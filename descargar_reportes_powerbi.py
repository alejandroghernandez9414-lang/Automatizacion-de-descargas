"""
descargar_reportes_powerbi.py

Automatiza la descarga de dos reportes desde un mismo dashboard de Power BI:
  1. Reporte MENSUAL (checkbox del mes actual en el slicer jerarquico)
  2. Reporte DIARIO  (checkbox del dia actual, dentro del mes, en el mismo slicer)

Diseñado con Playwright (sync API) para reutilizar sesion de Microsoft/Azure AD
mediante un perfil persistente, de forma que el MFA (codigo de 2 o 6 digitos)
solo se pida una vez cada ~24 horas, no en cada corrida.

REQUISITOS:
    pip install -r requirements.txt
    playwright install chromium
    Copiar .env.example como .env y completar los valores.

USO NORMAL (headless, programado por cron / Task Scheduler):
    python descargar_reportes_powerbi.py

CUANDO EXPIRA LA SESION (~cada 24h):
    El script detecta la pantalla de login/MFA, escribe un aviso claro en el
    log y termina SIN intentar automatizar el MFA. Ese dia debes correr:

        python descargar_reportes_powerbi.py --interactivo

    Esto abre el navegador CON ventana para que ingreses usuario, contraseña
    y el codigo de MFA una sola vez. La sesion queda guardada en el perfil
    persistente y las siguientes corridas headless vuelven a funcionar solas
    hasta que la sesion vuelva a expirar.
"""

import argparse
import logging
import os
import shutil
import sys
import re
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from config import opcional, requerido


class SesionExpiradaError(RuntimeError):
    """
    Se lanza cuando se detecta que terminamos en la pantalla de login de
    Microsoft (sesion vencida) en vez de en el reporte. Se distingue de un
    RuntimeError generico porque main() NO debe reintentar la corrida en
    este caso -- reintentar no arregla una sesion vencida, solo desperdicia
    minutos; lo que se necesita es correr --interactivo.
    """
    pass

# ---------------------------------------------------------------------------
# CONFIGURACION - ajusta estos valores a tu entorno
# ---------------------------------------------------------------------------

# URL completa del reporte embebido (incluye reportId y el id del tenant),
# por eso vive en .env y no en el codigo.
REPORT_URL = requerido("POWERBI_REPORT_URL")

# Nombre de la tabla del modelo de datos que se exporta (aparece en el
# atributo data-query-ref de las celdas del visual).
TABLA_EXPORTAR = requerido("POWERBI_TABLA_EXPORTAR")

# Ruta al ejecutable de Brave instalado en el equipo. Playwright no tiene
# un "channel" nativo para Brave (solo chrome/msedge), asi que se apunta
# directo al .exe ya instalado y confiado por el software de seguridad del equipo.
# Ajusta esta ruta si tu instalacion de Brave esta en otro lugar.
BRAVE_EXECUTABLE_PATH = opcional(
    "BRAVE_EXECUTABLE_PATH",
    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
)

# Nombre (o parte del nombre) del slicer jerarquico de fecha que se debe
# usar para filtrar el reporte. El dashboard tiene varios slicers en la
# misma pagina y todos comparten el mismo data-testid, asi que se
# identifica por su aria-label. IMPORTANTE: el aria-label real de estos
# slicers de fecha es "Año, Mes, Día" (el nombre tecnico del campo
# jerarquico), NO el titulo visible de la tarjeta ("Fecha Entrada").
# Ademas hay MAS DE UN slicer con este mismo aria-label (Fecha Entrada y
# Fecha Ejecucion comparten el mismo campo), asi que se distinguen por
# posicion: EL PRIMERO que aparece en el DOM corresponde a "Fecha
# Entrada" segun el layout visual (izquierda a derecha). Si esto
# resultara estar equivocado, cambiar SLICER_FECHA_INDICE a 1 para
# tomar el segundo en vez del primero.
NOMBRE_SLICER_FECHA = "Año, Mes, Día"
SLICER_FECHA_INDICE = 1  # 0 = primero (Fecha Entrada), 1 = segundo (Fecha Ejecucion)

# Carpeta donde Playwright guarda la sesion (cookies, localStorage, etc.)
# Usa una carpeta fija y persistente en disco, NO una carpeta temporal.
# Cada USUARIO de Microsoft/Azure AD necesita su PROPIA carpeta de perfil,
# ya que cada uno tiene su propio login/MFA. La carpeta base se sufija
# con --perfil (ver main()) para poder alternar entre varios usuarios sin
# que una sesion sobrescriba a la otra.
USER_DATA_DIR_BASE = Path.home() / ".pw_powerbi_profile"

# Se sobreescribe en main() segun el argumento --perfil. Por defecto usa
# el perfil base (retrocompatible si no se especifica --perfil).
USER_DATA_DIR = USER_DATA_DIR_BASE

# "brave" o "edge". Se sobreescribe en main() segun --navegador.
NAVEGADOR = "brave"

# Carpeta donde Playwright deja los archivos descargados ANTES de que
# save_as() los mueva a su destino final. Normalmente Playwright usa una
# carpeta temporal aleatoria que borra al cerrar; aqui se fija a una ruta
# conocida a proposito, para poder RESCATAR el archivo cuando el navegador
# se muere justo despues de terminar la descarga pero antes de que
# save_as() alcance a copiarlo (el error "Target page, context or browser
# has been closed" que venimos viendo). En ese escenario el .xlsx ya
# existe completo aqui -- lo unico que se perdio fue la conexion para
# moverlo, y eso se puede hacer con Python puro, sin navegador.
DOWNLOADS_TMP_DIR = Path.home() / ".pw_powerbi_downloads"


def kwargs_navegador() -> dict:
    """
    Arma los kwargs de launch_persistent_context segun NAVEGADOR.
    - brave: usa el ejecutable de Brave instalado (BRAVE_EXECUTABLE_PATH).
    - edge: usa channel="msedge", que le dice a Playwright que use el
      Microsoft Edge ya instalado en el sistema (no descarga uno propio).
      Se prueba como alternativa a Brave porque los crashes del navegador
      durante la descarga podrian ser especificos de Brave (su binario,
      o como el software de seguridad del equipo interactua especificamente con
      el, dado el --no-sandbox forzado que se detecto ahi).
    """
    if NAVEGADOR == "edge":
        return {"channel": "msedge"}
    return {"executable_path": BRAVE_EXECUTABLE_PATH}

# Carpeta de destino de los archivos exportados
# Usar ruta UNC (\\servidor\carpeta\...) en vez de una letra de unidad
# mapeada: las letras mapeadas pertenecen a la sesion interactiva del
# usuario y no siempre estan disponibles para una tarea programada.
OUTPUT_DIR = Path(requerido("POWERBI_OUTPUT_DIR"))

# Meses en español, en el mismo formato de texto que muestra el slicer
MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre",
    11: "noviembre", 12: "diciembre",
}

LOG_FILE = OUTPUT_DIR / "descarga_powerbi.log"

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

def setup_logging():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


# ---------------------------------------------------------------------------
# DETECCION DE LOGIN / MFA
# ---------------------------------------------------------------------------

def sesion_requiere_login(page) -> bool:
    """
    Revisa si la pagina actual es la pantalla de login de Microsoft
    (usuario/contraseña) o de MFA (numero de 2 digitos / codigo de 6 digitos).
    Ajusta los selectores si tu tenant usa un flujo distinto.

    IMPORTANTE: cuando la sesion esta vencida, Power BI NO redirige a
    login.microsoftonline.com de inmediato -- primero termina de cargar
    su propia pagina (dispara "domcontentloaded") y SOLO DESPUES, via JS,
    hace el redirect de SSO. Revisar page.url una sola vez justo despues
    de wait_for_load_state("domcontentloaded") puede pillar el momento
    "de en medio", antes de que el redirect ocurra, y concluir
    erroneamente que la sesion esta bien -- eso hace que el script se
    meta a buscar el reporte dentro de lo que en realidad es la pantalla
    de login, agotando reintentos sin sentido (visto en corrida real:
    9+ min de reintentos contra login.microsoftonline.com sin detectarlo).
    Por eso aqui se reintenta la revision varias veces con una pequeña
    espera entre cada una, dandole tiempo al redirect a completarse.
    """
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except PWTimeout:
        pass

    dominios_login = ("login.microsoftonline.com", "login.live.com", "adfs")
    indicadores_login = [
        "input[type='email']",
        "input[name='loginfmt']",
        "input[type='password']",
        "text=Enter code",
        "text=Ingresa el código",
        "text=Approve sign in request",
        "text=Aprobar solicitud de inicio de sesión",
    ]

    intentos_deteccion = 5
    for intento in range(intentos_deteccion):
        # Cualquier redireccion a dominios de auth de Microsoft es señal fuerte
        if any(d in page.url for d in dominios_login):
            return True

        for selector in indicadores_login:
            try:
                if page.locator(selector).first.is_visible(timeout=1000):
                    return True
            except Exception:
                continue

        if intento < intentos_deteccion - 1:
            # Todavia no hay señal de login, pero puede que el redirect de
            # SSO aun no haya disparado -- dar tiempo y volver a revisar.
            page.wait_for_timeout(1500)

    return False


def modo_interactivo_login(playwright):
    """
    Abre un navegador CON ventana para que el usuario complete manualmente
    usuario, contraseña y MFA. Al cerrar, la sesion queda guardada en
    USER_DATA_DIR para las siguientes corridas headless.
    """
    logging.info(f"Abriendo navegador visible ({NAVEGADOR}) para login manual (MFA)...")
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=str(USER_DATA_DIR),
        headless=False,
        accept_downloads=True,
        downloads_path=str(DOWNLOADS_TMP_DIR),  # ver comentario en la constante
        **kwargs_navegador(),  # executable_path=Brave o channel="msedge"
        # Ver comentario detallado en el launch_persistent_context del modo
        # headless (mas abajo en el archivo) sobre por que se agrega este
        # flag: el software de seguridad del equipo fuerza --no-sandbox en
        # cualquier Chromium que se abra aqui (confirmado viendo el banner
        # de advertencia en una ventana --interactivo real), lo que deja
        # expuesto el proceso a que Windows lo mate por "Renderer Code
        # Integrity" cuando el software de seguridad inyecta su DLL de monitoreo.
        # Este flag y los de restauracion de sesion de abajo son flags
        # de Chromium estandar, funcionan igual en Brave y en Edge.
        #
        # Los flags de restauracion de sesion: evitan que el navegador
        # ofrezca o aplique "restaurar pestañas anteriores" tras un
        # cierre anormal -- confirmado que eso puede dejar varias
        # pestañas del reporte pesado cargando a la vez en la siguiente
        # corrida.
        args=[
            "--disable-features=RendererCodeIntegrity",
            "--disable-session-crashed-bubble",
            "--hide-crash-restore-bubble",
            "--no-restore-session-state",
        ],
    )

    # Cerrar cualquier pestaña sobrante que haya quedado restaurada de
    # una corrida anterior, para empezar siempre con una sola.
    while len(context.pages) > 1:
        context.pages[-1].close()

    page = context.pages[0] if context.pages else context.new_page()
    page.goto(REPORT_URL, wait_until="domcontentloaded")

    print("\n" + "=" * 70)
    print("Completa el login (usuario, contraseña y MFA) en la ventana abierta.")
    print("Cuando el reporte cargue correctamente, vuelve aqui y presiona ENTER.")
    print("=" * 70 + "\n")
    input("Presiona ENTER cuando el reporte haya cargado... ")

    context.close()
    logging.info("Sesion guardada. Ya puedes correr el script en modo headless normal.")


# ---------------------------------------------------------------------------
# NAVEGACION DENTRO DEL REPORTE (maneja iframe si aplica)
# ---------------------------------------------------------------------------

def obtener_frame_reporte(page):
    """
    Power BI a veces renderiza el reporte dentro de un iframe interno,
    y puede reemplazar ese iframe una vez mas poco despues de la carga
    inicial (lo que puede causar "Frame was detached" si se captura la
    referencia demasiado pronto). Se espera un poco mas para reducir esa
    posibilidad, aunque esperar_reporte_cargado tiene reintento por si
    aun asi ocurre.

    Puede haber MAS DE UN frame con 'powerbi' en la URL (ej. uno de
    telemetria/autenticacion vacio ademas del frame real del reporte),
    asi que se devuelve una LISTA de candidatos en vez de asumir que el
    primero encontrado es el correcto -- esperar_reporte_cargado prueba
    cada uno hasta encontrar el que realmente tiene los visuales.
    """
    page.wait_for_timeout(5000)
    logging.info(
        f"Frames totales en la pagina: "
        f"{[(f.url[:80], f == page.main_frame) for f in page.frames]}"
    )
    candidatos = [
        frame for frame in page.frames
        if "powerbi" in frame.url.lower() and frame != page.main_frame
    ]
    if not candidatos:
        return [page]
    # Se agrega tambien page.main_frame como candidato de respaldo, por si
    # el reporte esta directamente en la pagina principal (sin iframe) en
    # algunas cargas, o por si los frames con 'powerbi' en la URL no son
    # realmente donde vive el contenido visible.
    candidatos.append(page)
    return candidatos


def esperar_reporte_cargado(page, destino, reintentos=2):
    """Espera a que el canvas del reporte este visible, y luego da tiempo
    extra para que el resto de visuales (slicers incluidos) terminen de
    renderizar, ya que la pagina puede ir lenta.

    Si el frame se desconecta ("Frame was detached") mientras se espera
    -- Power BI a veces reemplaza el iframe interno tras la carga inicial
    -- se vuelve a resolver el frame y se reintenta, hasta `reintentos`
    veces, antes de darse por vencido.

    Devuelve el frame/page vigente (puede ser distinto de `destino` si
    hubo que reemplazarlo).

    Si falla definitivamente, guarda screenshot + HTML para diagnostico.
    """
    intento = 0
    # Lista de selectores candidatos para detectar que el reporte esta
    # renderizado. No todos los reportes de Power BI generan los mismos
    # data-testid/clases (depende de la version del motor de render), asi
    # que se prueban varios en orden y basta con que UNO aparezca.
    # slicer-dropdown se confirmo visualmente que existe en este reporte.
    SELECTORES_REPORTE_CARGADO = (
        "[data-testid='slicer-dropdown'], "
        "[data-testid='visual-container'], "
        ".visualContainerHost, "
        ".visual-container, "
        ".card.visual, "
        "visual-modern"
    )

    # destino puede venir como una LISTA de frames candidatos (de
    # obtener_frame_reporte) o como un solo frame/page (si ya se resolvio
    # en un reintento anterior). Se normaliza a lista para probar cada uno.
    candidatos = destino if isinstance(destino, list) else [destino]

    while True:
        for candidato in candidatos:
            try:
                # Se usa state="attached" (existe en el DOM) en vez de
                # "visible" (por defecto). En corridas reales el reporte
                # visualmente cargo completo (confirmado con screenshots)
                # pero wait_for_selector con "visible" seguia fallando --
                # posible falso negativo de Playwright por CSS/overlay/
                # timing de renderizado. "attached" es una condicion mas
                # laxa y suficiente: si el elemento existe, el reporte ya
                # esta ahi aunque Playwright dude de su "visibilidad".
                candidato.wait_for_selector(
                    SELECTORES_REPORTE_CARGADO, timeout=30000, state="attached"
                )
                # Espera adicional para que terminen de cargar visuales mas
                # lentos (como el slicer jerarquico) despues de que aparece
                # el primer visual.
                candidato.wait_for_timeout(5000)
                return candidato
            except Exception as e:
                if "detached" in str(e).lower():
                    continue  # probar el siguiente candidato o reintentar abajo
                continue  # este candidato no tenia los visuales; probar el siguiente
        # Ninguno de los candidatos actuales funciono.
        if intento < reintentos:
            intento += 1
            logging.warning(
                f"Ningun frame candidato mostro el reporte cargado "
                f"(reintento {intento}/{reintentos}); volviendo a resolver frames..."
            )
            page.wait_for_timeout(2000)
            candidatos = obtener_frame_reporte(page)
            continue

        # Antes de rendirse: si en algun momento la pagina termino en el
        # dominio de login de Microsoft, esto NUNCA fue un problema del
        # reporte/frames -- es sesion vencida que sesion_requiere_login()
        # no alcanzo a detectar a tiempo (redirect de SSO tardio). No tiene
        # sentido reintentar 4 veces contra una pantalla de login.
        if any(d in page.url for d in ("login.microsoftonline.com", "login.live.com", "adfs")):
            logging.warning(
                f"El reporte nunca cargo porque la sesion esta vencida "
                f"(URL termino en pantalla de login: {page.url[:120]}...). "
                f"Corre 'python descargar_reportes_powerbi.py --interactivo' "
                f"para renovar la sesion."
            )
            raise SesionExpiradaError(
                "Sesion vencida: se detecto pantalla de login de Microsoft "
                "en vez del reporte."
            )

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            page.screenshot(path=str(OUTPUT_DIR / f"debug_{ts}.png"), full_page=True)
        except Exception:
            logging.exception("No se pudo tomar screenshot de diagnostico.")
        try:
            (OUTPUT_DIR / f"debug_{ts}.html").write_text(page.content(), encoding="utf-8")
        except Exception:
            logging.exception("No se pudo guardar HTML de diagnostico.")
        logging.error(
            f"Reporte no cargo tras probar {len(candidatos)} frame(s) candidato(s). "
            f"URL actual: {page.url}. "
            f"Diagnostico guardado en debug_{ts}.png / .html"
        )
        raise RuntimeError(
            "El reporte de Power BI no termino de cargar en ningun frame candidato."
        )


# ---------------------------------------------------------------------------
# MANEJO DEL SLICER JERARQUICO (Año > Mes > Dia)
# ---------------------------------------------------------------------------

def abrir_slicer(page, destino):
    """
    Abre el panel del slicer jerarquico (Año/Mes/Dia) si esta colapsado,
    y devuelve el LOCATOR DEL POPUP donde vive el arbol de fechas.

    IMPORTANTE: Power BI monta el popup del slicer (el arbol Año/Mes/Dia)
    en un contenedor SEPARADO del boton dropdown, referenciado por el
    atributo aria-controls del boton (ej. "slicer-dropdown-popup-<guid>").
    Ese popup puede estar fuera del iframe del reporte (montado directo
    en la pagina), asi que buscamos primero en destino y si no aparece,
    en page.

    IMPORTANTE 2: el dashboard tiene varios slicers en la misma pagina
    que comparten el mismo data-testid='slicer-dropdown' (Descripcion
    Mercado, Pto Responsable, Municipio, Fecha Entrada, Fecha Ejecucion,
    etc.). Se filtra por aria-label para coger el correcto
    (NOMBRE_SLICER_FECHA), NO el primero que aparezca en el DOM.

    Selector confirmado: div.slicer-dropdown-menu[data-testid='slicer-dropdown']
    con aria-expanded indicando si el popup ya esta abierto.
    Si no lo encuentra, guarda screenshot + HTML para diagnostico.
    """
    dropdown = destino.locator(
        f"[data-testid='slicer-dropdown'][aria-label*='{NOMBRE_SLICER_FECHA}']"
    ).nth(SLICER_FECHA_INDICE)
    try:
        dropdown.wait_for(state="visible", timeout=60000)
    except PWTimeout:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            page.screenshot(path=str(OUTPUT_DIR / f"debug_slicer_{ts}.png"), full_page=True)
        except Exception:
            logging.exception("No se pudo tomar screenshot de diagnostico del slicer.")
        try:
            labels_disponibles = destino.locator(
                "[data-testid='slicer-dropdown']"
            ).evaluate_all("els => els.map(e => e.getAttribute('aria-label'))")
            logging.error(f"aria-label de slicers encontrados en la pagina: {labels_disponibles}")
        except Exception:
            logging.exception("No se pudieron listar los aria-label disponibles.")
        try:
            (OUTPUT_DIR / f"debug_slicer_{ts}.html").write_text(destino.content() if hasattr(destino, "content") else page.content(), encoding="utf-8")
        except Exception:
            logging.exception("No se pudo guardar HTML de diagnostico del slicer.")
        logging.error(
            f"No se encontro el slicer con aria-label que contenga "
            f"'{NOMBRE_SLICER_FECHA}'. "
            f"Diagnostico guardado en debug_slicer_{ts}.png / .html"
        )
        raise

    popup_id = dropdown.get_attribute("aria-controls")

    expandido = dropdown.get_attribute("aria-expanded")
    if expandido != "true":
        dropdown.click(timeout=5000)
        destino.wait_for_timeout(500)

    # Localizar el contenedor real del popup por su id, probando primero
    # dentro del frame del reporte y luego en la pagina completa.
    popup = None
    if popup_id:
        candidato_frame = destino.locator(f"#{popup_id}")
        candidato_page = page.locator(f"#{popup_id}")
        try:
            if candidato_frame.count() > 0:
                popup = candidato_frame
                logging.info(f"Popup del slicer encontrado dentro del frame del reporte (#{popup_id}).")
            elif candidato_page.count() > 0:
                popup = candidato_page
                logging.info(f"Popup del slicer encontrado en la pagina, fuera del frame (#{popup_id}).")
        except Exception:
            logging.exception("Error localizando el popup del slicer por id.")

    if popup is None:
        # Fallback: usar el mismo destino de siempre (comportamiento previo)
        logging.warning(
            f"No se pudo ubicar el popup del slicer por aria-controls "
            f"('{popup_id}'); usando el frame del reporte como fallback."
        )
        popup = destino

    return dropdown, popup


def cerrar_slicer(page, dropdown):
    """
    Cierra el popup del slicer haciendo clic de nuevo en el boton dropdown,
    si esta expandido. Necesario porque el popup abierto puede quedar
    superpuesto sobre otros visuales (ej. el boton 'Mas opciones' de la
    tabla a exportar) y bloquear los clics posteriores.
    """
    try:
        if dropdown.get_attribute("aria-expanded") == "true":
            dropdown.click(timeout=5000)
            page.wait_for_timeout(500)
    except Exception:
        logging.exception("No se pudo cerrar el popup del slicer (se continua de todas formas).")


def limpiar_seleccion_slicer(destino):
    """
    Deja el slicer sin selecciones para partir limpio antes de marcar
    la opcion que corresponda (mes o dia).
    """
    boton_limpiar = destino.get_by_role("button", name=re.compile("Clear|Borrar|Limpiar", re.IGNORECASE))
    try:
        boton_limpiar.first.click(timeout=3000)
    except Exception:
        logging.info("No se encontro boton de limpiar selección; continuo sin limpiar.")


def expandir_anio(page, destino, anio: int, timeout_espera=30000):
    """
    Expande el nodo del año en el arbol jerarquico del slicer (nivel 1)
    para que sus hijos (los meses) aparezcan en el DOM. El arbol arranca
    colapsado mostrando solo 'Seleccionar todo' y el año (ej. '2026'),
    y hasta no expandirlo los meses no existen como treeitem.

    Si el popup se reabrio despues de un scroll previo (ej. tras
    seleccionar el mes anterior con buscar_con_scroll), el nodo del año
    puede haber quedado fuera del viewport visible (scrolleado hacia
    abajo). Primero se intenta hacer scroll hacia ARRIBA para volver al
    tope de la lista, y si aun asi no aparece, se usa buscar_con_scroll
    normal (que scrollea hacia abajo) como respaldo.
    """
    # ANTES de scrollear: esperar a que la lista realmente tenga contenido.
    # El arbol del slicer es virtualizado y se puebla por JS despues de que
    # el popup aparece. Si se empieza a scrollear cuando todavia esta vacia
    # (o apenas con los primeros años), la busqueda falla aunque el año SI
    # exista -- se vio en corridas reales fallando con la lista en [] y con
    # solo ['Seleccionar todo', '(En blanco)', '2004'...'2009'] cargados.
    for _ in range(20):
        try:
            n = destino.locator("[role='treeitem']").count()
        except Exception:
            n = 0
        if n > 3:  # mas que 'Seleccionar todo' + '(En blanco)' + 1
            break
        page.wait_for_timeout(500)
    else:
        logging.warning(
            "La lista del slicer sigue casi vacia tras esperar 10s; "
            "se intenta buscar el año igual, pero es probable que falle."
        )

    # Intentar volver al tope del popup por si quedo scrolleado abajo.
    try:
        caja = destino.bounding_box()
        if caja:
            cx = caja["x"] + caja["width"] / 2
            cy = caja["y"] + caja["height"] / 2
            page.mouse.move(cx, cy)
            page.mouse.wheel(0, -10000)  # scroll grande hacia arriba
            page.wait_for_timeout(300)
    except Exception:
        logging.exception("Error haciendo scroll hacia arriba en el popup del slicer.")

    # El arbol de años sigue creciendo (2021...2027 y seguira sumando),
    # asi que se dan mas intentos de scroll de los que bastaban antes.
    fila_anio = buscar_con_scroll(
        page, destino, f"[role='treeitem'][title='{anio}']", max_intentos=35
    )
    if fila_anio is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            page.screenshot(path=str(OUTPUT_DIR / f"debug_anio_{ts}.png"), full_page=True)
        except Exception:
            logging.exception("No se pudo tomar screenshot de diagnostico del año.")
        try:
            titles_disponibles = destino.locator("[role='treeitem']").evaluate_all(
                "els => els.map(e => e.getAttribute('title'))"
            )
            logging.error(f"Titles de treeitem encontrados al buscar el año: {titles_disponibles}")
        except Exception:
            logging.exception("No se pudieron listar los titles disponibles.")
        logging.error(
            f"No se encontro treeitem con title='{anio}'. "
            f"Diagnostico guardado en debug_anio_{ts}.png"
        )
        raise RuntimeError(f"No se encontro el año '{anio}' en el slicer tras hacer scroll.")
    if fila_anio.get_attribute("aria-expanded") != "true":
        fila_anio.locator(".expandButton").click(timeout=5000)
        page.wait_for_timeout(800)


def marcar_checkbox_si_no_esta(fila, activar: bool = True, timeout=10000):
    """
    Marca (o desmarca) el checkbox de una fila del slicer SOLO SI no esta
    ya en el estado deseado. Hacer clic sobre un checkbox ya marcado lo
    DESMARCA (toggle), asi que clickear a ciegas sin revisar el estado
    previo puede invertir una seleccion existente -- esto paso realmente:
    al seleccionar el mes completo primero y despues intentar marcar un
    dia especifico, el dia ya estaba marcado (por ser parte del mes) y el
    clic lo desmarcaba en vez de dejarlo como unica seleccion.
    """
    estado_actual = fila.get_attribute("aria-selected")
    ya_esta_activo = estado_actual == "true"
    if ya_esta_activo != activar:
        fila.locator(".slicerCheckbox").click(timeout=timeout)


def buscar_con_scroll(page, destino, selector: str, max_intentos=25, paso_scroll=100):
    """
    Busca un elemento dentro de una lista virtualizada del slicer (Power BI
    solo renderiza en el DOM los nodos visibles en el area de scroll, asi
    que un mes o dia especifico puede no existir aun si esta fuera del
    area de scroll actual). Hace scroll progresivo dentro del popup y
    reintenta la busqueda hasta encontrarlo o agotar los intentos.

    Si tras hacer scroll hacia ABAJO no aparece (posible overshoot: el
    paso de scroll salto por encima del elemento sin que el DOM llegara
    a listarlo), se intenta un segundo barrido hacia ARRIBA en pasos
    finos para "volver" y encontrarlo.

    Devuelve el locator si lo encuentra (ya visible), o None si no
    aparecio tras ambos barridos.
    """
    def _intentar(direccion, cantidad):
        for _ in range(cantidad):
            elemento = destino.locator(selector).first
            if elemento.count() > 0:
                try:
                    elemento.scroll_into_view_if_needed(timeout=3000)
                    elemento.wait_for(state="visible", timeout=3000)
                    return elemento
                except Exception:
                    pass
            try:
                caja = destino.bounding_box()
                if caja:
                    cx = caja["x"] + caja["width"] / 2
                    cy = caja["y"] + caja["height"] / 2
                    page.mouse.move(cx, cy)
                    page.mouse.wheel(0, direccion * paso_scroll)
            except Exception:
                logging.exception("Error haciendo scroll dentro del popup del slicer.")
            page.wait_for_timeout(350)
        return None

    encontrado = _intentar(1, max_intentos)
    if encontrado is not None:
        return encontrado

    # Posible overshoot: reintentar con barrido hacia arriba en pasos finos.
    logging.info(
        f"No se encontro '{selector}' bajando; reintentando subiendo "
        f"(posible overshoot del scroll)."
    )
    return _intentar(-1, max_intentos)


def seleccionar_mes(page, destino, anio: int, mes: int):
    """
    Expande el año y marca el checkbox del mes actual (vista mensual).
    La lista de meses esta virtualizada (Power BI solo renderiza los
    nodos visibles), asi que se hace scroll progresivo hasta encontrar
    el mes buscado.
    Selector confirmado: div.slicerItemContainer[role='treeitem'][title='agosto']
    Si no lo encuentra, guarda diagnostico con los titles reales disponibles.
    """
    nombre_mes = MESES_ES[mes]
    logging.info(f"Seleccionando mes: {nombre_mes} {anio}")
    limpiar_seleccion_slicer(destino)
    expandir_anio(page, destino, anio)

    fila_mes = buscar_con_scroll(
        page, destino, f"[role='treeitem'][title='{nombre_mes}']"
    )
    if fila_mes is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            page.screenshot(path=str(OUTPUT_DIR / f"debug_mes_{ts}.png"), full_page=True)
        except Exception:
            logging.exception("No se pudo tomar screenshot de diagnostico del mes.")
        try:
            titles_disponibles = destino.locator("[role='treeitem']").evaluate_all(
                "els => els.map(e => e.getAttribute('title'))"
            )
            logging.error(f"Titles de treeitem encontrados en el slicer: {titles_disponibles}")
        except Exception:
            logging.exception("No se pudieron listar los titles disponibles.")
        try:
            (OUTPUT_DIR / f"debug_mes_{ts}.html").write_text(destino.content() if hasattr(destino, "content") else page.content(), encoding="utf-8")
        except Exception:
            logging.exception("No se pudo guardar HTML de diagnostico del mes.")
        logging.error(
            f"No se encontro treeitem con title='{nombre_mes}' tras hacer scroll. "
            f"Diagnostico guardado en debug_mes_{ts}.png / .html"
        )
        raise RuntimeError(f"No se encontro el mes '{nombre_mes}' en el slicer tras hacer scroll.")
    marcar_checkbox_si_no_esta(fila_mes, activar=True)


def ubicar_dia_en_mes_expandido(page, destino, fila_mes, dia: int, max_intentos=30, paso_scroll=100):
    """
    Busca el treeitem del dia dentro de un mes YA EXPANDIDO, re-anclando
    la vista al propio mes en cada intento (en vez de scrollear a ciegas
    sobre todo el popup), para evitar que el scroll se pase de largo y
    termine colapsando de vuelta al nivel de año/mes.
    """
    for intento in range(max_intentos):
        # Re-anclar al mes en cada intento: si el DOM cambio o el scroll
        # se desvio, esto nos vuelve a poner cerca de donde estan los dias.
        try:
            fila_mes.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass

        fila_dia = destino.locator(
            f"[role='treeitem'][aria-level='3'][title='{dia}']"
        ).first
        if fila_dia.count() > 0:
            try:
                fila_dia.scroll_into_view_if_needed(timeout=3000)
                fila_dia.wait_for(state="visible", timeout=3000)
                return fila_dia
            except Exception:
                pass

        # Scroll pequeño hacia abajo desde la posicion actual (ya anclada
        # cerca del mes), para ir revelando los dias progresivamente.
        try:
            caja = destino.bounding_box()
            if caja:
                cx = caja["x"] + caja["width"] / 2
                cy = caja["y"] + caja["height"] / 2
                page.mouse.move(cx, cy)
                page.mouse.wheel(0, paso_scroll)
        except Exception:
            logging.exception("Error haciendo scroll dentro del popup del slicer (dia).")
        page.wait_for_timeout(350)
    return None


def seleccionar_dia(page, destino, anio: int, mes: int, dia: int):
    """
    Expande el año y el mes actual, y marca el checkbox del dia actual
    (vista diaria).
    Selectores confirmados:
      - Fila:   div.slicerItemContainer[role='treeitem'][title='<nombre>']
      - Expandir: .expandButton dentro de la fila, solo si aria-expanded='false'
      - Dia:    misma estructura, title='<numero de dia>', aria-level='3'
    """
    nombre_mes = MESES_ES[mes]
    logging.info(f"Seleccionando dia: {dia} de {nombre_mes} {anio}")
    limpiar_seleccion_slicer(destino)
    expandir_anio(page, destino, anio)

    fila_mes = buscar_con_scroll(
        page, destino, f"[role='treeitem'][title='{nombre_mes}']"
    )
    if fila_mes is None:
        raise RuntimeError(f"No se encontro el mes '{nombre_mes}' en el slicer tras hacer scroll.")

    if fila_mes.get_attribute("aria-expanded") != "true":
        fila_mes.locator(".expandButton").click(timeout=5000)
        page.wait_for_timeout(500)
        try:
            fila_mes.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            logging.exception("No se pudo hacer scroll hacia el mes recien expandido.")

    # IMPORTANTE: limpiar_seleccion_slicer no siempre encuentra el boton
    # de limpiar (se confirmo en logs reales), asi que la seleccion del
    # mes completo de la corrida anterior (reporte mensual) puede seguir
    # activa aqui. Si no se desmarca el mes ahora, el dia buscado ya
    # apareceria marcado (por ser parte del mes) y el reporte diario
    # terminaria exportando TODO el mes menos el dia elegido, en vez de
    # solo ese dia. Se desmarca el mes explicitamente antes de buscar
    # el dia especifico.
    marcar_checkbox_si_no_esta(fila_mes, activar=False)

    # Se re-ancla al mes (ya expandido) en cada intento de scroll, en vez
    # de confiar en scroll libre sobre todo el popup, para evitar el
    # overshoot que colapsaba de vuelta al nivel de año.
    fila_dia = ubicar_dia_en_mes_expandido(page, destino, fila_mes, dia)
    if fila_dia is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            page.screenshot(path=str(OUTPUT_DIR / f"debug_dia_{ts}.png"), full_page=True)
        except Exception:
            logging.exception("No se pudo tomar screenshot de diagnostico del dia.")
        try:
            items_disponibles = destino.locator("[role='treeitem']").evaluate_all(
                "els => els.map(e => ({title: e.getAttribute('title'), level: e.getAttribute('aria-level')}))"
            )
            logging.error(f"treeitems encontrados en el slicer al buscar el dia: {items_disponibles}")
        except Exception:
            logging.exception("No se pudieron listar los treeitems disponibles.")
        try:
            (OUTPUT_DIR / f"debug_dia_{ts}.html").write_text(destino.content() if hasattr(destino, "content") else page.content(), encoding="utf-8")
        except Exception:
            logging.exception("No se pudo guardar HTML de diagnostico del dia.")
        logging.error(
            f"No se encontro treeitem aria-level=3 title='{dia}'. "
            f"Diagnostico guardado en debug_dia_{ts}.png / .html"
        )
        raise RuntimeError(f"No se encontro el dia '{dia}' en el slicer tras hacer scroll.")
    marcar_checkbox_si_no_esta(fila_dia, activar=True)


# ---------------------------------------------------------------------------
# EXPORTACION
# ---------------------------------------------------------------------------

def exportar_datos_actuales(page, destino, nombre_archivo: str, carpeta_destino: Path = None):
    """
    Ejecuta el flujo completo de exportacion sobre la tabla configurada en TABLA_EXPORTAR:
      1. Ubica el visual por su data-query-ref (identifica la tabla correcta
         entre varios visuales de la pagina).
      2. Hover sobre el visual para que aparezca el boton "Mas opciones" (...).
      3. Clic en ese boton -> aparece menu contextual -> clic en "Exportar datos".
      4. La ventana emergente ya trae "Datos con diseño actual" seleccionado
         por defecto -> clic directo en el boton "Exportar" (data-testid='export-btn').
      5. Captura la descarga y la guarda con el nombre indicado.

    carpeta_destino: carpeta donde se guarda el archivo. Si no se indica,
    usa OUTPUT_DIR (la misma carpeta para diario y mensual).
    """
    if carpeta_destino is None:
        carpeta_destino = OUTPUT_DIR
    logging.info(f"Exportando -> {nombre_archivo}")

    # 1. Ubicar el contenedor del visual correcto usando el nombre de tabla
    #    del modelo de datos, confirmado en el data-query-ref de sus columnas.
    celda_tabla = destino.locator(f"[data-query-ref*='{TABLA_EXPORTAR}']").first
    celda_tabla.wait_for(state="visible", timeout=40000)
    contenedor_visual = celda_tabla.locator(
        "xpath=ancestor::*[@data-testid='visual-container'][1]"
    )

    # 2. Hover para que aparezca el boton de "Mas opciones" (esta oculto
    #    por CSS hasta que el mouse esta sobre el visual).
    contenedor_visual.hover(timeout=10000)
    destino.wait_for_timeout(300)

    # 3. Clic en "Mas opciones" (...) -> clic en "Exportar datos"
    boton_mas_opciones = contenedor_visual.locator(
        "[data-testid='visual-more-options-btn']"
    ).first
    boton_mas_opciones.click(timeout=10000)

    opcion_exportar = page.get_by_text("Exportar datos", exact=False).first
    opcion_exportar.click(timeout=10000)

    # 4. La ventana emergente ya trae "Datos con diseño actual" seleccionado
    #    por defecto, asi que vamos directo al boton de confirmar.
    boton_confirmar = page.locator("[data-testid='export-btn']").first
    boton_confirmar.wait_for(state="visible", timeout=30000)

    # Se anota que archivos ya existian en la carpeta temporal de descargas
    # ANTES de disparar esta descarga, para poder identificar despues cual
    # es el archivo nuevo si hay que rescatarlo a mano (ver mas abajo).
    DOWNLOADS_TMP_DIR.mkdir(parents=True, exist_ok=True)
    archivos_previos = set(DOWNLOADS_TMP_DIR.rglob("*"))

    with page.expect_download(timeout=60000) as info_descarga:
        boton_confirmar.click()
    descarga = info_descarga.value

    carpeta_destino.mkdir(parents=True, exist_ok=True)
    ruta_destino = carpeta_destino / nombre_archivo

    # Si el archivo de destino esta abierto en Excel (u otro programa que
    # lo bloquee), Windows impide sobrescribirlo. Se detecta esto ANTES
    # de intentar guardar, para dar un mensaje claro en vez de un error
    # generico de Playwright/permisos.
    if ruta_destino.exists():
        try:
            # Intentar abrir el archivo en modo append binario es la forma
            # mas confiable en Windows de detectar si otro proceso tiene
            # un lock exclusivo sobre el (Excel abierto lo bloquea asi).
            with open(ruta_destino, "ab"):
                pass
        except PermissionError:
            logging.error(
                f"No se pudo guardar '{ruta_destino.name}': el archivo esta "
                f"abierto en Excel (u otro programa). Cierralo y vuelve a "
                f"correr el script."
            )
            raise RuntimeError(
                f"El archivo '{ruta_destino.name}' esta abierto en otro "
                f"programa (probablemente Excel). Cierralo y reintenta."
            )

    try:
        descarga.save_as(str(ruta_destino))
    except Exception as e:
        msg = str(e).lower()

        if "permission" in msg or "denied" in msg:
            logging.error(
                f"No se pudo guardar '{ruta_destino.name}': posible archivo "
                f"abierto en Excel u otro programa bloqueandolo."
            )
            raise

        # RESCATE: el navegador murio durante/despues de la descarga
        # ("Target page, context or browser has been closed"). Playwright
        # ya habia bajado el archivo COMPLETO a DOWNLOADS_TMP_DIR; lo unico
        # que se perdio fue la conexion para moverlo. Asi que se busca el
        # archivo nuevo ahi y se copia con Python puro, sin navegador.
        if "closed" in msg or "target" in msg:
            logging.warning(
                f"El navegador se cerro durante save_as. Intentando rescatar "
                f"'{nombre_archivo}' desde la carpeta temporal de descargas..."
            )
            try:
                candidatos = [
                    p for p in DOWNLOADS_TMP_DIR.rglob("*")
                    if p.is_file() and p not in archivos_previos
                ]
                if candidatos:
                    # El mas reciente por fecha de modificacion es el de
                    # esta descarga.
                    origen = max(candidatos, key=lambda p: p.stat().st_mtime)
                    tam = origen.stat().st_size
                    # Sanidad: un .xlsx valido nunca es de unos pocos bytes.
                    # Si esta vacio o truncado, la descarga alcanzo a morir
                    # antes de terminar y NO sirve -- mejor fallar y que el
                    # bucle de reintentos haga otra corrida limpia.
                    if tam > 5000:
                        shutil.copy2(str(origen), str(ruta_destino))
                        logging.info(
                            f"RESCATADO desde temporal ({tam:,} bytes). "
                            f"Guardado en: {ruta_destino}"
                        )
                        try:
                            origen.unlink()  # limpiar el temporal rescatado
                        except Exception:
                            pass
                        return
                    logging.warning(
                        f"El archivo temporal esta truncado ({tam:,} bytes), "
                        f"la descarga no alcanzo a completarse. Se reintentara."
                    )
                else:
                    logging.warning(
                        "No se encontro ningun archivo nuevo en la carpeta "
                        "temporal; la descarga no alcanzo a escribirse."
                    )
            except Exception:
                logging.exception("Fallo el intento de rescate desde el temporal.")

        raise
    logging.info(f"Guardado en: {ruta_destino}")


# ---------------------------------------------------------------------------
# FLUJO PRINCIPAL
# ---------------------------------------------------------------------------

def correr_descargas():
    hoy = datetime.now()
    fecha_str = hoy.strftime("%Y%m%d")

    with sync_playwright() as playwright:
        logging.info(f"Abriendo navegador headless ({NAVEGADOR})...")
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(USER_DATA_DIR),
            headless=True,
            accept_downloads=True,
            downloads_path=str(DOWNLOADS_TMP_DIR),  # ver comentario en la constante
            **kwargs_navegador(),  # executable_path=Brave o channel="msedge"
            # --headless=old: ver comentario de arriba (modo de descargas
            # mas maduro que el headless "nuevo" por defecto). Nota: este
            # flag es especifico de Chromium/Brave -- si NAVEGADOR="edge",
            # Edge lo ignora sin problema si no aplica.
            #
            # --disable-features=RendererCodeIntegrity: se descubrio que
            # el software de seguridad del equipo esta forzando
            # --no-sandbox en CUALQUIER navegador Chromium que se abre
            # ahi (se ve el banner de advertencia hasta en las ventanas
            # que abre este mismo script). Con --no-sandbox activo,
            # Windows sigue bloqueando DLLs no firmadas por Microsoft
            # dentro del proceso del navegador (proteccion "Renderer
            # Code Integrity") -- y el DLL que el software de seguridad inyecta para
            # vigilar NO esta firmado por Microsoft, asi que Windows lo
            # trata como violacion y mata el proceso. Esto coincide con
            # el patron visto: se cae ~20s despues de iniciar una
            # descarga pesada, justo cuando hay mas actividad para que
            # --disable-session-crashed-bubble / --hide-crash-restore-bubble:
            # se detecto que despues de un cierre forzado del navegador
            # (justo lo que hemos estado viendo), Brave reabre ofreciendo
            # "restaurar pestañas anteriores" -- si eso llega a aceptarse
            # (o si Brave lo hace solo por --restore-last-session
            # implicito), la siguiente corrida puede terminar con VARIAS
            # copias del reporte pesado cargando al mismo tiempo en
            # pestañas distintas, compitiendo por memoria/CPU. Esto se
            # confirmo viendo 3 pestañas de Power BI abiertas a la vez
            # tras varias corridas fallidas seguidas. Estos flags evitan
            # que se ofrezca/aplique esa restauracion.
            args=[
                "--headless=old",
                "--disable-features=RendererCodeIntegrity",
                "--disable-session-crashed-bubble",
                "--hide-crash-restore-bubble",
                "--no-restore-session-state",
            ],
        )

        # Defensa adicional, sin importar si los flags de arriba
        # funcionaron o no: cerrar cualquier pestaña sobrante que haya
        # quedado abierta (restaurada de una corrida anterior) antes de
        # seguir, para garantizar que solo haya UNA pestaña cargando el
        # reporte a la vez.
        while len(context.pages) > 1:
            context.pages[-1].close()

        page = context.pages[0] if context.pages else context.new_page()

        logging.info("Navegando al reporte de Power BI...")
        page.goto(REPORT_URL, wait_until="domcontentloaded")

        if sesion_requiere_login(page):
            logging.warning(
                "La sesion expiro y se requiere login/MFA manual. "
                "Corre este script con --interactivo para completarlo. "
                "No se realizara ninguna descarga en esta corrida."
            )
            context.close()
            sys.exit(2)

        destino = obtener_frame_reporte(page)
        destino = esperar_reporte_cargado(page, destino)
        dropdown_slicer, popup_slicer = abrir_slicer(page, destino)

        # --- Descarga MENSUAL ---
        seleccionar_mes(page, popup_slicer, hoy.year, hoy.month)
        cerrar_slicer(page, dropdown_slicer)
        exportar_datos_actuales(page, destino, f"reporte_mensual_{fecha_str}.xlsx")

        # --- Descarga DIARIA ---
        abrir_slicer(page, destino)  # reabrir para cambiar la seleccion
        seleccionar_dia(page, popup_slicer, hoy.year, hoy.month, hoy.day)
        cerrar_slicer(page, dropdown_slicer)
        exportar_datos_actuales(page, destino, f"reporte_diario_{fecha_str}.xlsx")

        context.close()
        logging.info("Proceso completado sin errores.")


def main():
    global USER_DATA_DIR, NAVEGADOR

    parser = argparse.ArgumentParser(description="Descarga reportes mensual y diario de Power BI")
    parser.add_argument(
        "--interactivo",
        action="store_true",
        help="Abre el navegador visible para completar login/MFA manualmente.",
    )
    parser.add_argument(
        "--perfil",
        default=None,
        help=(
            "Nombre del perfil de sesion a usar (ej. --perfil usuario2). "
            "Cada perfil tiene su propia carpeta de sesion/cookies, asi "
            "puedes alternar entre varios usuarios de Microsoft/Azure AD "
            "sin que una sesion sobrescriba a la otra. Si no se indica, "
            "usa el perfil por defecto (el mismo de siempre)."
        ),
    )
    parser.add_argument(
        "--navegador",
        choices=["brave", "edge"],
        default="brave",
        help=(
            "Que navegador usar. 'brave' (por defecto) usa el Brave "
            "instalado en el equipo. 'edge' usa el Microsoft Edge del "
            "sistema -- util para probar si los crashes durante la "
            "descarga son especificos de Brave/su software de seguridad, ya que Edge "
            "usa un binario y una integracion con Windows distintos. "
            "IMPORTANTE: --navegador edge usa su PROPIO perfil de sesion "
            "(carpeta separada), asi que la primera vez necesitas volver "
            "a hacer --interactivo --navegador edge para loguearte ahi "
            "tambien -- no comparte sesion con el perfil de Brave."
        ),
    )
    args = parser.parse_args()

    NAVEGADOR = args.navegador

    if args.perfil:
        USER_DATA_DIR = Path(str(USER_DATA_DIR_BASE) + f"_{args.perfil}")
    else:
        USER_DATA_DIR = USER_DATA_DIR_BASE

    if NAVEGADOR == "edge":
        # Perfil separado por navegador: la sesion de Brave y la de Edge
        # nunca deben mezclarse en la misma carpeta.
        USER_DATA_DIR = Path(str(USER_DATA_DIR) + "_edge")

    setup_logging()
    logging.info(f"Usando perfil de sesion: {USER_DATA_DIR} (navegador: {NAVEGADOR})")

    if args.interactivo:
        with sync_playwright() as playwright:
            modo_interactivo_login(playwright)
        return

    # 6 en vez de 4: el crash del navegador durante la descarga sigue sin
    # tener solucion de raiz (viene del software de seguridad del equipo, ocurre
    # igual en Brave y en Edge), asi que la defensa practica es reintentar
    # hasta que una corrida pase. Cada intento fallido cuesta ~2-3 min,
    # asi que 6 intentos son ~15-18 min en el peor caso -- cabe holgado
    # dentro de la hora que hay entre corridas del orquestador.
    max_intentos_corrida = 6
    for intento in range(1, max_intentos_corrida + 1):
        try:
            correr_descargas()
            break
        except SesionExpiradaError:
            # Detectada tarde (dentro de esperar_reporte_cargado, tras dar
            # vueltas buscando el reporte en lo que era pantalla de login).
            # NO reintentar: la sesion sigue vencida en el intento 2, 3 y 4
            # igual que en el 1 -- solo desperdicia minutos. Se necesita
            # --interactivo. Mismo codigo de salida que la deteccion
            # temprana (sys.exit(2) dentro de correr_descargas) para que
            # cualquier automatizacion externa que revise el exit code lo
            # trate igual en ambos casos.
            logging.error(
                "Sesion vencida detectada. Corre "
                "'python descargar_reportes_powerbi.py --interactivo' para "
                "renovarla. No se reintenta esta corrida."
            )
            sys.exit(2)
        except Exception as e:
            # Se reintenta ante CUALQUIER OTRO error (no solo "navegador
            # cerrado"): timeout esperando el reporte, slicer no
            # encontrado, boton de exportar que no aparecio, frame
            # detached que agoto los reintentos internos, etc. Casi todos
            # estos son transitorios (la pagina fue lenta esa vez, un
            # elemento tardo en renderizar) y una segunda corrida completa
            # desde cero suele resolverlos solos, igual que ya pasaba con
            # "navegador cerrado".
            if intento < max_intentos_corrida:
                espera_seg = 15 * intento  # espera creciente: 15s, 30s, 45s, 60s
                logging.warning(
                    f"Fallo la corrida (intento {intento}/{max_intentos_corrida}): "
                    f"{e}. Esperando {espera_seg}s antes de reintentar la corrida "
                    f"completa..."
                )
                import time
                time.sleep(espera_seg)
                continue
            logging.exception(
                f"Error durante la descarga tras {max_intentos_corrida} "
                f"intentos, se abandona esta corrida: {e}"
            )
            sys.exit(1)


if __name__ == "__main__":
    main()