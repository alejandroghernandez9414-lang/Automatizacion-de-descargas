"""
orquestador.py

Corre los scripts de descarga en sus propios procesos, en los horarios
indicados, cada uno en un hilo separado para que uno lento (hay reportes
que tardan 15-20 min) no bloquee a los demas.

descarga_directa.py es un tercer script OPCIONAL que no forma parte de este
repositorio: si el archivo no existe en la carpeta, su tarea simplemente
no se programa.

Horarios (todos los dias que este programa quede corriendo):
    - descarga_directa.py            -> en punto,  08:00 a 16:00 (9 corridas)
    - descargar_reportes_powerbi.py -> y 15,       08:15 a 16:15 (9 corridas)
    - descargar_y_generar_bases.py  -> en bucle continuo entre 08:45 y 15:45:
                                       arranca, termina, espera 5 min y vuelve
                                       a arrancar (no tiene horas fijas)

Requisitos:
    pip install -r requirements.txt
    Copiar .env.example como .env y completar los valores.

Uso:
    Dejar este script corriendo todo el dia (por ejemplo como tarea de inicio
    de sesion en Task Scheduler, o con pythonw.exe para que no muestre consola).
    NO es un script de "un solo disparo": se queda en un bucle infinito
    revisando la hora.

    python orquestador.py
"""

import logging
import subprocess
import sys
import threading
import time
from datetime import datetime, time as dt_time
from pathlib import Path

import schedule

# ---------------------------------------------------------------------------
# CONFIGURACION
# ---------------------------------------------------------------------------

# Carpeta donde estan los 3 scripts (ajusta si los mueves de lugar).
CARPETA_SCRIPTS = Path(__file__).resolve().parent

SCRIPT_BASES = CARPETA_SCRIPTS / "descargar_y_generar_bases.py"
SCRIPT_POWERBI = CARPETA_SCRIPTS / "descargar_reportes_powerbi.py"
SCRIPT_DIRECTA = CARPETA_SCRIPTS / "descarga_directa.py"

# Las credenciales NO se pasan por linea de comandos (quedarian visibles en
# la lista de procesos y en el log): cada script las lee del archivo .env
# a traves de config.py.

CARPETA_LOGS = CARPETA_SCRIPTS / "logs_orquestador"
CARPETA_LOGS.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(CARPETA_LOGS / "orquestador.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

# Evita que una misma tarea se dispare dos veces en paralelo si la corrida
# anterior todavia no termina (por ejemplo, un reporte tardando mas de lo normal).
_tareas_en_curso = set()
_lock = threading.Lock()


def _ejecutar(nombre_tarea, comando):
    """Corre 'comando' en un proceso aparte y registra el resultado.
    Se llama siempre dentro de un hilo propio para no bloquear el scheduler.
    """
    with _lock:
        if nombre_tarea in _tareas_en_curso:
            logging.warning(
                f"[{nombre_tarea}] Se omite esta corrida: la corrida anterior "
                f"todavia esta en curso."
            )
            return
        _tareas_en_curso.add(nombre_tarea)

    inicio = datetime.now()
    logging.info(f"[{nombre_tarea}] INICIO -> {' '.join(comando)}")
    log_salida = CARPETA_LOGS / f"{nombre_tarea}_{inicio:%Y%m%d_%H%M%S}.log"

    try:
        with open(log_salida, "w", encoding="utf-8") as f:
            resultado = subprocess.run(
                comando,
                cwd=str(CARPETA_SCRIPTS),
                stdout=f,
                stderr=subprocess.STDOUT,
            )
        duracion = (datetime.now() - inicio).total_seconds() / 60
        if resultado.returncode == 0:
            logging.info(
                f"[{nombre_tarea}] OK - duro {duracion:.1f} min "
                f"(log: {log_salida.name})"
            )
        else:
            logging.error(
                f"[{nombre_tarea}] TERMINO CON ERROR (codigo {resultado.returncode}) "
                f"- duro {duracion:.1f} min - revisa {log_salida.name}"
            )
    except Exception as e:
        logging.exception(f"[{nombre_tarea}] Excepcion al lanzar el proceso: {e}")
    finally:
        with _lock:
            _tareas_en_curso.discard(nombre_tarea)


def lanzar_en_hilo(nombre_tarea, comando):
    """Dispara _ejecutar en un hilo nuevo (no bloqueante) para que el
    scheduler pueda seguir revisando otras tareas mientras esta corre."""
    hilo = threading.Thread(
        target=_ejecutar, args=(nombre_tarea, comando), daemon=True
    )
    hilo.start()


def bucle_bases():
    """Corre descargar_y_generar_bases.py en bucle: apenas termina una
    corrida (bien o con error), espera ESPERA_BASES_MIN minutos y lanza la
    siguiente. Solo arranca corridas NUEVAS dentro de la ventana
    VENTANA_BASES_INICIO - VENTANA_BASES_FIN; fuera de ella duerme y revisa
    cada minuto (asi retoma solo al dia siguiente).

    Como _ejecutar() espera a que el subproceso termine, nunca puede haber
    dos corridas de bases solapadas: la siguiente siempre empieza despues
    de que la anterior acabo + la espera.
    """
    comando = [sys.executable, str(SCRIPT_BASES)]
    while True:
        ahora = datetime.now().time()
        if VENTANA_BASES_INICIO <= ahora <= VENTANA_BASES_FIN:
            _ejecutar("descargar_y_generar_bases", comando)  # bloquea hasta terminar
            logging.info(
                f"[descargar_y_generar_bases] Proxima corrida en "
                f"{ESPERA_BASES_MIN} min."
            )
            time.sleep(ESPERA_BASES_MIN * 60)
        else:
            time.sleep(60)


def tarea_powerbi():
    # --navegador edge: se probo que el crash del navegador durante la
    # descarga ("Target page, context or browser has been closed") ocurre
    # tanto en Brave como en Edge -- la causa raiz esta en el equipo
    # (software de seguridad del equipo que fuerza --no-sandbox en cualquier
    # navegador Chromium), no en el navegador. Pero en la practica Edge
    # viene fallando bastante menos que Brave, asi que se usa Edge.
    #
    # OJO: Edge usa su PROPIO perfil de sesion (.pw_powerbi_profile_edge).
    # Cuando la sesion venza (~24h), hay que renovarla con:
    #     python descargar_reportes_powerbi.py --interactivo --navegador edge
    # Si se corre --interactivo SIN --navegador edge, se renueva el perfil
    # de Brave (el equivocado) y las corridas de aqui seguiran fallando.
    lanzar_en_hilo(
        "descargar_reportes_powerbi",
        [sys.executable, str(SCRIPT_POWERBI), "--navegador", "edge"],
    )


def tarea_directa():
    lanzar_en_hilo(
        "descarga_directa",
        [sys.executable, str(SCRIPT_DIRECTA)],
    )


# ---------------------------------------------------------------------------
# PROGRAMACION DE HORARIOS
# ---------------------------------------------------------------------------

# descargar_y_generar_bases.py -> bucle continuo (ver bucle_bases):
# termina, espera ESPERA_BASES_MIN y vuelve a correr. Solo se ARRANCAN
# corridas nuevas dentro de esta ventana; si una corrida empezo a las 15:44
# la deja terminar normal aunque se pase de las 15:45.
VENTANA_BASES_INICIO = dt_time(8, 45)
VENTANA_BASES_FIN = dt_time(15, 45)
ESPERA_BASES_MIN = 5

# descargar_reportes_powerbi.py -> cada hora en el minuto 15, 08:15 a 16:15
HORAS_POWERBI = [
    "08:15", "09:15", "10:15", "11:15", "12:15",
    "13:15", "14:15", "15:15", "16:15",
]

# descarga_directa.py -> cada hora en punto, 08:00 a 16:00
HORAS_DIRECTA = [
    "08:00", "09:00", "10:00", "11:00", "12:00",
    "13:00", "14:00", "15:00", "16:00",
]

for hora in HORAS_POWERBI:
    schedule.every().day.at(hora).do(tarea_powerbi)

HAY_DIRECTA = SCRIPT_DIRECTA.exists()
if HAY_DIRECTA:
    for hora in HORAS_DIRECTA:
        schedule.every().day.at(hora).do(tarea_directa)


def main():
    logging.info("Orquestador iniciado. Horarios programados:")
    logging.info(
        f"  descargar_y_generar_bases  -> bucle continuo "
        f"{VENTANA_BASES_INICIO:%H:%M}-{VENTANA_BASES_FIN:%H:%M}, "
        f"{ESPERA_BASES_MIN} min de espera tras cada corrida"
    )
    logging.info(f"  descargar_reportes_powerbi -> {HORAS_POWERBI}")
    if HAY_DIRECTA:
        logging.info(f"  descarga_directa           -> {HORAS_DIRECTA}")
    else:
        logging.info("  descarga_directa           -> no instalado, se omite")

    threading.Thread(target=bucle_bases, daemon=True).start()

    while True:
        schedule.run_pending()
        time.sleep(20)  # revisa cada 20s; suficiente resolucion para horas fijas


if __name__ == "__main__":
    main()