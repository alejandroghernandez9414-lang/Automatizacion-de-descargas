# Orquestador de descargas

**Automatización en Python que descarga reportes operativos de dos plataformas web a horas fijas, sin intervención manual, y consolida los datos en un Excel listo para análisis.**

Antes este trabajo se hacía a mano varias veces al día: entrar a cada plataforma, filtrar, exportar, esperar y pegar. Ahora un solo programa queda corriendo y lo hace solo.

---

## Cómo funciona

```
                    orquestador.py
            (queda corriendo todo el día)
                         │
        ┌────────────────┴────────────────┐
        ▼                                 ▼
descargar_y_generar_bases.py   descargar_reportes_powerbi.py
  bucle continuo 08:45–15:45     cada hora, al minuto 15
        │                                 │
  3 reportes en paralelo           reporte mensual + diario
  (Playwright)                     (Playwright, sesión persistente)
        │                                 │
        ▼                                 ▼
  BASES_PSR.xlsx                   reporte_*.xlsx
```

| Archivo | Qué hace |
|---|---|
| `orquestador.py` | Programa y lanza cada script en su propio proceso e hilo, con logs por corrida. |
| `descargar_y_generar_bases.py` | Inicia sesión en la plataforma de órdenes, exporta 3 reportes en paralelo y genera un Excel con los cruces y columnas calculadas. |
| `descargar_reportes_powerbi.py` | Abre un reporte de Power BI, selecciona mes y día en un filtro jerárquico y exporta los datos. |
| `config.py` | Lee usuarios, claves, URLs y rutas desde `.env`. |

---

## Problemas que tuve que resolver

- **Una tarea lenta no puede frenar a las demás.** Cada script corre en un proceso aparte dentro de su propio hilo, y un candado evita que la misma tarea se dispare dos veces si la corrida anterior no ha terminado.
- **Inicio de sesión con doble factor en Power BI.** El script reutiliza un perfil persistente del navegador, así el código de verificación se pide una vez al día y no en cada corrida. Si la sesión vence, lo detecta, avisa en el log y termina sin reintentar en vano.
- **Listas virtualizadas.** El filtro de fechas de Power BI solo dibuja los elementos visibles, así que el script hace scroll progresivo hasta encontrar el año, el mes y el día, con un segundo barrido en sentido contrario por si se pasó de largo.
- **El navegador se cierra a mitad de una descarga.** Las descargas van a una carpeta conocida; si el navegador muere después de bajar el archivo pero antes de moverlo, el script lo rescata desde ahí y valida su tamaño.
- **Fallos transitorios.** La corrida completa se reintenta hasta 6 veces con espera creciente, y cada fallo deja captura de pantalla y HTML para diagnóstico.
- **Archivo abierto en Excel.** Se detecta el bloqueo antes de guardar y se informa con un mensaje claro.
- **Lógica heredada de macros VBA.** El ordenamiento, la detección de órdenes repetidas y los cruces entre bases se reescribieron en pandas.

---

## Tecnologías

Python · Playwright · pandas · openpyxl · schedule · python-dotenv

---

## Uso

```
pip install -r requirements.txt
playwright install chromium
```

1. Copiar `.env.example` como `.env` y completar los valores.
2. La primera vez, iniciar sesión en Power BI:
   `python descargar_reportes_powerbi.py --interactivo --navegador edge`
3. Dejar corriendo el orquestador:
   `python orquestador.py`

---

## Seguridad

- Ninguna credencial, URL interna ni ruta de red está en el código: todo se lee de `.env`, que no se sube al repositorio.
- Las claves no se pasan por línea de comandos ni se escriben en los logs.
- El repositorio no incluye datos, reportes descargados ni sesiones del navegador.

---

## Autor

**Luis Alejandro Gamboa Hernández**
