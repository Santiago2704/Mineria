# Minería de prospectos de crédito

Herramienta para identificar y priorizar empresas candidatas a productos de
crédito (simple, PyME/capital de trabajo, línea revolvente, arrendamiento de
equipo), a partir del **DENUE** (Directorio Estadístico Nacional de Unidades
Económicas) de INEGI. Prioriza el Área Metropolitana de Monterrey y Nuevo
León, pero funciona con cualquier estado del país.

## 1. Obtener los datos (tú, fuera de esta sesión)

Este entorno no tiene salida de red hacia inegi.org.mx, así que la descarga
la debes hacer tú:

1. Entra a la descarga masiva del DENUE:
   https://www.inegi.org.mx/app/mapa/denue/ (o desde
   https://www.inegi.org.mx/app/descarga/default.html buscando "DENUE").
2. Selecciona **Nuevo León** (o el estado que quieras analizar) y descarga en
   formato **CSV**.
3. Descomprime el `.zip` y sube el archivo `denue_inegi_19_*.csv` a este
   proyecto, en la carpeta `data/` (ej. `data/denue_nuevo_leon.csv`).

No necesitas token ni cuenta para la descarga masiva por CSV (el token solo
es necesario para la API en vivo).

## 2. Ejecutar el análisis

```bash
pip install -r requirements.txt
python scripts/generar_prospectos.py data/denue_nuevo_leon.csv --salida prospectos.xlsx
```

Puedes pasar varios CSV (por ejemplo Nuevo León + otro estado) en el mismo
comando; el script los combina.

## 3. Qué hace el script

- **Filtra** giros que normalmente no son sujetos de crédito empresarial
  (unidades económicas del gobierno, asociaciones religiosas, etc.).
- **Prioriza por ubicación**: Área Metropolitana de Monterrey > resto de
  Nuevo León > resto de México.
- **Calcula un score de 0-100** combinando: tamaño de la empresa (rango de
  empleados), qué tan formalizada está (tiene teléfono/correo/sitio web) y el
  giro (sectores con más rotación de capital o necesidad típica de
  financiamiento puntúan más alto).
- **Sugiere producto de crédito** por empresa según su sector y tamaño
  (crédito simple, PyME/capital de trabajo, línea revolvente o arrendamiento
  de equipo).
- **Exporta un Excel** (`prospectos.xlsx`) con una hoja por prioridad de zona,
  ordenado por score descendente, listo para que el equipo de banca de
  empresas empiece a contactar.

Ver `scripts/generar_prospectos.py` para ajustar los criterios de scoring o
la lista de municipios del AMM.
