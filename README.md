# Extractor de Datos Semanales de Excel

Script en Python para extraer automáticamente datos de la semana pasada de archivos Excel (.xlsx) basándose en las columnas de fecha de matrícula (columna M) y fecha de renovación (columna N).

## 📋 Características

- ✅ Calcula automáticamente el rango de la semana pasada (lunes a domingo)
- ✅ Lee múltiples archivos .xlsx de una carpeta
- ✅ Filtra datos por columnas M (matrícula) y N (renovación)
- ✅ Soporta formato de fecha DD/MM/AAAA
- ✅ Crea archivos solo si hay datos nuevos (no genera archivos vacíos)
- ✅ Mantiene la estructura original de los datos
- ✅ Configuración mediante archivo .env

## 🚀 Instalación

1. Asegúrate de tener Python 3.8 o superior instalado

2. Instala las dependencias:
```powershell
pip install -r requirements.txt
```

## ⚙️ Configuración

Edita el archivo `.env` para especificar las rutas:

```env
CARPETA_ARCHIVOS=C:\ruta\a\tus\archivos\excel
CARPETA_SALIDA=C:\ruta\para\guardar\resultados
```

- **CARPETA_ARCHIVOS**: Carpeta donde se encuentran los archivos .xlsx originales
- **CARPETA_SALIDA**: Carpeta donde se guardarán los archivos con datos filtrados

## 📖 Uso

Ejecuta el script:

```powershell
python extractor_datos.py
```

El script:
1. Calculará el rango de la semana pasada
2. Leerá todos los archivos .xlsx de la carpeta configurada
3. Filtrará los datos por las columnas M y N
4. Guardará nuevos archivos con el sufijo `_semana_pasada.xlsx`
5. Mostrará un resumen del proceso

## 📊 Estructura de Datos

El script espera que los archivos Excel tengan:
- **Columna M** (índice 12): Fecha de matrícula
- **Columna N** (índice 13): Fecha de renovación
- **Formato de fecha**: DD/MM/AAAA

## 📁 Estructura del Proyecto

```
Prueba pythonCompite/
├── Data/                    # Carpeta para archivos de entrada (configurable)
├── Output/                  # Carpeta para archivos de salida (se crea automáticamente)
├── extractor_datos.py       # Script principal
├── requirements.txt         # Dependencias de Python
├── .env                     # Configuración de rutas
└── README.md               # Este archivo
```

## 🔍 Ejemplo de Salida

```
======================================================================
  EXTRACTOR DE DATOS SEMANALES DE EXCEL
======================================================================

📅 Buscando datos de la semana pasada:
   Desde: 28/10/2025
   Hasta: 03/11/2025

📂 Carpeta de entrada: C:\Users\...\Data
📂 Carpeta de salida: C:\Users\...\Output
----------------------------------------------------------------------

🔍 Se encontraron 3 archivo(s) .xlsx

📄 Procesando: datos_alumnos.xlsx
  ✓ Archivo leído correctamente (250 filas)
  📋 Columna matrícula: 'Fecha Matrícula'
  📋 Columna renovación: 'Fecha Renovación'
  ✅ 15 registro(s) encontrado(s)
  💾 Guardado en: datos_alumnos_semana_pasada.xlsx

======================================================================
📊 RESUMEN
======================================================================
Archivos procesados: 3
Archivos con datos nuevos: 1
Total de registros extraídos: 15

✓ Proceso completado
```

## ⚠️ Notas Importantes

- El script no generará archivos vacíos si no hay datos de la semana pasada
- Mantiene la estructura y formato original de los datos
- Copia las filas completas que cumplan con el criterio de fecha
- Crea automáticamente la carpeta de salida si no existe

## 🛠️ Solución de Problemas

**Error: "No se encontró la variable CARPETA_ARCHIVOS"**
- Verifica que el archivo `.env` existe en el mismo directorio que el script
- Asegúrate de que las variables están correctamente definidas

**Error: "La carpeta no existe"**
- Verifica que la ruta en `CARPETA_ARCHIVOS` es correcta y existe
- Usa rutas absolutas en Windows (ej: `C:\Users\...`)

**No se encuentran datos**
- Verifica que las columnas M y N contienen fechas en formato DD/MM/AAAA
- Confirma que hay datos de la semana pasada en los archivos
