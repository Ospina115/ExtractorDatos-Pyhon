"""
Script para extraer datos de la semana pasada de archivos Excel.
Filtra por columnas M (fecha matrícula) y N (fecha renovación).
"""

import os
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()

def obtener_rango_semana_pasada():
    """
    Obtiene el rango de fechas de la semana pasada (lunes a domingo).
    
    Returns:
        tuple: (fecha_inicio, fecha_fin) de la semana pasada
    """
    hoy = datetime.now()
    
    # Calcular el lunes de esta semana
    dias_desde_lunes = hoy.weekday()  # 0 = lunes, 6 = domingo
    lunes_esta_semana = hoy - timedelta(days=dias_desde_lunes)
    
    # Calcular el lunes y domingo de la semana pasada
    lunes_semana_pasada = lunes_esta_semana - timedelta(days=7)
    domingo_semana_pasada = lunes_semana_pasada + timedelta(days=6)
    
    # Normalizar a inicio y fin del día
    fecha_inicio = lunes_semana_pasada.replace(hour=0, minute=0, second=0, microsecond=0)
    fecha_fin = domingo_semana_pasada.replace(hour=23, minute=59, second=59, microsecond=999999)
    
    return fecha_inicio, fecha_fin


def convertir_fecha(valor_fecha):
    """
    Convierte un valor de fecha a datetime, manejando diferentes formatos.
    
    Args:
        valor_fecha: Valor que puede ser string, datetime, o None
        
    Returns:
        datetime o None si no se puede convertir
    """
    if pd.isna(valor_fecha):
        return None
    
    if isinstance(valor_fecha, datetime):
        return valor_fecha
    
    if isinstance(valor_fecha, str):
        try:
            # Formato DD/MM/AAAA
            return datetime.strptime(valor_fecha, '%d/%m/%Y')
        except ValueError:
            try:
                # Intentar otros formatos comunes
                return datetime.strptime(valor_fecha, '%d-%m-%Y')
            except ValueError:
                return None
    
    return None


def filtrar_datos_semana_pasada(df, fecha_inicio, fecha_fin):
    """
    Filtra el DataFrame por fechas de matrícula (columna M) y renovación (columna N).
    
    Args:
        df: DataFrame con los datos
        fecha_inicio: Fecha de inicio de la semana pasada
        fecha_fin: Fecha de fin de la semana pasada
        
    Returns:
        DataFrame filtrado con los datos de la semana pasada
    """
    # Obtener los nombres de las columnas M y N (índices 12 y 13)
    columnas = df.columns
    
    if len(columnas) < 14:
        print(f"  ⚠ El archivo no tiene suficientes columnas (tiene {len(columnas)}, se esperan al menos 14)")
        return pd.DataFrame()
    
    columna_matricula = columnas[12]  # Columna M (índice 12)
    columna_renovacion = columnas[13]  # Columna N (índice 13)
    
    print(f"  📋 Columna matrícula: '{columna_matricula}'")
    print(f"  📋 Columna renovación: '{columna_renovacion}'")
    
    # Convertir las columnas a datetime
    df['fecha_matricula_dt'] = df[columna_matricula].apply(convertir_fecha)
    df['fecha_renovacion_dt'] = df[columna_renovacion].apply(convertir_fecha)
    
    # Filtrar filas donde alguna de las fechas esté en el rango
    mascara = (
        ((df['fecha_matricula_dt'] >= fecha_inicio) & (df['fecha_matricula_dt'] <= fecha_fin)) |
        ((df['fecha_renovacion_dt'] >= fecha_inicio) & (df['fecha_renovacion_dt'] <= fecha_fin))
    )
    
    # Eliminar las columnas auxiliares antes de retornar
    df_filtrado = df[mascara].copy()
    df_filtrado = df_filtrado.drop(columns=['fecha_matricula_dt', 'fecha_renovacion_dt'])
    
    return df_filtrado


def procesar_archivos():
    """
    Procesa todos los archivos Excel en la carpeta especificada.
    """
    # Obtener rutas de las variables de entorno
    carpeta_entrada = os.getenv('CARPETA_ARCHIVOS')
    carpeta_salida = os.getenv('CARPETA_SALIDA')
    
    if not carpeta_entrada:
        print("❌ Error: No se encontró la variable CARPETA_ARCHIVOS en el archivo .env")
        return
    
    if not carpeta_salida:
        print("❌ Error: No se encontró la variable CARPETA_SALIDA en el archivo .env")
        return
    
    # Convertir a Path
    ruta_entrada = Path(carpeta_entrada)
    ruta_salida = Path(carpeta_salida)
    
    # Verificar que la carpeta de entrada existe
    if not ruta_entrada.exists():
        print(f"❌ Error: La carpeta '{carpeta_entrada}' no existe")
        return
    
    # Crear carpeta de salida si no existe
    ruta_salida.mkdir(parents=True, exist_ok=True)
    
    # Obtener rango de fechas de la semana pasada
    fecha_inicio, fecha_fin = obtener_rango_semana_pasada()
    print(f"\n📅 Buscando datos de la semana pasada:")
    print(f"   Desde: {fecha_inicio.strftime('%d/%m/%Y')}")
    print(f"   Hasta: {fecha_fin.strftime('%d/%m/%Y')}")
    print(f"\n📂 Carpeta de entrada: {ruta_entrada}")
    print(f"📂 Carpeta de salida: {ruta_salida}")
    print("-" * 70)
    
    # Buscar archivos .xlsx
    archivos_xlsx = list(ruta_entrada.glob('*.xlsx'))
    
    if not archivos_xlsx:
        print(f"\n⚠ No se encontraron archivos .xlsx en '{carpeta_entrada}'")
        return
    
    print(f"\n🔍 Se encontraron {len(archivos_xlsx)} archivo(s) .xlsx\n")
    
    archivos_procesados = 0
    archivos_con_datos = 0
    total_registros = 0
    
    # Procesar cada archivo
    for archivo in archivos_xlsx:
        print(f"\n📄 Procesando: {archivo.name}")
        
        try:
            # Leer el archivo Excel
            df = pd.read_excel(archivo)
            print(f"  ✓ Archivo leído correctamente ({len(df)} filas)")
            
            # Filtrar datos de la semana pasada
            df_filtrado = filtrar_datos_semana_pasada(df, fecha_inicio, fecha_fin)
            
            if len(df_filtrado) > 0:
                # Crear nombre del archivo de salida
                nombre_base = archivo.stem
                nombre_salida = f"{nombre_base}_semana_pasada.xlsx"
                ruta_archivo_salida = ruta_salida / nombre_salida
                
                # Guardar el archivo
                df_filtrado.to_excel(ruta_archivo_salida, index=False)
                
                print(f"  ✅ {len(df_filtrado)} registro(s) encontrado(s)")
                print(f"  💾 Guardado en: {nombre_salida}")
                
                archivos_con_datos += 1
                total_registros += len(df_filtrado)
            else:
                print(f"  ℹ No se encontraron datos de la semana pasada")
            
            archivos_procesados += 1
            
        except Exception as e:
            print(f"  ❌ Error al procesar el archivo: {str(e)}")
            continue
    
    # Resumen final
    print("\n" + "=" * 70)
    print("📊 RESUMEN")
    print("=" * 70)
    print(f"Archivos procesados: {archivos_procesados}")
    print(f"Archivos con datos nuevos: {archivos_con_datos}")
    print(f"Total de registros extraídos: {total_registros}")
    print(f"\n✓ Proceso completado")


if __name__ == "__main__":
    print("=" * 70)
    print("  EXTRACTOR DE DATOS SEMANALES DE EXCEL")
    print("=" * 70)
    procesar_archivos()
