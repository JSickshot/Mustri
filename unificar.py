import pandas as pd
import os

carpeta = r'C:\Users\Monica\OneDrive - Grupo M Arquitectos\LC Monica Soriano\Personas Fisicas\Categorias'
salida = r'C:\Users\Monica\OneDrive - Grupo M Arquitectos\LC Monica Soriano\Personas Fisicas\Categorias\unificado.xlsx'

archivos = [f for f in os.listdir(carpeta) if f.endswith('.xlsx')]

with pd.ExcelWriter(salida, engine='openpyxl') as writer:
    for archivo in archivos:
        ruta_archivo = os.path.join(carpeta, archivo)
        nombre_hoja = os.path.splitext(archivo)[0][:31]

        try:
            df = pd.read_excel(ruta_archivo)
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            print(f"✔: {nombre_hoja}")
        except Exception as e:
            print(f"Error con {archivo}: {e}")

print("todo en orden")
