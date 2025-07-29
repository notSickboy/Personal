import pandas as pd
import os
from tkinter import filedialog, Tk,messagebox

def consolidar_archivos_en_carpeta():
    root = Tk()
    root.withdraw()

    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los reportes .xlsx")
    if not carpeta:
        print("No se seleccionó carpeta.")
        return

    """
    Enlista todos los archivos de una dirección y los almacena en la variable
    si terminan con la extensión .xls o .xlsx
    """

    archivos = [f for f in os.listdir(carpeta) if f.endswith(('.xls', '.xlsx'))]

    if not archivos:
        print("No se encontraron archivos Excel en la carpeta.")
        return
    
    ruta_archivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")],
        title="Guardar archivo consolidado como...",
        initialfile="Consolidado.xlsx")

    lista_df = []

    for archivo in archivos:
        # Consigue la ruta de cada archivo en la carpeta
        path = os.path.join(carpeta, archivo)
        try:
            """ Guarda 4 variables con los datos que se repiten, los toma de la columna B
                de la ruta path
            """
            encabezados = pd.read_excel(path, header=None, nrows=4, usecols='B')
            operador = encabezados.iloc[0, 0]
            n_nomi = encabezados.iloc[1, 0]
            fecha = encabezados.iloc[2, 0]
            zipper = encabezados.iloc[3, 0]


            # Lee el excel de la ruta path y los lee desde la fila 6
            df = pd.read_excel(path, header=5)

            # Elimina los espacios antes y después del nombre en las celdas de título
            df.columns = [str(c).strip() for c in df.columns]
            
            # Crea nuevas columnas en el df y les asigna los valores
            df['N° Nomi'] = n_nomi
            df['Operador'] = operador
            df['Fecha'] = fecha
            df['Zipper'] = zipper


            # Renombra las columnas según el diccionario, izquierda es antes y derecha
            # después del cambio
            # Renombramiento basado en coincidencia parcial
            mapa_columnas = {
                'linea': 'Linea',
                'punto': 'Punto',
                'coordenada x': 'Coord X',
                'coordenada y': 'Coord Y',
                'CoordX': 'Coord X',
                'CoordY': 'Coord Y',
                'evento': 'N° de Evento',
                'swnomi': 'SW Nomi',
                'sw nomi': 'SW Nomi',
                'hora': 'Hora',
                'infraestructura': 'Infraestructura',
                'localidad': 'Localidad',
                'comentarios': 'Localidad',
                'percep.': 'Percepción de Velocidad',
                'percepción de velocidad': 'Percepción de Velocidad',
                'impacto social': 'Percepción Social',
                'imp. social': 'Percepción Social',
                'impacto audible': 'Impacto Audible',
                'imp. audible': 'Impacto Audible',
                'coordx': 'Coord X',
                'coordy': 'Coord Y',

            }

            # Convertimos todos los nombres de columnas a minúsculas y los limpiamos
            nuevos_nombres = {}
            for col in df.columns:
                nombre = col.lower().strip()
                for clave, nuevo in mapa_columnas.items():
                    if clave in nombre:
                        nuevos_nombres[col] = nuevo
                        break

            df.rename(columns=nuevos_nombres, inplace=True)


            # Convertir columnas problemáticas a numérico
            columnas_a_convertir = [
                'Coord X',
                'Coord Y',
                'N° de Evento',
                'SW Nomi',
                'Percepción de Velocidad',
                'Percepción Social',
                'Impacto Audible'
            ]

            for col in columnas_a_convertir:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')


            # Limpia el DataFrame eliminando columnas innecesarias.
            columnas_ordenadas = [
                'N° Nomi',
                'Operador',
                'Fecha',
                'Linea',
                'Punto',
                'Coord X',
                'Coord Y',
                'N° de Evento',
                'SW Nomi',
                'Hora',
                'Infraestructura',
                'Localidad',
                'Percepción de Velocidad',
                'Percepción Social',
                'Impacto Audible',
                'Zipper'
            ]
            df = df[[col for col in columnas_ordenadas if col in df.columns]]
            
            # Se agrega cada df creado a la variable lista
            lista_df.append(df)
            print(f"✅ Procesado: {archivo} ({len(df)} filas)")

        except Exception as e:
            print(f"❌ Error procesando {archivo}: {e}")

    if lista_df:

        # Todos los archivos procesados se concatenan verticalmente
        #  en un solo DataFrame con todas las filas.
        df_final = pd.concat(lista_df, ignore_index=True)

        # COMPLETAR VALORES VACÍOS SOLO SI EL OPERADOR ES EL MISMO
        columnas_a_completar = ['Linea', 'Punto', 'Coord X', 'Coord Y',
                                'N° de Evento', 'SW Nomi', 'Infraestructura', 'Localidad']
        
        # Para cada columna, aplicamos fillna agrupando por Operador
        for col in columnas_a_completar:
            if col in df_final.columns:
                df_final[col] = df_final.groupby('Operador')[col].ffill()

    # Guardar el archivo
    try:
        df_final.to_excel(ruta_archivo, index=False)
        print(f"\n✅ Consolidación completada. Total filas: {len(df_final)}")
        print(f"📁 Archivo guardado como: {ruta_archivo}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el archivo.\n{e}")
        print(f"❌ Error al guardar el archivo: {e}")        

if __name__ == "__main__":
    consolidar_archivos_en_carpeta()


