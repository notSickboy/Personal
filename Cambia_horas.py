import pandas as pd
import re
import os
from tkinter import Tk, messagebox,filedialog
from tkinter.filedialog import askopenfilename, askopenfilenames


def extraer_eventos(texto):
    # Extrae: Evento, Hora, Graph
    patron = r'Event\s+#(\d+)\s*/\s*\d{2}/\d{2}/\d{4}\s+(\d{2}:\d{2}:\d{2})\s+[ap]\. m\.\s+Graph:\s+(\d+)'
    coincidencias = re.findall(patron, texto)
    df = pd.DataFrame(coincidencias, columns=['Evento', 'Hora', 'Graph'])
    df['Evento'] = df['Evento'].astype(int)
    df['Graph'] = df['Graph'].astype(str)
    return df

def main():
    # Oculta la ventana principal de tkinter
    Tk().withdraw()

    archivo_txt = askopenfilename(title='Selecciona el archivo .txt con los eventos y horas...',filetypes=[("Text files", "*.txt")])
    if not archivo_txt:
        messagebox.showwarning("Advertencia", "No se seleccionó archivo de texto.")
        return

    with open(archivo_txt, 'r', encoding='utf-8') as f:
        contenido_txt = f.read()

    df_eventos = extraer_eventos(contenido_txt)

    archivos_csv = askopenfilenames(title='Selecciona uno o más archivos .CSV(SPG) para modificar...',filetypes=[("CSV Files", "*.csv")])

    ruta_salida = filedialog.askdirectory(title="Selecciona la carpeta donde guardar los archivos actualizados")

    if not archivos_csv:
        print("❌ No se seleccionaron archivos CSV.")
        return

    for archivo_csv in archivos_csv:
        print(f"\nProcesando archivo: {os.path.basename(archivo_csv)}")

        # Leer CSV
        df_csv = pd.read_csv(archivo_csv)

        # Verificar columnas necesarias
        if not {'Event #', 'Graph Serial', 'Time'}.issubset(df_csv.columns):
            print(f"❌ Columnas requeridas no encontradas en {archivo_csv}")
            continue

        # Convertir tipos compatibles para comparación
        df_csv['Event #'] = pd.to_numeric(df_csv['Event #'], errors='coerce').astype('Int64')
        df_csv['Graph Serial'] = pd.to_numeric(df_csv['Graph Serial'], errors='coerce').dropna().astype(int).astype(str)

        # Reemplazo de horas
        reemplazos = 0
        for i, row in df_csv.iterrows():
            evento = row['Event #']
            graph = row['Graph Serial']

            if pd.notna(evento):
                match = df_eventos[
                    (df_eventos['Evento'] == evento) &
                    (df_eventos['Graph'] == graph)
                ]
                if not match.empty:
                    nueva_hora = match.iloc[0]['Hora']
                    df_csv.at[i, 'Time'] = nueva_hora
                    reemplazos += 1

        print(f"✅ Reemplazos realizados: {reemplazos}")

        # Guardar cada archivo con sufijo "_actualizado"
        nombre_salida = os.path.splitext(os.path.basename(archivo_csv))[0] + "_actualizado.csv"
        salida_completa = os.path.join(ruta_salida, nombre_salida)

        df_csv.to_csv(salida_completa, index=False)
        print(f"💾 Guardado en: {salida_completa}")

if __name__ == "__main__":
    main()