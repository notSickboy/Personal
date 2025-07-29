import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def main():
    # Inicializar Tkinter oculto
    root = tk.Tk()
    root.withdraw()

    # Selección de archivos
    path_consolidado = filedialog.askopenfilename(
        title="Selecciona el archivo Consolidado",
        filetypes=[("Excel", "*.xlsx;*.xls")]
    )
    if not path_consolidado:
        messagebox.showwarning("Advertencia", "No seleccionaste el Consolidado.")
        return

    path_comparado = filedialog.askopenfilename(
        title="Selecciona el archivo Comparado",
        filetypes=[("CSV", "*.csv")]
    )
    if not path_comparado:
        messagebox.showwarning("Advertencia", "No seleccionaste el Comparado.")
        return

    # Leer datos
    consolidado_df = pd.read_excel(path_consolidado)
    comparado_df    = pd.read_csv(path_comparado)

    # Columnas plantilla
    template_cols = [
        'CANT.', 'SERIE # NOMI', 'N° NOMI', 'OPERADOR', 'FECHA', 'SW PT',
        'PROF.POZO (m)', 'CARGA (KG)', 'NOMBRE PT', 'LINEA', 'PUNTO',
        'COORD. X PT', 'COORD. Y PT', 'COORD. Z PT', 'COORD. X NOMI',
        'COORD. Y NOMI', 'DISTANCIA NOMI-PT (m)', '# EVENTO', 'SW NOMI',
        'VEL. MAX RADIAL (mm/s)', 'VEL. MAX TRANSVERSAL (mm/s)',
        'VEL. MAX VERTICAL (mm/s)', 'VEL. VECTOR SUM (mm/s)',
        'FREC. MAX RADIAL (Hz)', 'FREC. MAX TRANSVERSAL (Hz)',
        'FREC. MAX VERTICAL (Hz)', 'FREC. VECTOR SUMA (HZ)',
        'INTENSIDAD (dBL)', 'FREC. MIC. (Hz)', 'HORA',
        'PERCEPCION DE VELOCIDAD', 'PERCEPCIÓN SOCIAL',
        'IMPACTO ACUSTICO', 'INFRAESTRUCTURA', 'LOCALIDAD',
        'MUNICIPIO', 'ESTADO', 'ZIPPER'
    ]

    # Hacer LEFT JOIN para conservar todas las filas de comparado
    merged = pd.merge(
        comparado_df,
        consolidado_df,
        left_on = ['SERIE # NOMI', 'SPG_Event #'],
        right_on= ['N° Nomi',     'N° de Evento'],
        how='left'
    )

    # Crear DataFrame de resultado usando solo las columnas de la plantilla
    result_df = pd.DataFrame(columns=template_cols)

    # Rellenar columnas
    result_df['CANT.']                       = ''
    result_df['SERIE # NOMI']               = merged['SERIE # NOMI']
    result_df['N° NOMI']                    = ''  # según tu formato original
    result_df['OPERADOR']                   = merged['Operador'].fillna('')
    result_df['FECHA']                      = merged['Fecha'].fillna('')
    result_df['SW PT']                      = ''
    result_df['PROF.POZO (m)']              = 28
    result_df['CARGA (KG)']                 = ''
    result_df['NOMBRE PT']                  = ''
    result_df['LINEA']                      = merged['PFS_Line']
    result_df['PUNTO']                      = merged['PFS_Station']
    result_df['COORD. X PT']                = ''
    result_df['COORD. Y PT']                = ''
    result_df['COORD. Z PT']                = ''
    result_df['COORD. X NOMI']              = merged['Coord X']
    result_df['COORD. Y NOMI']              = merged['Coord Y']
    result_df['DISTANCIA NOMI-PT (m)']      = ''
    result_df['# EVENTO']                   = merged['SPG_Event #']
    result_df['SW NOMI']                    = merged['SW Nomi'].fillna('')
    result_df['VEL. MAX RADIAL (mm/s)']     = merged['SPG_Radial mm/s']
    result_df['VEL. MAX TRANSVERSAL (mm/s)']= merged['SPG_Transverse mm/s']
    result_df['VEL. MAX VERTICAL (mm/s)']   = merged['SPG_Vertical mm/s']
    result_df['VEL. VECTOR SUM (mm/s)']     = merged['SPG_Vector Sum mm/s']
    result_df['FREC. MAX RADIAL (Hz)']      = merged['SPG_Radial (Hz)']
    result_df['FREC. MAX TRANSVERSAL (Hz)'] = merged['SPG_Transverse (Hz)']
    result_df['FREC. MAX VERTICAL (Hz)']    = merged['SPG_Vertical (Hz)']
    result_df['FREC. VECTOR SUMA (HZ)']     = ''
    result_df['INTENSIDAD (dBL)']          = merged['SPG_Air(dBL)']
    result_df['FREC. MIC. (Hz)']           = merged['SPG_Air (Hz)']
    result_df['HORA']                      = merged['SPG_Time']
    # Llenar percepción e impacto solo donde hay datos, dejar vacío en caso contrario
    result_df['PERCEPCION DE VELOCIDAD']   = merged['Percepción de Velocidad'].fillna('')
    result_df['PERCEPCIÓN SOCIAL']         = merged['Percepción Social'].fillna('')
    result_df['IMPACTO ACUSTICO']          = merged['Impacto Audible'].fillna('')
    result_df['INFRAESTRUCTURA']           = merged['Infraestructura'].fillna('')
    result_df['LOCALIDAD']                 = merged['Localidad'].fillna('')
    result_df['MUNICIPIO']                 = ''
    result_df['ESTADO']                    = ''
    result_df['ZIPPER']                    = merged['Zipper'].fillna('')

    # Guardar resultado
    save_path = filedialog.asksaveasfilename(
        title="Guardar resultado como",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")]
    )
    if not save_path:
        messagebox.showwarning("Advertencia", "No seleccionaste ruta de guardado.")
        return

    result_df.to_excel(save_path, index=False)
    messagebox.showinfo("Éxito", f"Archivo guardado en:\n{save_path}")

if __name__ == "__main__":
    main()

