import pandas as pd
from datetime import datetime
from tkinter import Tk, filedialog
import os
from dateutil import parser

# Inicializar Tk
root = Tk()
root.withdraw()

# Seleccionar múltiples archivos SPG
print("Seleccione uno o más archivos SPG (Monitoreo):")
spg_files = filedialog.askopenfilenames(
    title="Seleccione uno o más archivos SPG (Monitoreo)",
    filetypes=[("CSV Files", "*.csv;*.CSV")]
)

# Seleccionar archivo PFS
print("Seleccione el archivo PFS:")
pfs_file = filedialog.askopenfilename(
    title="Seleccione el archivo PFS",
    filetypes=[("CSV Files", "*.csv;*.CSV")]
)

root.update()
root.destroy()


import tkinter as tk
from tkinter import simpledialog, messagebox

def pedir_rango_segundos():
    root = tk.Tk()
    root.withdraw()  # Ocultar ventana principal

    while True:
        try:
            entrada = simpledialog.askstring("Rango de tiempo", "⏱️ Ingresa el rango de segundos de tolerancia:")
            if entrada is None:
                messagebox.showinfo("Cancelado", "No se ingresó ningún valor.")
                return None
            rango = int(entrada)
            if rango < 0:
                raise ValueError
            return rango
        except ValueError:
            messagebox.showerror("Error", "Por favor ingresa un número entero positivo.")

# Rango de segundos de tolerancia
rango_segundos = pedir_rango_segundos()
if rango_segundos is None:
    exit()  # O return, si estás dentro de una función


# Funciones de parseo seguras
def parse_spg_datetime(date_str, time_str):
    try:
        if pd.isna(date_str) or pd.isna(time_str):
            return pd.NaT
        date_clean = str(date_str).strip('#')
        date_dt = datetime.strptime(date_clean, "%Y-%m-%d")

        time_clean = str(time_str).strip().lower() \
            .replace('.', '') \
            .replace('a m', 'am').replace('p m0', 'pm') \
            .replace('a.m.', 'am').replace('p.m.', 'pm') \
            .replace('a.m', 'am').replace('p.m', 'pm') \
            .replace('a. m.', 'am').replace('p. m.', 'pm') \
            .replace('a. m', 'am').replace('p. m', 'pm')

        try:
            time_dt = datetime.strptime(time_clean, "%I:%M %p").time()
        except ValueError:
            try:
                time_dt = datetime.strptime(time_clean, "%H:%M:%S").time()
            except ValueError:
                time_dt = parser.parse(time_clean).time()

        return datetime.combine(date_dt, time_dt)
    except Exception as e:
        print(f"Error al convertir SPG: {date_str} {time_str} -> {e}")
        return pd.NaT

def parse_pfs_datetime(date_str, time_str):
    try:
        if pd.isna(date_str) or pd.isna(time_str):
            return pd.NaT
        date_dt = datetime.strptime(str(date_str), "%m/%d/%Y")
        time_dt = datetime.strptime(str(time_str), "%H:%M:%S").time()
        return datetime.combine(date_dt, time_dt)
    except Exception as e:
        print(f"Error al convertir PFS: {date_str} {time_str} -> {e}")
        return pd.NaT

# Leer y combinar todos los archivos SPG seleccionados
spg_df_list = []
for spg_file in spg_files:
    temp_df = pd.read_csv(spg_file, encoding='utf-8', skip_blank_lines=True)
    temp_df = temp_df.dropna(subset=['Date', 'Time'])
    temp_df['DateTime'] = temp_df.apply(lambda row: parse_spg_datetime(row['Date'], row['Time']), axis=1)
    temp_df = temp_df.dropna(subset=['DateTime'])
    spg_df_list.append(temp_df)

if not spg_df_list:
    print("\n⚠️ No se cargaron datos SPG válidos. Finalizando.")
    exit()

spg_df = pd.concat(spg_df_list, ignore_index=True)

# Leer PFS
pfs_df = pd.read_csv(pfs_file, encoding='utf-8', skip_blank_lines=True)
pfs_df = pfs_df.dropna(subset=['Date', 'Time'])
pfs_df['DateTime'] = pfs_df.apply(lambda row: parse_pfs_datetime(row['Date'], row['Time']), axis=1)
pfs_df = pfs_df.dropna(subset=['DateTime'])

# Buscar coincidencias
resultados = []
print("Buscando coincidencias...")

for idx_spg, row_spg in spg_df.iterrows():
    dt_spg = row_spg['DateTime']
    matches = pfs_df[(pfs_df['DateTime'] - dt_spg).abs().dt.total_seconds() <= rango_segundos]
    for idx_pfs, row_pfs in matches.iterrows():
        fila_resultado = {
            'PFS_Line': row_pfs['Line'],
            'PFS_Station': row_pfs['Station'],
            'PFS_Time': row_pfs['Time'],
            'SPG_Event #': row_spg.get('Event #', ''),
            'SPG_Time': row_spg.get('Time', ''),
            'SPG_Radial mm/s': row_spg.get('Radial mm/s', ''),
            'SPG_Radial (Hz)': row_spg.get('Radial (Hz)', ''),
            'SPG_Transverse mm/s': row_spg.get('Transverse mm/s', ''),
            'SPG_Transverse (Hz)': row_spg.get('Transverse (Hz)', ''),
            'SPG_Vertical mm/s': row_spg.get('Vertical mm/s', ''),
            'SPG_Vertical (Hz)': row_spg.get('Vertical (Hz)', ''),
            'SPG_Vector Sum mm/s': row_spg.get('Vector Sum mm/s', ''),
            'SPG_Air (mb)': row_spg.get('Air (mb)', ''),
            'SPG_Air(dBL)': row_spg.get('Air(dBL)', ''),
            'SPG_Air (Hz)': row_spg.get('Air (Hz)', ''),
            'SPG_Air (kPA)': row_spg.get('Air (kPA)', ''),
            'SERIE # NOMI': row_spg.get('Graph Serial', ''),
            'PFS_Line_dup': row_pfs['Line'],
            'PFS_Station_dup': row_pfs['Station'],
            'PFS_Time_dup': row_pfs['Time'],
        }
        resultados.append(fila_resultado)

# Exportar resultados
if resultados:
    resultados_df = pd.DataFrame(resultados)
    columnas_orden = [
        'PFS_Line', 'PFS_Station', 'PFS_Time',
        'SPG_Event #', 'SPG_Time', 'SPG_Radial mm/s', 'SPG_Radial (Hz)',
        'SPG_Transverse mm/s', 'SPG_Transverse (Hz)', 'SPG_Vertical mm/s', 'SPG_Vertical (Hz)',
        'SPG_Vector Sum mm/s', 'SPG_Air (mb)', 'SPG_Air(dBL)', 'SPG_Air (Hz)', 'SPG_Air (kPA)',
        'SERIE # NOMI',
        'PFS_Line_dup', 'PFS_Station_dup', 'PFS_Time_dup'
    ]
    resultados_df = resultados_df[columnas_orden]
    resultados_df = resultados_df.rename(columns={
        'PFS_Line_dup': 'PFS_Line',
        'PFS_Station_dup': 'PFS_Station',
        'PFS_Time_dup': 'PFS_Time'
    })

       # Mostrar diálogo para guardar archivo
    Tk().withdraw()
    ruta_archivo = filedialog.asksaveasfilename(
        title="Guardar archivo de resultados",
        defaultextension=".csv",
        filetypes=[("Archivos CSV", "*.csv")],
        initialfile="resultado_comparacion_completo.csv"
    )

    if not ruta_archivo:
        messagebox.showinfo("Operación cancelada", "No se seleccionó una ruta para guardar.")
        print("🚫 Operación cancelada por el usuario.")
    else:
        try:
            resultados_df.to_csv(ruta_archivo, index=False, encoding='utf-8-sig')
            print(f"\n✅ Archivo generado con {len(resultados_df)} coincidencias en:\n{ruta_archivo}")
        except Exception as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo:\n{e}")
            print(f"❌ Error al guardar el archivo: {e}")
else:
    print("\n⚠️ No se encontraron coincidencias con el rango especificado.")
