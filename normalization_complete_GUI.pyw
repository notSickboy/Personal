import os
import math
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# (Opcional) Mejor manejo de DPI en Windows
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# === Matplotlib para vista previa ===
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ==========================
#   UTILIDADES / CORE
# ==========================
def to_numeric(s):
    return pd.to_numeric(s, errors="coerce")

def safe_median(s):
    s = to_numeric(s)
    return float(s.dropna().median()) if s.notna().any() else np.nan

def safe_mean(s):
    s = to_numeric(s)
    return float(s.dropna().mean()) if s.notna().any() else np.nan

def iqr_mask(series, k=1.5):
    x = to_numeric(series)
    q1 = x.quantile(0.25)
    q3 = x.quantile(0.75)
    i = q3 - q1
    lo = q1 - k*i
    hi = q3 + k*i
    return (x >= lo) & (x <= hi)

def zscore_mask(series, z=3.0):
    x = to_numeric(series)
    mu = x.mean()
    sd = x.std(ddof=0)
    if sd == 0 or pd.isna(sd):
        return pd.Series([True]*len(x), index=x.index)
    return (np.abs((x - mu) / sd) <= z)

def estimate_n_loglog(SD_vals, Z_vals, min_points=3, base="e"):
    """
    Ajuste log-log: log(Z) = a + b*log(SD) => n = -b.
    base: 'e' o '10' (afecta reporte de k únicamente).
    """
    sd = to_numeric(SD_vals)
    z  = to_numeric(Z_vals)

    m = (sd > 0) & (z > 0)
    sd = sd[m]; z = z[m]
    if len(sd) < min_points:
        return dict(ok=False, n=np.nan, r2=np.nan, k=np.nan, a=np.nan, b=np.nan, n_points=len(sd))

    if base == "10":
        log_sd = np.log10(sd.values)
        log_z  = np.log10(z.values)
        b, a = np.polyfit(log_sd, log_z, 1)  # pendiente, intercepto en base10
        y_pred = a + b*log_sd
        n = -b
        k = 10**a
    else:
        log_sd = np.log(sd.values)
        log_z  = np.log(z.values)
        b, a = np.polyfit(log_sd, log_z, 1)  # pendiente, intercepto en base e
        y_pred = a + b*log_sd
        n = -b
        k = math.e**a

    ss_res = np.sum((log_z - y_pred)**2)
    ss_tot = np.sum((log_z - np.mean(log_z))**2)
    r2 = 1 - ss_res/ss_tot if ss_tot > 0 else np.nan
    return dict(ok=True, n=n, r2=r2, k=k, a=a, b=b, n_points=len(sd))

def choose_stat(series, mode="Median"):
    if mode == "Mean":
        return safe_mean(series)
    return safe_median(series)

def apply_outlier_filter(df, col_z, col_sd, mode="None", param=1.5, zthr=3.0):
    """
    Devuelve máscara booleana de filas válidas, combinando SD>0, Z>0 y filtro de outliers
    (en log-espacio para mayor estabilidad).
    """
    z = to_numeric(df[col_z])
    sd = to_numeric(df[col_sd])
    m = (z > 0) & (sd > 0)

    if mode == "IQR":
        with np.errstate(invalid='ignore'):
            m &= iqr_mask(np.log(z), k=param)
            m &= iqr_mask(np.log(sd), k=param)
    elif mode == "Z-score":
        with np.errstate(invalid='ignore'):
            m &= zscore_mask(np.log(z), z=zthr)
            m &= zscore_mask(np.log(sd), z=zthr)
    return m

def clip_if_needed(arr, vmin=None, vmax=None):
    if vmin is None and vmax is None:
        return arr
    a = arr.copy()
    if vmin is not None:
        a = np.maximum(a, vmin)
    if vmax is not None:
        a = np.minimum(a, vmax)
    return a

# ==========================
#   GUI
# ==========================
class NormalizationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Normalización para Interpolación — GUI")
        # Ventana más grande + redimensionable
        self.geometry("1100x900")
        self.minsize(1000, 820)
        self.resizable(True, True)

        # ---------- ESTILO ----------
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TLabel", padding=2)
        style.configure("TButton", padding=6)
        style.configure("TEntry", padding=3)
        style.configure("TCombobox", padding=3)

        # DataFrame cargado
        self.df = None

        self.file_in = tk.StringVar(value="")
        self.file_out = tk.StringVar(value="")

        # Columnas por defecto (editables)
        self.col_x = tk.StringVar(value="COORD. X NOMI")
        self.col_y = tk.StringVar(value="COORD. Y NOMI")
        self.col_carga = tk.StringVar(value="CARGA (KG)")
        self.col_dist = tk.StringVar(value="DISTANCIA NOMI-PT (m)")

        # SD_ref
        self.sdref_mode = tk.StringVar(value="Median")
        self.sdref_custom = tk.StringVar(value="")

        # Regresión
        self.reg_mode = tk.StringVar(value="OLS (log-log)")
        self.fixed_n = tk.StringVar(value="1.6")
        self.min_points = tk.StringVar(value="3")
        self.log_base = tk.StringVar(value="e")  # e o 10

        # Outliers
        self.outlier_mode = tk.StringVar(value="None")  # None, IQR, Z-score
        self.iqr_k = tk.StringVar(value="1.5")
        self.z_thr = tk.StringVar(value="3.0")

        # Clipping
        self.clip_min = tk.StringVar(value="")  # factor mínimo del ajuste (SD/SDref)^n
        self.clip_max = tk.StringVar(value="")

        # Normalización (por estadístico) y frecuencias
        self.norm_stat = tk.StringVar(value="Median")  # Median o Mean
        self.freq_action = tk.StringVar(value="Normalize by statistic")  # None / Normalize by statistic
        self.freq_adjust_by_sd = tk.BooleanVar(value=False)

        # Extras
        self.export_adjusted = tk.BooleanVar(value=True)  # exportar columna Ajustada@SD_ref
        self.export_params = tk.BooleanVar(value=True)
        self.export_QA = tk.BooleanVar(value=True)

        # Vista previa
        self.preview_y_col = tk.StringVar(value="")
        self.preview_series_mode = tk.StringVar(value="Original")  # Original / Ajustada@SD_ref / Normalizada
        self.preview_logx = tk.BooleanVar(value=True)
        self.preview_logy = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self):
        # === Frame de Archivo ===
        frm_io = ttk.LabelFrame(self, text="Archivo")
        frm_io.place(x=20, y=15, width=1040, height=120)
        frm_io.grid_columnconfigure(1, weight=1)

        ttk.Label(frm_io, text="Entrada BASE DE DATOS NOMIS (.xlsx, headers en fila 3):")\
            .grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(frm_io, textvariable=self.file_in)\
            .grid(row=0, column=1, padx=5, sticky="we")
        ttk.Button(frm_io, text="Seleccionar...", command=self.select_input)\
            .grid(row=0, column=2, padx=5)

        ttk.Label(frm_io, text="Salida (.xlsx):")\
            .grid(row=1, column=0, sticky="w", padx=5)
        ttk.Entry(frm_io, textvariable=self.file_out)\
            .grid(row=1, column=1, padx=5, sticky="we")
        ttk.Button(frm_io, text="Guardar como...", command=self.select_output)\
            .grid(row=1, column=2, padx=5)

        # === Frame Columnas ===
        frm_cols = ttk.LabelFrame(self, text="Columnas (puedes editar si difieren)")
        frm_cols.place(x=20, y=145, width=1040, height=110)

        ttk.Label(frm_cols, text="X:").grid(row=0, column=0, sticky="e", padx=5, pady=3)
        ttk.Entry(frm_cols, textvariable=self.col_x, width=32).grid(row=0, column=1, padx=5)
        ttk.Label(frm_cols, text="Y:").grid(row=0, column=2, sticky="e", padx=5)
        ttk.Entry(frm_cols, textvariable=self.col_y, width=32).grid(row=0, column=3, padx=5)

        ttk.Label(frm_cols, text="Carga (kg):").grid(row=1, column=0, sticky="e", padx=5)
        ttk.Entry(frm_cols, textvariable=self.col_carga, width=32).grid(row=1, column=1, padx=5)
        ttk.Label(frm_cols, text="Distancia (m):").grid(row=1, column=2, sticky="e", padx=5)
        ttk.Entry(frm_cols, textvariable=self.col_dist, width=32).grid(row=1, column=3, padx=5)

        # === Frame SD_ref ===
        frm_sdref = ttk.LabelFrame(self, text="SD_ref (Distancia Escalada de referencia)")
        frm_sdref.place(x=20, y=265, width=510, height=120)

        ttk.Label(frm_sdref, text="Modo:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        cb_sd = ttk.Combobox(frm_sdref, textvariable=self.sdref_mode, width=24, state="readonly",
                             values=["Median", "Mean", "1", "Custom"])
        cb_sd.grid(row=0, column=1, padx=5)
        ttk.Label(frm_sdref, text="Custom:").grid(row=1, column=0, sticky="e", padx=5)
        ttk.Entry(frm_sdref, textvariable=self.sdref_custom, width=26).grid(row=1, column=1, padx=5)

        # === Frame Regresión ===
        frm_reg = ttk.LabelFrame(self, text="Regresión / Exponente n")
        frm_reg.place(x=550, y=265, width=510, height=120)

        ttk.Label(frm_reg, text="Modo:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        cb_rm = ttk.Combobox(frm_reg, textvariable=self.reg_mode, width=24, state="readonly",
                             values=["OLS (log-log)", "Fixed n (override)"])
        cb_rm.grid(row=0, column=1, padx=5)

        ttk.Label(frm_reg, text="n fijo:").grid(row=1, column=0, sticky="e", padx=5)
        ttk.Entry(frm_reg, textvariable=self.fixed_n, width=10).grid(row=1, column=1, sticky="w", padx=5)

        ttk.Label(frm_reg, text="Puntos mínimos:").grid(row=1, column=2, sticky="e", padx=5)
        ttk.Entry(frm_reg, textvariable=self.min_points, width=8).grid(row=1, column=3, padx=5)

        ttk.Label(frm_reg, text="Log base:").grid(row=2, column=0, sticky="e", padx=5)
        cb_lb = ttk.Combobox(frm_reg, textvariable=self.log_base, width=10, state="readonly", values=["e", "10"])
        cb_lb.grid(row=2, column=1, sticky="w", padx=5)

        # === Frame Outliers / Clipping ===
        frm_out = ttk.LabelFrame(self, text="Outliers y Clipping")
        frm_out.place(x=20, y=395, width=1040, height=120)

        ttk.Label(frm_out, text="Filtro outliers:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        cb_om = ttk.Combobox(frm_out, textvariable=self.outlier_mode, width=16, state="readonly",
                             values=["None", "IQR", "Z-score"])
        cb_om.grid(row=0, column=1, padx=5)

        ttk.Label(frm_out, text="IQR k:").grid(row=0, column=2, sticky="e", padx=5)
        ttk.Entry(frm_out, textvariable=self.iqr_k, width=8).grid(row=0, column=3, padx=5)

        ttk.Label(frm_out, text="Z-score thr:").grid(row=0, column=4, sticky="e", padx=5)
        ttk.Entry(frm_out, textvariable=self.z_thr, width=8).grid(row=0, column=5, padx=5)

        ttk.Label(frm_out, text="Clipping factor (min/max):").grid(row=1, column=0, sticky="e", padx=5)
        ttk.Entry(frm_out, textvariable=self.clip_min, width=10).grid(row=1, column=1, padx=5)
        ttk.Entry(frm_out, textvariable=self.clip_max, width=10).grid(row=1, column=2, padx=5)

        # === Frame Normalización / Frecuencias ===
        frm_norm = ttk.LabelFrame(self, text="Normalización y Frecuencias")
        frm_norm.place(x=20, y=525, width=1040, height=120)

        ttk.Label(frm_norm, text="Estadístico (PPV y Freq):").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        cb_ns = ttk.Combobox(frm_norm, textvariable=self.norm_stat, width=16, state="readonly",
                             values=["Median", "Mean"])
        cb_ns.grid(row=0, column=1, padx=5)

        ttk.Label(frm_norm, text="Frecuencias:").grid(row=0, column=2, sticky="e", padx=5)
        cb_fa = ttk.Combobox(frm_norm, textvariable=self.freq_action, width=24, state="readonly",
                             values=["None", "Normalize by statistic"])
        cb_fa.grid(row=0, column=3, padx=5)

        ttk.Checkbutton(frm_norm, text="Ajustar frecuencia por SD (experimental)",
                        variable=self.freq_adjust_by_sd).grid(row=1, column=1, columnspan=3, sticky="w", padx=5)

        # === Frame Vista previa ===
        frm_prev = ttk.LabelFrame(self, text="Vista previa (X = SD, Y configurable)")
        frm_prev.place(x=20, y=655, width=820, height=100)
        frm_prev.grid_columnconfigure(1, weight=1)

        ttk.Label(frm_prev, text="Columna Y:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.cb_ycol = ttk.Combobox(frm_prev, textvariable=self.preview_y_col, width=45, state="readonly", values=[])
        self.cb_ycol.grid(row=0, column=1, padx=5, sticky="we")

        ttk.Label(frm_prev, text="Serie:").grid(row=0, column=2, sticky="e", padx=5)
        ttk.Combobox(frm_prev, textvariable=self.preview_series_mode, width=18, state="readonly",
                     values=["Original", "Ajustada@SD_ref", "Normalizada"]).grid(row=0, column=3, padx=5)

        ttk.Checkbutton(frm_prev, text="Log X", variable=self.preview_logx).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Checkbutton(frm_prev, text="Log Y", variable=self.preview_logy).grid(row=1, column=3, sticky="w", padx=5)

        ttk.Button(self, text="👁 Ver gráfica", command=self.show_preview)\
            .place(x=860, y=670, width=200, height=36)

        # === Frame Salida / Ejecutar ===
        frm_extras = ttk.LabelFrame(self, text="Salida")
        frm_extras.place(x=20, y=760, width=820, height=60)

        ttk.Checkbutton(frm_extras, text="Exportar columna Ajustada@SD_ref", variable=self.export_adjusted)\
            .grid(row=0, column=0, padx=5, sticky="w")
        ttk.Checkbutton(frm_extras, text="Exportar hoja Parámetros", variable=self.export_params)\
            .grid(row=0, column=1, padx=5, sticky="w")
        ttk.Checkbutton(frm_extras, text="Exportar hoja QA", variable=self.export_QA)\
            .grid(row=0, column=2, padx=5, sticky="w")

        ttk.Button(self, text="▶ Ejecutar normalización", command=self.run)\
            .place(x=860, y=810, width=200, height=40)

    # === Handlers de archivo
    def select_input(self):
        p = filedialog.askopenfilename(
            title="Selecciona el archivo XLSX",
            filetypes=[("Excel", "*.xlsx")]
        )
        if p:
            self.file_in.set(p)
            base = os.path.splitext(os.path.basename(p))[0]
            out = os.path.join(os.path.dirname(p), f"{base}_normalizado.xlsx")
            self.file_out.set(out)

            try:
                # Carga y prepara df para poder poblar combos de columnas Y
                self.df = pd.read_excel(p, header=2)
                self.refresh_y_columns()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")
                self.df = None
                self.cb_ycol["values"] = []
                self.preview_y_col.set("")

    def select_output(self):
        p = filedialog.asksaveasfilename(
            title="Guardar como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if p:
            self.file_out.set(p)

    def refresh_y_columns(self):
        """Detecta columnas de PPV y Frecuencia para la vista previa."""
        if self.df is None:
            self.cb_ycol["values"] = []
            self.preview_y_col.set("")
            return
        all_cols = list(self.df.columns)
        ppv_cols = [c for c in all_cols if "mm/s" in c]
        freq_cols = [c for c in all_cols if ("Hz" in c or "HZ" in c)]
        y_candidates = ppv_cols + freq_cols
        self.cb_ycol["values"] = y_candidates
        if y_candidates:
            self.preview_y_col.set(y_candidates[0])
        else:
            self.preview_y_col.set("")

    # === Lógica común de cálculo (sin guardar) ===
    def _compute_sd_and_series(self, df, y_col):
        """Con los parámetros actuales de la GUI, devuelve:
           - sd (Serie): distancia escalada
           - y_series (Serie): según modo 'Original' / 'Ajustada@SD_ref' / 'Normalizada'
           - info dict: {n, r2, med_used, sd_ref, npoints, fallback}
        """
        # Columnas base
        COL_CARGA = self.col_carga.get().strip()
        COL_DIST = self.col_dist.get().strip()

        # Validaciones
        for must in [COL_CARGA, COL_DIST, y_col]:
            if must not in df.columns:
                raise ValueError(f"Columna faltante: {must}")

        # Calcular SD
        carga = to_numeric(df[COL_CARGA])
        dist = to_numeric(df[COL_DIST])
        sd = dist / (carga ** (1/3))
        SD_COL = "SD (Distancia Escalada)"

        # SD_ref
        sd_ref_mode = self.sdref_mode.get()
        if sd_ref_mode == "Median":
            sd_ref = choose_stat(sd, "Median")
        elif sd_ref_mode == "Mean":
            sd_ref = choose_stat(sd, "Mean")
        elif sd_ref_mode == "1":
            sd_ref = 1.0
        else:
            try:
                sd_ref = float(self.sdref_custom.get().strip())
            except:
                sd_ref = np.nan
        if pd.isna(sd_ref) or sd_ref <= 0:
            raise ValueError("SD_ref inválido.")

        # Parámetros globales
        reg_mode = self.reg_mode.get()
        try:
            n_fixed = float(self.fixed_n.get())
        except:
            n_fixed = 1.6
        try:
            min_points = int(self.min_points.get())
        except:
            min_points = 3
        log_base = self.log_base.get()

        out_mode = self.outlier_mode.get()
        try:
            iqr_k = float(self.iqr_k.get())
        except:
            iqr_k = 1.5
        try:
            z_thr = float(self.z_thr.get())
        except:
            z_thr = 3.0

        vmin = None if not self.clip_min.get().strip() else float(self.clip_min.get())
        vmax = None if not self.clip_max.get().strip() else float(self.clip_max.get())

        norm_stat = self.norm_stat.get()

        z = to_numeric(df[y_col])

        # Filtro outliers (para estimar n)
        tmp_df = df.copy()
        tmp_df[SD_COL] = sd
        tmp_df[y_col] = z
        mask = apply_outlier_filter(tmp_df, y_col, SD_COL, mode=out_mode, param=iqr_k, zthr=z_thr)

        # Estimar n si es PPV (si es frecuencia, igual seguimos misma lógica para estabilidad)
        if reg_mode.startswith("OLS"):
            est = estimate_n_loglog(sd[mask], z[mask], min_points=min_points, base=log_base)
            if not est["ok"]:
                n_use = n_fixed
                r2 = np.nan
                fallback = True
                npoints = est["n_points"]
            else:
                n_use = est["n"]; r2 = est["r2"]; fallback = False; npoints = est["n_points"]
        else:
            n_use = n_fixed
            r2 = np.nan
            fallback = True
            npoints = int(mask.sum())

        # Construir serie según modo
        mode = self.preview_series_mode.get()
        with np.errstate(invalid='ignore'):
            factor = (sd / sd_ref) ** n_use
        factor = clip_if_needed(factor, vmin, vmax)

        if mode == "Original":
            y_series = z
            med_used = np.nan
        elif mode == "Ajustada@SD_ref":
            y_series = z * factor
            med_used = np.nan
        else:  # Normalizada
            z_adj = z * factor
            med_used = choose_stat(z_adj, mode=norm_stat)
            y_series = z_adj / med_used if (not pd.isna(med_used) and med_used != 0) else np.nan

        info = dict(n=n_use, r2=r2, med_used=med_used, sd_ref=sd_ref, npoints=npoints, fallback=fallback)
        return sd, y_series, info

    def show_preview(self):
        """Genera una ventana con la gráfica SD (X) vs serie elegida (Y)."""
        try:
            if self.df is None:
                # Si aún no se ha cargado, intenta cargar ahora
                in_path = self.file_in.get().strip()
                if not in_path or not os.path.exists(in_path):
                    messagebox.showerror("Error", "Selecciona primero un archivo de entrada.")
                    return
                self.df = pd.read_excel(in_path, header=2)
                self.refresh_y_columns()

            y_col = self.preview_y_col.get().strip()
            if not y_col:
                messagebox.showerror("Error", "Selecciona la columna Y para la vista previa.")
                return

            sd, y_series, info = self._compute_sd_and_series(self.df.copy(), y_col)

            # Preparar datos válidos
            x = to_numeric(sd).to_numpy()
            y = to_numeric(y_series).to_numpy()
            m = np.isfinite(x) & np.isfinite(y)
            x = x[m]; y = y[m]

            if x.size == 0:
                messagebox.showwarning("Sin datos", "No hay puntos válidos para graficar con la configuración actual.")
                return

            # Crear ventana Toplevel con la figura
            win = tk.Toplevel(self)
            win.title(f"Vista previa — SD vs {y_col} [{self.preview_series_mode.get()}]")
            win.geometry("900x650")

            fig = Figure(figsize=(8.5, 6.0), dpi=100)
            ax = fig.add_subplot(111)
            ax.scatter(x, y, s=16, alpha=0.7)

            # Ejes y escala
            ax.set_xlabel("SD (Distancia Escalada)")
            ax.set_ylabel(f"{y_col}  ({self.preview_series_mode.get()})")

            if self.preview_logx.get():
                ax.set_xscale("log")
            if self.preview_logy.get():
                ax.set_yscale("log")

            # Título con info clave
            n_str = f"n={info['n']:.3f}" if np.isfinite(info['n']) else "n=NA"
            r2_str = f"R²={info['r2']:.3f}" if np.isfinite(info['r2']) else "R²=NA"
            sdref_str = f"SD_ref={info['sd_ref']:.3g}" if np.isfinite(info['sd_ref']) else "SD_ref=NA"
            extra = " (fallback)" if info.get("fallback") else ""
            ax.set_title(f"SD vs {y_col} — {self.preview_series_mode.get()} | {n_str}, {r2_str}, {sdref_str}{extra}")

            ax.grid(True, which="both", alpha=0.25)

            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except Exception as e:
            messagebox.showerror("Error en vista previa", str(e))

    # === Lógica de cálculo y guardado (igual que antes) ===
    def _get_sd_ref(self, sd_series):
        mode = self.sdref_mode.get()
        if mode == "Median":
            return choose_stat(sd_series, "Median")
        elif mode == "Mean":
            return choose_stat(sd_series, "Mean")
        elif mode == "1":
            return 1.0
        else:
            try:
                v = float(self.sdref_custom.get().strip())
                return v if v > 0 else np.nan
            except:
                return np.nan

    def run(self):
        try:
            in_path = self.file_in.get().strip()
            out_path = self.file_out.get().strip()
            if not in_path or not os.path.exists(in_path):
                messagebox.showerror("Error", "Selecciona un archivo de entrada válido.")
                return
            if not out_path:
                messagebox.showerror("Error", "Indica una ruta de salida.")
                return

            # Leer Excel (headers en fila 3)
            df = pd.read_excel(in_path, header=2)
            self.df = df.copy()  # actualizar cache por si cambió
            self.refresh_y_columns()

            # Columnas
            COL_X = self.col_x.get().strip()
            COL_Y = self.col_y.get().strip()
            COL_CARGA = self.col_carga.get().strip()
            COL_DIST = self.col_dist.get().strip()

            # Validar columnas mínimas
            faltantes = [c for c in [COL_CARGA, COL_DIST] if c not in df.columns]
            if faltantes:
                messagebox.showerror("Error", f"Faltan columnas requeridas: {faltantes}")
                return

            # Detectar componentes
            ALL_COLS = list(df.columns)
            ppv_cols = [c for c in ALL_COLS if "mm/s" in c]
            freq_cols = [c for c in ALL_COLS if ("Hz" in c or "HZ" in c)]

            # SD
            carga = to_numeric(df[COL_CARGA])
            dist = to_numeric(df[COL_DIST])
            SD_COL = "SD (Distancia Escalada)"
            df[SD_COL] = dist / (carga ** (1/3))

            # SD_ref
            sd_ref = self._get_sd_ref(df[SD_COL])
            if pd.isna(sd_ref) or sd_ref <= 0:
                messagebox.showerror("Error", "SD_ref inválido. Revisa el modo/valor.")
                return

            # Parámetros globales
            reg_mode = self.reg_mode.get()
            try:
                n_fixed = float(self.fixed_n.get())
            except:
                n_fixed = 1.6
            try:
                min_points = int(self.min_points.get())
            except:
                min_points = 3
            log_base = self.log_base.get()

            out_mode = self.outlier_mode.get()
            try:
                iqr_k = float(self.iqr_k.get())
            except:
                iqr_k = 1.5
            try:
                z_thr = float(self.z_thr.get())
            except:
                z_thr = 3.0

            vmin = None if not self.clip_min.get().strip() else float(self.clip_min.get())
            vmax = None if not self.clip_max.get().strip() else float(self.clip_max.get())

            norm_stat = self.norm_stat.get()
            freq_action = self.freq_action.get()
            freq_by_sd = self.freq_adjust_by_sd.get()

            export_adjusted = self.export_adjusted.get()
            export_params = self.export_params.get()
            export_QA = self.export_QA.get()

            # --- Resultado
            result = pd.DataFrame()
            if COL_X in df.columns: result[COL_X] = df[COL_X]
            if COL_Y in df.columns: result[COL_Y] = df[COL_Y]
            result[SD_COL] = df[SD_COL]

            params_rows = []
            qa_rows = []

            # Registrar configuración
            cfg = {
                "SD_ref_mode": self.sdref_mode.get(),
                "SD_ref_value": sd_ref,
                "Reg_mode": reg_mode,
                "n_fixed": n_fixed,
                "min_points": min_points,
                "log_base": log_base,
                "Outlier_mode": out_mode,
                "IQR_k": iqr_k,
                "Z_thr": z_thr,
                "Clip_min": vmin,
                "Clip_max": vmax,
                "Norm_stat": norm_stat,
                "Freq_action": freq_action,
                "Freq_adjust_by_SD": freq_by_sd
            }
            params_rows.append({"Parametro": "SD_ref", "Valor": sd_ref})
            for k,v in cfg.items():
                params_rows.append({"Parametro": f"CFG::{k}", "Valor": v})

            # --- Procesar PPV
            for col in ppv_cols:
                z = to_numeric(df[col])
                sd = to_numeric(df[SD_COL])

                # Filtro outliers
                mask = apply_outlier_filter(df.assign(**{col: z}), col, SD_COL,
                                            mode=out_mode, param=iqr_k, zthr=z_thr)
                # Estimar n
                if reg_mode.startswith("OLS"):
                    est = estimate_n_loglog(sd[mask], z[mask], min_points=min_points, base=log_base)
                    if not est["ok"]:
                        n_use = n_fixed
                        r2, k, a, b, npoints = np.nan, np.nan, np.nan, np.nan, est["n_points"]
                        fallback = True
                    else:
                        n_use = est["n"]; r2 = est["r2"]; k = est["k"]; a = est["a"]; b = est["b"]
                        npoints = est["n_points"]
                        fallback = False
                else:
                    # n fijo
                    n_use = n_fixed
                    r2 = k = a = b = np.nan
                    npoints = int(mask.sum())
                    fallback = True

                # Factor y ajuste
                with np.errstate(invalid='ignore'):
                    factor = (sd / sd_ref) ** n_use
                factor = clip_if_needed(factor, vmin, vmax)
                z_adj = z * factor

                # Normalización por estadístico elegido
                med_used = choose_stat(z_adj, mode=norm_stat)
                z_norm = z_adj / med_used if (not pd.isna(med_used) and med_used != 0) else np.nan

                # Guardar
                col_norm = f"{col} (Norm)"
                result[col_norm] = z_norm

                if export_adjusted:
                    col_adj = f"{col} (Ajustada@SD_ref)"
                    result[col_adj] = z_adj

                params_rows.append({
                    "Parametro": f"n ({col})",
                    "Valor": n_use,
                    "R2_regresion": r2,
                    "k({base})".format(base=("e" if log_base=="e" else "10")): k,
                    "Pendiente_b": -n_use,
                    "Estadistico_norm({})".format(col): norm_stat,
                    "Valor_estadistico({})".format(col): med_used,
                    "Puntos_validos": npoints,
                    "Fallback_n_fijo_o_insuficiente": "Sí" if fallback else "No"
                })

                # QA percentiles
                zn = result[col_norm].to_numpy(dtype=float)
                zn_ok = zn[np.isfinite(zn)]
                p5  = np.nan if zn_ok.size == 0 else np.nanpercentile(zn_ok, 5)
                p50 = np.nan if zn_ok.size == 0 else np.nanpercentile(zn_ok, 50)
                p95 = np.nan if zn_ok.size == 0 else np.nanpercentile(zn_ok, 95)

                qa_rows.append({
                    "Columna": col,
                    "N_total": int(z.notna().sum()),
                    "N_usados_en_fit": int(npoints),
                    "P5_norm": p5,
                    "P50_norm": p50,
                    "P95_norm": p95,
                })

            # --- Frecuencias
            for col in freq_cols:
                fvals = to_numeric(df[col])
                if freq_by_sd:
                    with np.errstate(invalid='ignore'):
                        factor = (to_numeric(df[SD_COL]) / sd_ref) ** n_fixed
                    factor = clip_if_needed(factor, vmin, vmax)
                    f_adj = fvals * factor
                    base_series = f_adj
                else:
                    base_series = fvals

                if freq_action.startswith("Normalize"):
                    stat_used = choose_stat(base_series, mode=norm_stat)
                    f_norm = base_series / stat_used if (not pd.isna(stat_used) and stat_used != 0) else np.nan
                    result[f"{col} (Norm)"] = f_norm
                    params_rows.append({
                        "Parametro": f"Estadistico_norm ({col})",
                        "Valor": norm_stat
                    })

            # --- Exportar
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                # Ordenar columnas: X, Y, SD, Ajustadas (si hay), Norm
                base_cols = [c for c in [COL_X, COL_Y, SD_COL] if c in df.columns]
                adj_cols = [c for c in result.columns if c.endswith("(Ajustada@SD_ref)")]
                norm_cols = [c for c in result.columns if c.endswith("(Norm)")]
                cols_order = base_cols + adj_cols + norm_cols
                result[cols_order].to_excel(writer, index=False, sheet_name="Normalizado")

                if self.export_params.get():
                    params_df = pd.DataFrame(params_rows)
                    params_df.to_excel(writer, index=False, sheet_name="Parametros")

                if self.export_QA.get():
                    qa_df = pd.DataFrame(qa_rows)
                    qa_df.to_excel(writer, index=False, sheet_name="QA")

            messagebox.showinfo("Listo", f"Archivo generado:\n{out_path}")

        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = NormalizationApp()
    app.mainloop()
