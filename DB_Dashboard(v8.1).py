# -*- coding: utf-8 -*-
"""
NOM-016 • Visor de gráficas (v8 + toolbar nativa)

Cambio solicitado:
- Se elimina el zoom personalizado y se activa la barra nativa de Matplotlib (NavigationToolbar2Tk)
  en TODAS las gráficas (zoom, pan, home, guardar, etc.).

Se conserva TODO lo anterior:
- Pestaña Por Localidad con Barras/Serie/Heatmap/Pie, filtros de localidad y métrica,
  exportación de tabla/resumen, ver tabla.
- Tooltips/hover en barras, líneas y heatmap (leyenda inferior).
- Altura dinámica de figuras para aprovechar el área.
- Pestañas Por día, Por persona, Ubicaciones, Estadísticos, Desempeño.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# === Columnas esperadas en la hoja BASE ===
DATE_COL   = "FECHA"
PERSON_COL = "OPERADOR"
X_COL      = "COORD. X NOMI"
Y_COL      = "COORD. Y NOMI"
LOC_COL    = "LOCALIDAD"

# ---------- Utilidades de datos -----------
def parse_fecha(series: pd.Series) -> pd.Series:
    dt1 = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if dt1.dt.date.nunique() <= 1:
        s_num = pd.to_numeric(series, errors="coerce")
        dt2 = pd.to_datetime(s_num, unit="D", origin="1899-12-30", errors="coerce")
        return dt2
    return dt1

def safe_date(s: str):
    if not s or str(s).strip() == "":
        return None
    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None

def filter_by_dates(df: pd.DataFrame, start, end) -> pd.DataFrame:
    out = df.copy()
    if start is not None:
        out = out[out["__FECHA"] >= start]
    if end is not None:
        out = out[out["__FECHA"] <= end]
    return out

def compute_events_per_person(df: pd.DataFrame, ascending=False) -> pd.DataFrame:
    if PERSON_COL not in df.columns:
        return pd.DataFrame(columns=[PERSON_COL, "Monitoreos"])
    tmp = (
        df.dropna(subset=[PERSON_COL])
          .groupby(PERSON_COL)
          .size()
          .reset_index(name="Monitoreos")
          .sort_values("Monitoreos", ascending=ascending)
    )
    return tmp

def compute_locations_per_person(df: pd.DataFrame, round_decimals:int=3, ascending=False) -> pd.DataFrame:
    if PERSON_COL not in df.columns or X_COL not in df.columns or Y_COL not in df.columns:
        return pd.DataFrame(columns=[PERSON_COL, "Ubicaciones únicas"])

    d2 = df.dropna(subset=[PERSON_COL, X_COL, Y_COL]).copy()
    d2[X_COL] = pd.to_numeric(d2[X_COL], errors="coerce")
    d2[Y_COL] = pd.to_numeric(d2[Y_COL], errors="coerce")
    d2 = d2.dropna(subset=[X_COL, Y_COL])

    d2["Xr"] = d2[X_COL].round(round_decimals)
    d2["Yr"] = d2[Y_COL].round(round_decimals)
    d2 = d2.drop_duplicates(subset=[PERSON_COL, "Xr", "Yr"])

    tmp = (
        d2.groupby(PERSON_COL)
          .size()
          .reset_index(name="Ubicaciones únicas")
          .sort_values("Ubicaciones únicas", ascending=ascending)
    )
    return tmp

# === Localidades: helpers ===
def list_localidades(df: pd.DataFrame) -> list:
    if LOC_COL not in df.columns:
        return []
    vals = df[LOC_COL].dropna().astype(str).str.strip()
    return sorted(v for v in vals.unique() if v)

def compute_summary_by_localidad(df: pd.DataFrame, round_decimals:int=3) -> pd.DataFrame:
    if LOC_COL not in df.columns:
        return pd.DataFrame(columns=[LOC_COL,"Visitas","Operadores únicos","Primera visita","Última visita","Días únicos","Ubicaciones únicas"])
    d = df.dropna(subset=[LOC_COL]).copy()
    d[LOC_COL] = d[LOC_COL].astype(str).str.strip()
    g = d.groupby(LOC_COL)

    locs = d.dropna(subset=[X_COL, Y_COL]).copy()
    if not locs.empty and X_COL in d.columns and Y_COL in d.columns:
        locs[X_COL] = pd.to_numeric(locs[X_COL], errors="coerce")
        locs[Y_COL] = pd.to_numeric(locs[Y_COL], errors="coerce")
        locs = locs.dropna(subset=[X_COL, Y_COL])
        locs["Xr"] = locs[X_COL].round(round_decimals)
        locs["Yr"] = locs[Y_COL].round(round_decimals)
        locs_u = locs.drop_duplicates(subset=[LOC_COL,"Xr","Yr"]).groupby(LOC_COL).size().rename("Ubicaciones únicas")
    else:
        locs_u = pd.Series(dtype="int64", name="Ubicaciones únicas")

    out = pd.DataFrame({
        "Visitas": g.size(),
        "Operadores únicos": g[PERSON_COL].nunique(),
        "Primera visita": g["__FECHA"].min(),
        "Última visita": g["__FECHA"].max(),
        "Días únicos": g["__FECHA"].nunique(),
    })
    if not locs_u.empty:
        out = out.join(locs_u, how="left")
    else:
        out["Ubicaciones únicas"] = 0
    out = out.reset_index().rename(columns={LOC_COL:"Localidad"})
    return out.sort_values("Visitas", ascending=False)

def compute_metric_by_localidad(df: pd.DataFrame, metric:str, round_decimals:int=3) -> pd.DataFrame:
    """metric: 'Monitoreos' (conteo de filas) o 'Ubicaciones únicas' (Xr,Yr únicas por localidad)."""
    if LOC_COL not in df.columns:
        return pd.DataFrame(columns=["Localidad", metric])
    d = df.dropna(subset=[LOC_COL]).copy()
    d["_loc"] = d[LOC_COL].astype(str).str.strip()
    if metric == "Monitoreos":
        g = d.groupby("_loc").size().reset_index(name="valor")
    else:
        if X_COL not in d.columns or Y_COL not in d.columns:
            return pd.DataFrame(columns=["Localidad", metric])
        d[X_COL] = pd.to_numeric(d[X_COL], errors="coerce")
        d[Y_COL] = pd.to_numeric(d[Y_COL], errors="coerce")
        d = d.dropna(subset=[X_COL, Y_COL])
        d["Xr"] = d[X_COL].round(round_decimals)
        d["Yr"] = d[Y_COL].round(round_decimals)
        d = d.drop_duplicates(subset=["_loc","Xr","Yr"])
        g = d.groupby("_loc").size().reset_index(name="valor")
    g = g.rename(columns={"_loc":"Localidad"})
    return g.sort_values("valor", ascending=False)

def series_metric_localidad(df: pd.DataFrame, metric:str, group:str="Día", round_decimals:int=3) -> pd.DataFrame:
    d = df.dropna(subset=["__FECHA"]).copy()
    if d.empty:
        return pd.DataFrame(columns=["_idx","valor","_lab"])
    d["_idx"] = pd.to_datetime(d["__FECHA"])
    if metric == "Monitoreos":
        if group == "Semana":
            g = d.groupby(pd.Grouper(key="_idx", freq="W-MON")).size().reset_index(name="valor")
        elif group == "Mes":
            g = d.groupby(pd.Grouper(key="_idx", freq="MS")).size().reset_index(name="valor")
        else:
            g = d.groupby("_idx").size().reset_index(name="valor")
    else:
        if X_COL not in d.columns or Y_COL not in d.columns:
            return pd.DataFrame(columns=["_idx","valor","_lab"])
        d[X_COL] = pd.to_numeric(d[X_COL], errors="coerce")
        d[Y_COL] = pd.to_numeric(d[Y_COL], errors="coerce")
        d = d.dropna(subset=[X_COL, Y_COL])
        d["Xr"] = d[X_COL].round(round_decimals)
        d["Yr"] = d[Y_COL].round(round_decimals)
        if group == "Semana":
            d["PER"] = d["_idx"].dt.to_period("W-MON")
        elif group == "Mes":
            d["PER"] = d["_idx"].dt.to_period("M")
        else:
            d["PER"] = d["_idx"].dt.to_period("D")
        d = d.drop_duplicates(subset=["PER","Xr","Yr"])
        g = d.groupby("PER").size().reset_index(name="valor")
        idx = g["PER"].dt.start_time
        g["_idx"] = pd.to_datetime(idx)

    def _lab(idx_series):
        if group == "Semana":
            return idx_series.dt.strftime("Sem %U (%Y-%m-%d)")
        elif group == "Mes":
            return idx_series.dt.strftime("%Y-%m")
        else:
            return idx_series.dt.strftime("%Y-%m-%d")
    g["_lab"] = _lab(g["_idx"])
    return g[["_idx","valor","_lab"]].sort_values("_idx")

# ---------- UI -----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NOM-016 • Visor de gráficas v8")
        self.geometry("1380x920")
        self.minsize(1180, 760)
        self.configure(bg="#0f172a")

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#0f172a")
        style.configure("Card.TFrame", background="#111827")
        style.configure("TLabel", background="#0f172a", foreground="#e5e7eb", font=("Segoe UI", 10))
        style.configure("Title.TLabel", background="#0f172a", foreground="#f8fafc", font=("Segoe UI", 16, "bold"))
        style.configure("Sub.TLabel", background="#111827", foreground="#e5e7eb", font=("Segoe UI", 11, "bold"))
        style.configure("TButton", padding=8)
        style.configure("TCheckbutton", background="#111827", foreground="#e5e7eb")
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)

        self.df = None
        self.out_dir = None
        self.min_date = None
        self.max_date = None

        header = ttk.Frame(self, style="TFrame")
        header.pack(fill="x", padx=16, pady=(16, 8))
        ttk.Label(header, text="PANEL ESTADÍSTICO DE MONITOREO", style="Title.TLabel").pack(side="left")

        g = ttk.Frame(self, style="Card.TFrame")
        g.pack(fill="x", padx=16, pady=8)
        ttk.Button(g, text="📂 Cargar XLSX", command=self.on_open).grid(row=0, column=0, padx=8, pady=8, sticky="w")
        self.lbl_src = ttk.Label(g, text="Sin archivo...", style="TLabel")
        self.lbl_src.grid(row=0, column=1, padx=8, pady=8, sticky="w")
        ttk.Button(g, text="🖴 Elegir carpeta de guardado", command=self.on_choose_out).grid(row=0, column=2, padx=8, pady=8, sticky="w")
        self.lbl_out = ttk.Label(g, text="Sin carpeta...", style="TLabel")
        self.lbl_out.grid(row=0, column=3, padx=8, pady=8, sticky="w")
        g.grid_columnconfigure(1, weight=1)

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=16, pady=8)
        self.tab_day = ChartTabByDay(nb, title="Por período (monitoreos / ubicaciones)")
        self.tab_person = ChartTabByPerson(nb, title="Monitoreos por persona")
        self.tab_locs = ChartTabLocations(nb, title="Ubicaciones únicas por persona")
        self.tab_stats = StatsTab(nb, title="Estadísticos")
        self.tab_perf = PerfTab(nb, title="Desempeño por operador")
        self.tab_localidad = LocalidadTab(nb, title="Por Localidad")

        nb.add(self.tab_day, text="📅 Por día")
        nb.add(self.tab_person, text="👤 Por persona")
        nb.add(self.tab_locs, text="📍 Ubicaciones")
        nb.add(self.tab_stats, text="📊 Estadísticos")
        nb.add(self.tab_perf, text="📈 Desempeño")
        nb.add(self.tab_localidad, text="📍 Por Localidad")

    def on_open(self):
        path = filedialog.askopenfilename(
            title="Selecciona el XLSX (hoja BASE)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
        if not path: return
        try:
            df = None
            last_err = None
            for hdr in (2, 3, 1, 0):
                try:
                    tmp = pd.read_excel(path, sheet_name="BASE", header=hdr, engine="openpyxl")
                    if DATE_COL in tmp.columns and PERSON_COL in tmp.columns:
                        df = tmp
                        break
                except Exception as e:
                    last_err = e
            if df is None:
                if last_err:
                    raise last_err
                else:
                    raise RuntimeError("No pude detectar los encabezados de la hoja BASE. Revisa el archivo.")
        except Exception as e:
            messagebox.showerror("Error al leer XLSX", str(e)); return

        missing = [c for c in (DATE_COL, PERSON_COL) if c not in df.columns]
        if missing:
            messagebox.showerror("Columnas faltantes",
                                 f"No encuentro columnas requeridas: {missing}.\nRevisa la hoja 'BASE'.")
            return

        df[PERSON_COL] = df.get(PERSON_COL, pd.Series(dtype="object")).astype(str).str.strip()
        df["__FECHA"] = parse_fecha(df[DATE_COL]).dt.date

        self.df = df
        self.lbl_src.configure(text=os.path.basename(path))

        fechas_validas = df["__FECHA"].dropna()
        self.min_date = fechas_validas.min() if not fechas_validas.empty else None
        self.max_date = fechas_validas.max() if not fechas_validas.empty else None

        for tab in (self.tab_day, self.tab_person, self.tab_locs, self.tab_stats, self.tab_perf, self.tab_localidad):
            tab.set_data(self.df, self.min_date, self.max_date)

    def on_choose_out(self):
        d = filedialog.askdirectory(title="Elige carpeta de guardado")
        if not d: return
        self.out_dir = d
        for tab in (self.tab_day, self.tab_person, self.tab_locs, self.tab_stats, self.tab_perf, self.tab_localidad):
            tab.set_out_dir(d)

# ---------- Base Tabs -----------
class BaseTab(ttk.Frame):
    def __init__(self, parent, title: str):
        super().__init__(parent, style="Card.TFrame")
        self.df = None
        self.min_date = None
        self.max_date = None
        self.out_dir = None
        self.title_text = title
        self._hover_connections = []

        top = ttk.Frame(self, style="Card.TFrame"); top.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(top, text=title, style="Sub.TLabel").grid(row=0, column=0, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="Inicio (dd/mm/aaaa):").grid(row=0, column=1, sticky="e", padx=6)
        self.e_start = ttk.Entry(top, width=14); self.e_start.grid(row=0, column=2, sticky="w", padx=6)
        ttk.Label(top, text="Fin:").grid(row=0, column=3, sticky="e", padx=6)
        self.e_end = ttk.Entry(top, width=14); self.e_end.grid(row=0, column=4, sticky="w", padx=6)

        self.cmb_type = ttk.Combobox(top, state="readonly", width=22); self.cmb_type.grid(row=0, column=5, padx=6)
        ttk.Button(top, text="Actualizar", command=self.update_chart).grid(row=0, column=6, padx=6)
        ttk.Button(top, text="Todo", command=self.set_all).grid(row=0, column=7, padx=6)
        ttk.Button(top, text="Últimos 30", command=lambda: self.set_last_days(30)).grid(row=0, column=8, padx=6)
        ttk.Button(top, text="Últimos 14", command=lambda: self.set_last_days(14)).grid(row=0, column=9, padx=6)
        ttk.Button(top, text="Últimos 7", command=lambda: self.set_last_days(7)).grid(row=0, column=10, padx=6)

        dyn = ttk.Frame(self, style="Card.TFrame"); dyn.pack(fill="x", padx=12, pady=(0,4))
        self.var_grid = tk.BooleanVar(value=True)
        ttk.Checkbutton(dyn, text="Grid", variable=self.var_grid, command=self.update_chart).pack(side="left", padx=6)
        self.var_labels = tk.BooleanVar(value=False)
        ttk.Checkbutton(dyn, text="Etiquetas en barras", variable=self.var_labels, command=self.update_chart).pack(side="left", padx=6)
        self.var_logy = tk.BooleanVar(value=False)
        ttk.Checkbutton(dyn, text="Escala log(Y)", variable=self.var_logy, command=self.update_chart).pack(side="left", padx=6)

        # Cuerpo para canvas + toolbar
        self.body = ttk.Frame(self, style="Card.TFrame")
        self.body.pack(fill="both", expand=True, padx=12, pady=(4, 2))
        self.fig = None; self.ax = None; self.canvas = None
        self.ax_list = None
        self.toolbar = None  # <— toolbar nativa

        footer = ttk.Frame(self, style="Card.TFrame"); footer.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(footer, text="💾 Guardar imagen", command=self.save_png).pack(side="left", padx=6)
        self.lbl_status = ttk.Label(footer, text="", style="TLabel"); self.lbl_status.pack(side="right", padx=6)

    def _recreate_canvas(self, fig_w=10, fig_h=5, n_axes=1):
        # limita altura pero busca rellenar el cuerpo
        fig_h = min(fig_h, 7.8)
        # desconectar hovers previos
        if self.canvas is not None:
            try:
                for cid in self._hover_connections:
                    self.canvas.mpl_disconnect(cid)
            except Exception:
                pass
            self._hover_connections.clear()
            try:
                self.canvas.get_tk_widget().destroy()
            except Exception:
                pass
            self.canvas = None
        # destruir toolbar previa
        if getattr(self, "toolbar", None) is not None:
            try:
                self.toolbar.destroy()
            except Exception:
                pass
            self.toolbar = None

        self.fig = plt.Figure(figsize=(fig_w, fig_h), dpi=100)
        self.ax_list = []
        if n_axes == 1:
            self.ax = self.fig.add_subplot(111)
            self.ax_list = [self.ax]
        else:
            for i in range(n_axes):
                ax = self.fig.add_subplot(n_axes, 1, i+1)
                self.ax_list.append(ax)
            self.ax = self.ax_list[0]

        # canvas
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.body)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # toolbar nativa (zoom/pan/home/guardar). Siempre visible bajo el canvas.
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.body)
        self.toolbar.update()
        self.toolbar.pack(side="bottom", fill="x")

    def _apply_common_to(self, ax):
        ax.grid(self.var_grid.get(), which="both", axis="both", linestyle="--", linewidth=0.5, alpha=0.6)
        ax.set_yscale("log" if self.var_logy.get() else "linear")

    def _connect_hover(self, payload):
        def on_move(event):
            if event.inaxes is None:
                self.lbl_status.config(text=""); 
                return
            try:
                if payload.get("kind") == "bar":
                    best = None; dmin = 1e18
                    for p, lx, val in payload["bars"]:
                        ax = p.axes
                        if event.inaxes is not ax: 
                            continue
                        if p.get_width() > p.get_height():  # horizontal
                            cy = p.get_y() + p.get_height()/2.0
                            d = abs((event.ydata or 0) - cy)
                        else:
                            cx = p.get_x() + p.get_width()/2.0
                            d = abs((event.xdata or 0) - cx)
                        if d < dmin:
                            dmin = d; best = (lx, val)
                    if best:
                        self.lbl_status.config(text=f"{payload['title']} • {best[0]} = {best[1]}")
                elif payload.get("kind") == "heatmap":
                    ax = payload.get("ax", self.ax)
                    if event.inaxes is not ax or event.xdata is None or event.ydata is None:
                        self.lbl_status.config(text=""); return
                    xi = int(round(event.xdata)); yi = int(round(event.ydata))
                    if 0 <= yi < payload["data"].shape[0] and 0 <= xi < payload["data"].shape[1]:
                        val = int(payload["data"][yi, xi])
                        op  = payload["rows"][yi]
                        loc = payload["cols"][xi]
                        self.lbl_status.config(text=f"{payload['metric']} • {op} × {loc} = {val}")
                else:
                    xs = payload["x"]; ys = payload["y"]; ax = payload.get("ax", self.ax)
                    if event.inaxes is not ax or event.xdata is None:
                        self.lbl_status.config(text=""); return
                    idx = int(round(event.xdata))
                    if 0 <= idx < len(xs):
                        self.lbl_status.config(text=f"{payload['title']} • {xs[idx]} = {ys[idx]}")
            except Exception:
                pass
        if self.canvas:
            cid = self.canvas.mpl_connect("motion_notify_event", on_move)
            self._hover_connections.append(cid)

    def set_data(self, df, dmin, dmax):
        self.df = df; self.min_date = dmin; self.max_date = dmax
        if dmin and dmax:
            self.e_start.delete(0, tk.END); self.e_start.insert(0, dmin.strftime("%d/%m/%Y"))
            self.e_end.delete(0, tk.END); self.e_end.insert(0, dmax.strftime("%d/%m/%Y"))
        self.update_chart()

    def set_out_dir(self, d): self.out_dir = d
    def get_date_range(self): return safe_date(self.e_start.get().strip()), safe_date(self.e_end.get().strip())
    def set_all(self):
        if self.min_date and self.max_date:
            self.e_start.delete(0, tk.END); self.e_start.insert(0, self.min_date.strftime("%d/%m/%Y"))
            self.e_end.delete(0, tk.END); self.e_end.insert(0, self.max_date.strftime("%d/%m/%Y"))
            self.update_chart()
    def set_last_days(self, n):
        if self.max_date:
            start = pd.to_datetime(self.max_date) - pd.Timedelta(days=n-1)
            self.e_start.delete(0, tk.END); self.e_start.insert(0, start.strftime("%d/%m/%Y"))
            self.e_end.delete(0, tk.END); self.e_end.insert(0, self.max_date.strftime("%d/%m/%Y"))
            self.update_chart()
    def save_png(self):
        if self.fig is None: return
        if self.out_dir is None:
            messagebox.showwarning("Guardar", "Primero elige una carpeta de guardado (arriba)."); return
        path = os.path.join(self.out_dir, self.get_default_png_name())
        try:
            self.fig.tight_layout(rect=[0.02,0.05,0.98,0.95]); 
            self.fig.savefig(path, dpi=150, bbox_inches="tight")
            self.lbl_status.config(text=f"Guardado: {path}")
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))

    def get_default_png_name(self): ...
    def update_chart(self): ...

# ---------- Tab: Por día -----------
class ChartTabByDay(BaseTab):
    def __init__(self, parent, title):
        super().__init__(parent, title)
        extras = ttk.Frame(self, style="Card.TFrame"); extras.pack(fill="x", padx=12, pady=(0,4))
        self.var_trend = tk.BooleanVar(value=False)
        ttk.Checkbutton(extras, text="Línea de tendencia", variable=self.var_trend, command=self.update_chart).pack(side="left", padx=6)
        self.var_ma = tk.BooleanVar(value=False)
        ttk.Checkbutton(extras, text="Media móvil (7 días)", variable=self.var_ma, command=self.update_chart).pack(side="left", padx=6)

        ttk.Label(extras, text="Agrupar por:").pack(side="left", padx=(18,6))
        self.cmb_group = ttk.Combobox(extras, state="readonly", width=10, values=("Día","Semana","Mes"))
        self.cmb_group.current(0); self.cmb_group.bind("<<ComboboxSelected>>", lambda e: self.update_chart()); self.cmb_group.pack(side="left")

        ttk.Label(extras, text="Métrica:").pack(side="left", padx=(18,6))
        self.cmb_metric = ttk.Combobox(extras, state="readonly", width=20, values=("Monitoreos","Ubicaciones únicas"))
        self.cmb_metric.current(0); self.cmb_metric.bind("<<ComboboxSelected>>", lambda e: self.update_chart()); self.cmb_metric.pack(side="left")

        ttk.Label(extras, text="Redondeo XY:").pack(side="left", padx=(18,6))
        self.decimals = tk.IntVar(value=3)
        ttk.Spinbox(extras, from_=0, to=8, textvariable=self.decimals, width=5, command=self.update_chart).pack(side="left", padx=6)

        self.cmb_type["values"] = ("Barras", "Línea")
        self.cmb_type.current(0)

    def get_default_png_name(self): 
        return "ubicaciones_por_periodo.png" if self.cmb_metric.get()=="Ubicaciones únicas" else "monitoreos_por_periodo.png"

    def _group_label(self, idx_series):
        if self.cmb_group.get() == "Semana":
            return idx_series.dt.strftime("Sem %U (%Y-%m-%d)")
        elif self.cmb_group.get() == "Mes":
            return idx_series.dt.strftime("%Y-%m")
        else:
            return idx_series.dt.strftime("%Y-%m-%d")

    def _series_monitoreos(self, d):
        d = d.dropna(subset=["__FECHA"]).copy()
        idx = pd.to_datetime(d["__FECHA"])
        d["_idx"] = idx
        if self.cmb_group.get() == "Semana":
            g = d.groupby(pd.Grouper(key="_idx", freq="W-MON")).size().reset_index(name="valor")
        elif self.cmb_group.get() == "Mes":
            g = d.groupby(pd.Grouper(key="_idx", freq="MS")).size().reset_index(name="valor")
        else:
            g = d.groupby("_idx").size().reset_index(name="valor")
        g["_lab"] = self._group_label(g["_idx"])
        return g

    def _series_ubicaciones(self, d, decimals=3):
        d2 = d.dropna(subset=[X_COL, Y_COL, "__FECHA"]).copy()
        if d2.empty: 
            return pd.DataFrame(columns=["_idx","valor","_lab"])
        d2[X_COL] = pd.to_numeric(d2[X_COL], errors="coerce")
        d2[Y_COL] = pd.to_numeric(d2[Y_COL], errors="coerce")
        d2 = d2.dropna(subset=[X_COL, Y_COL])
        d2["Xr"] = d2[X_COL].round(decimals)
        d2["Yr"] = d2[Y_COL].round(decimals)
        d2["_idx"] = pd.to_datetime(d2["__FECHA"])

        if self.cmb_group.get() == "Semana":
            d2["PER"] = d2["_idx"].dt.to_period("W-MON")
        elif self.cmb_group.get() == "Mes":
            d2["PER"] = d2["_idx"].dt.to_period("M")
        else:
            d2["PER"] = d2["_idx"].dt.to_period("D")

        d2 = d2.drop_duplicates(subset=["PER","Xr","Yr"])
        g = d2.groupby("PER").size().reset_index(name="valor")
        idx = g["PER"].dt.start_time
        g["_idx"] = pd.to_datetime(idx)
        g["_lab"] = self._group_label(g["_idx"])
        return g[["_idx","valor","_lab"]].sort_values("_idx")

    def update_chart(self):
        if self.df is None: return
        start, end = self.get_date_range()
        dff = filter_by_dates(self.df, start, end)

        # altura dinámica: más puntos => más alto (dentro de topes)
        est_n = max(10, min(120, len(dff.dropna(subset=["__FECHA"])["__FECHA"].unique())))
        fig_h = min(7.6, max(5.5, 0.06 * est_n))
        self._recreate_canvas(fig_w=11.8, fig_h=fig_h)

        if dff.empty:
            self.ax.set_title("Sin datos en el rango seleccionado"); self.canvas.draw(); return

        metric = self.cmb_metric.get()
        if metric == "Ubicaciones únicas":
            g = self._series_ubicaciones(dff, decimals=int(self.decimals.get()))
            ylabel = "Ubicaciones únicas"
            title = "Ubicaciones únicas por período"
        else:
            g = self._series_monitoreos(dff)
            ylabel = "Cantidad de monitoreos"
            title = "Monitoreos por período"

        if g.empty:
            self.ax.set_title("Sin datos en el rango seleccionado"); self.canvas.draw(); return

        xlabels = g["_lab"].tolist()
        y = g["valor"].values

        if self.cmb_type.get() == "Línea":
            self.ax.plot(np.arange(len(y)), y, marker="o")
            payload = {"kind":"line", "x": xlabels, "y": y.tolist(), "title":ylabel, "ax": self.ax}
        else:
            bars = self.ax.bar(np.arange(len(y)), y)
            if self.var_labels.get():
                for b, v in zip(bars, y):
                    self.ax.text(b.get_x()+b.get_width()/2, v, f"{int(v)}", va="bottom", ha="center", fontsize=8)
            payload = {"kind":"bar", "bars":[(b, xl, int(v)) for b, xl, v in zip(bars, xlabels, y)], "title":ylabel}

        self.ax.set_xticks(np.arange(len(xlabels)))
        self.ax.set_xticklabels(xlabels, rotation=45, ha="right")
        self.ax.set_xlabel("Período"); self.ax.set_ylabel(ylabel)
        self.ax.set_title(title)
        self._apply_common_to(self.ax)

        if self.var_ma.get() and len(y) >= 3:
            s = pd.Series(y)
            window = 7 if self.cmb_group.get()=="Día" else 3
            ma = s.rolling(window, min_periods=1).mean()
            self.ax.plot(np.arange(len(y)), ma.values, linestyle="--")
        if self.var_trend.get() and len(y) >= 2:
            t = np.arange(len(y))
            m, b = np.polyfit(t, y, 1)
            self.ax.plot(t, m*t + b, linestyle=":")

        self.fig.tight_layout(rect=[0.03,0.26,0.98,0.94])
        self.canvas.draw()
        self._connect_hover(payload)

# ---------- Tab: Por persona ----------
class ChartTabByPerson(BaseTab):
    def __init__(self, parent, title):
        super().__init__(parent, title)
        extras = ttk.Frame(self, style="Card.TFrame"); extras.pack(fill="x", padx=12, pady=(0,4))
        self.var_mean = tk.BooleanVar(value=True)
        self.var_median = tk.BooleanVar(value=False)
        ttk.Checkbutton(extras, text="Mostrar media", variable=self.var_mean, command=self.update_chart).pack(side="left", padx=6)
        ttk.Checkbutton(extras, text="Mostrar mediana", variable=self.var_median, command=self.update_chart).pack(side="left", padx=6)

        ttk.Label(extras, text="Orden:").pack(side="left", padx=(18,6))
        self.cmb_sort = ttk.Combobox(extras, state="readonly", width=12, values=("Descendente","Ascendente"))
        self.cmb_sort.current(0); self.cmb_sort.bind("<<ComboboxSelected>>", lambda e: self.update_chart()); self.cmb_sort.pack(side="left")

        self.cmb_type["values"] = ("Barras", "Barras horizontales")
        self.cmb_type.current(1)

    def get_default_png_name(self): return "monitoreos_por_persona.png"

    def update_chart(self):
        if self.df is None: return
        start, end = self.get_date_range()
        dff = filter_by_dates(self.df, start, end)
        asc = (self.cmb_sort.get() == "Ascendente")
        data = compute_events_per_person(dff, ascending=asc)

        n = max(1, len(data))
        fig_h = min(7.6, max(4.8, 0.42 * n))
        self._recreate_canvas(fig_w=11.5, fig_h=fig_h)

        if data.empty:
            self.ax.set_title("Sin datos en el rango seleccionado"); self.canvas.draw(); return

        names = data[PERSON_COL].astype(str).tolist()
        vals = data["Monitoreos"].astype(int).to_numpy()

        if self.cmb_type.get() == "Barras horizontales":
            y_pos = np.arange(len(names))
            bars = self.ax.barh(y_pos, vals)
            self.ax.set_yticks(y_pos)
            self.ax.set_yticklabels(names)
            self.ax.set_ylabel("Operador"); self.ax.set_xlabel("Cantidad de monitoreos")
            if self.var_labels.get():
                for yi, (b, v) in enumerate(zip(bars, vals)):
                    self.ax.text(v, yi, f" {int(v)}", va="center", ha="left")
            if self.var_mean.get(): self.ax.axvline(np.mean(vals), linestyle="--")
            if self.var_median.get(): self.ax.axvline(np.median(vals), linestyle=":")
            self.fig.subplots_adjust(left=0.34, right=0.96, top=0.93, bottom=0.08)
            payload = {"kind":"bar", "bars":[(b, n, int(v)) for b, n, v in zip(bars, names, vals)], "title":"Monitoreos"}
        else:
            x_pos = np.arange(len(names))
            bars = self.ax.bar(x_pos, vals)
            self.ax.set_xticks(x_pos)
            self.ax.set_xticklabels(names, rotation=45, ha="right")
            self.ax.set_xlabel("Operador"); self.ax.set_ylabel("Cantidad de monitoreos")
            if self.var_labels.get():
                for b, v in zip(bars, vals):
                    self.ax.text(b.get_x()+b.get_width()/2, v, f"{int(v)}", va="bottom", ha="center")
            if self.var_mean.get(): self.ax.axhline(np.mean(vals), linestyle="--")
            if self.var_median.get(): self.ax.axhline(np.median(vals), linestyle=":")
            self.fig.subplots_adjust(left=0.07, right=0.98, top=0.93, bottom=0.30)
            payload = {"kind":"bar", "bars":[(b, n, int(v)) for b, n, v in zip(bars, names, vals)], "title":"Monitoreos"}

        self.ax.set_title("Monitoreos por persona")
        self._apply_common_to(self.ax)
        self.fig.tight_layout()
        self.canvas.draw()
        self._connect_hover(payload)

# ---------- Tab: Ubicaciones por persona ----------
class ChartTabLocations(BaseTab):
    def __init__(self, parent, title):
        super().__init__(parent, title)
        extras = ttk.Frame(self, style="Card.TFrame"); extras.pack(fill="x", padx=12, pady=(0,4))

        ttk.Label(extras, text="Redondeo XY (decimales):").pack(side="left", padx=6)
        self.decimals = tk.IntVar(value=3)
        ttk.Spinbox(extras, from_=0, to=8, textvariable=self.decimals, width=5, command=self.update_chart).pack(side="left", padx=6)

        ttk.Label(extras, text="Orden:").pack(side="left", padx=(18,6))
        self.cmb_sort = ttk.Combobox(extras, state="readonly", width=12, values=("Descendente","Ascendente"))
        self.cmb_sort.current(0); self.cmb_sort.bind("<<ComboboxSelected>>", lambda e: self.update_chart()); self.cmb_sort.pack(side="left")

        self.cmb_type["values"] = ("Barras", "Barras horizontales")
        self.cmb_type.current(1)

    def get_default_png_name(self): return "ubicaciones_por_persona.png"

    def update_chart(self):
        if self.df is None: return
        start, end = self.get_date_range()
        dff = filter_by_dates(self.df, start, end)
        asc = (self.cmb_sort.get() == "Ascendente")
        data = compute_locations_per_person(dff, round_decimals=int(self.decimals.get()), ascending=asc)

        n = max(1, len(data))
        fig_h = min(7.6, max(4.8, 0.42 * n))
        self._recreate_canvas(fig_w=11.5, fig_h=fig_h)

        if data.empty:
            self.ax.set_title("Sin datos en el rango seleccionado"); self.canvas.draw(); return

        names = data[PERSON_COL].astype(str).tolist()
        vals = data["Ubicaciones únicas"].astype(int).to_numpy()

        if self.cmb_type.get() == "Barras horizontales":
            y_pos = np.arange(len(names))
            bars = self.ax.barh(y_pos, vals)
            self.ax.set_yticks(y_pos)
            self.ax.set_yticklabels(names)
            self.ax.set_ylabel("Operador"); self.ax.set_xlabel("Cantidad de ubicaciones únicas")
            if self.var_labels.get():
                for yi, (b, v) in enumerate(zip(bars, vals)):
                    self.ax.text(v, yi, f" {int(v)}", va="center", ha="left")
            self.fig.subplots_adjust(left=0.34, right=0.96, top=0.93, bottom=0.08)
            payload = {"kind":"bar", "bars":[(b, n, int(v)) for b, n, v in zip(bars, names, vals)], "title":"Ubicaciones únicas"}
        else:
            x_pos = np.arange(len(names))
            bars = self.ax.bar(x_pos, vals)
            self.ax.set_xticks(x_pos)
            self.ax.set_xticklabels(names, rotation=45, ha="right")
            self.ax.set_xlabel("Operador"); self.ax.set_ylabel("Cantidad de ubicaciones únicas")
            if self.var_labels.get():
                for b, v in zip(bars, vals):
                    self.ax.text(b.get_x()+b.get_width()/2, v, f"{int(v)}", va="bottom", ha="center")
            self.fig.subplots_adjust(left=0.07, right=0.98, top=0.93, bottom=0.30)
            payload = {"kind":"bar", "bars":[(b, n, int(v)) for b, n, v in zip(bars, names, vals)], "title":"Ubicaciones únicas"}

        self.ax.set_title("Ubicaciones únicas por persona")
        self._apply_common_to(self.ax)
        self.fig.tight_layout()
        self.canvas.draw()
        self._connect_hover(payload)

# ---------- Tab: Estadísticos ----------
class StatsTab(ttk.Frame):
    def __init__(self, parent, title: str):
        super().__init__(parent, style="Card.TFrame")
        self.df = None; self.min_date=None; self.max_date=None; self.out_dir=None; self.title_text=title
        top = ttk.Frame(self, style="Card.TFrame"); top.pack(fill="x", padx=12, pady=(12, 6))
        ttk.Label(top, text=title, style="Sub.TLabel").grid(row=0, column=0, sticky="w", padx=6)
        ttk.Label(top, text="Inicio (dd/mm/aaaa):").grid(row=0, column=1, sticky="e", padx=6)
        self.e_start = ttk.Entry(top, width=14); self.e_start.grid(row=0, column=2, sticky="w", padx=6)
        ttk.Label(top, text="Fin:").grid(row=0, column=3, sticky="e", padx=6)
        self.e_end = ttk.Entry(top, width=14); self.e_end.grid(row=0, column=4, sticky="w", padx=6)
        ttk.Button(top, text="Actualizar", command=self.update_stats).grid(row=0, column=5, padx=6)
        ttk.Button(top, text="Todo", command=self.set_all).grid(row=0, column=6, padx=6)
        ttk.Button(top, text="Últimos 15", command=lambda: self.set_last_days(15)).grid(row=0, column=7, padx=6)
        ttk.Button(top, text="Últimos 7", command=lambda: self.set_last_days(7)).grid(row=0, column=8, padx=6)
        summary = ttk.Frame(self, style="Card.TFrame"); summary.pack(fill="x", padx=12, pady=(0, 8))
        self.lbl_total_events = ttk.Label(summary, text="Eventos: -", style="TLabel"); self.lbl_total_events.pack(side="left", padx=8)
        self.lbl_total_people = ttk.Label(summary, text="Operadores: -", style="TLabel"); self.lbl_total_people.pack(side="left", padx=8)
        self.lbl_total_locations = ttk.Label(summary, text="Ubicaciones únicas (global): -", style="TLabel"); self.lbl_total_locations.pack(side="left", padx=8)
        bottom = ttk.Panedwindow(self, orient="horizontal"); bottom.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        frame_events = ttk.Frame(bottom, style="Card.TFrame")
        ttk.Label(frame_events, text="Monitoreos por persona", style="Sub.TLabel").pack(anchor="w", padx=8, pady=(8,4))
        self.tree_events = self._make_tree(frame_events, ["Operador", "Monitoreos"])
        bottom.add(frame_events, weight=1)
        frame_locs = ttk.Frame(bottom, style="Card.TFrame")
        ttk.Label(frame_locs, text="Ubicaciones únicas por persona", style="Sub.TLabel").pack(anchor="w", padx=8, pady=(8,4))
        self.tree_locs = self._make_tree(frame_locs, ["Operador", "Ubicaciones únicas"])
        bottom.add(frame_locs, weight=1)
        footer = ttk.Frame(self, style="Card.TFrame"); footer.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(footer, text="⬇ Exportar CSV (eventos / ubicaciones)", command=self.export_csv).pack(side="left", padx=6)
        ttk.Button(footer, text="⬇ Exportar XLSX (3 hojas)", command=self.export_xlsx).pack(side="left", padx=6)
        self.lbl_status = ttk.Label(footer, text="", style="TLabel"); self.lbl_status.pack(side="right", padx=6)
        self._ev=None; self._loc=None; self._summary=None
    def _make_tree(self, parent, columns):
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=16)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8,0), pady=(0,8))
        vsb.pack(side="left", fill="y", padx=(0,8), pady=(0,8))
        for c in columns:
            tree.heading(c, text=c)
            tree.column(c, anchor="w", stretch=True, width=220)
        return tree
    def set_data(self, df, dmin, dmax):
        self.df = df; self.min_date=dmin; self.max_date=dmax
        if dmin and dmax:
            self.e_start.delete(0, tk.END); self.e_start.insert(0, dmin.strftime("%d/%m/%Y"))
            self.e_end.delete(0, tk.END); self.e_end.insert(0, dmax.strftime("%d/%m/%Y"))
        self.update_stats()
    def set_out_dir(self, d): self.out_dir=d
    def get_date_range(self): 
        return safe_date(self.e_start.get().strip()), safe_date(self.e_end.get().strip())
    def set_all(self):
        if self.min_date and self.max_date:
            self.e_start.delete(0, tk.END); self.e_start.insert(0, self.min_date.strftime("%d/%m/%Y"))
            self.e_end.delete(0, tk.END); self.e_end.insert(0, self.max_date.strftime("%d/%m/%Y"))
            self.update_stats()
    def set_last_days(self, n):
        if self.max_date:
            start = pd.to_datetime(self.max_date) - pd.Timedelta(days=n-1)
            self.e_start.delete(0, tk.END); self.e_start.insert(0, start.strftime("%d/%m/%Y"))
            self.e_end.delete(0, tk.END); self.e_end.insert(0, self.max_date.strftime("%d/%m/%Y"))
            self.update_stats()
    def update_stats(self):
        if self.df is None: return
        start, end = self.get_date_range()
        dff = filter_by_dates(self.df, start, end)
        total_events = len(dff)
        total_people = dff[PERSON_COL].dropna().nunique() if PERSON_COL in dff.columns else 0
        locs_global = dff.dropna(subset=[X_COL, Y_COL]).copy()
        if not locs_global.empty:
            locs_global[X_COL] = pd.to_numeric(locs_global[X_COL], errors="coerce")
            locs_global[Y_COL] = pd.to_numeric(locs_global[Y_COL], errors="coerce")
            locs_global = locs_global.dropna(subset=[X_COL, Y_COL])
            locs_global = locs_global[[X_COL, Y_COL]].round(3).drop_duplicates()
            total_locations = len(locs_global)
        else:
            total_locations = 0
        self.lbl_total_events.config(text=f"Eventos: {total_events:,}".replace(",", " "))
        self.lbl_total_people.config(text=f"Operadores: {total_people}")
        self.lbl_total_locations.config(text=f"Ubicaciones únicas (global): {total_locations}")
        ev = compute_events_per_person(dff)
        loc = compute_locations_per_person(dff, round_decimals=3)
        avg_locs = loc["Ubicaciones únicas"].mean() if not loc.empty else 0
        avg_events = ev["Monitoreos"].mean() if not ev.empty else 0
        self.lbl_status.config(text=f"Promedio ubic./persona: {avg_locs:.2f} • Promedio eventos/persona: {avg_events:.2f}")
        self._fill_tree(self.tree_events, ev, ["Operador", "Monitoreos"], rename_map={PERSON_COL:"Operador"})
        self._fill_tree(self.tree_locs, loc, ["Operador", "Ubicaciones únicas"], rename_map={PERSON_COL:"Operador"})
        self._ev=ev.copy(); self._loc=loc.copy()
        self._summary = pd.DataFrame({
            "Total eventos":[total_events],
            "Total operadores":[total_people],
            "Ubicaciones únicas (global)":[total_locations],
            "Promedio ubicaciones/persona":[round(avg_locs,2)],
            "Promedio eventos/persona":[round(avg_events,2)]
        })
    def _fill_tree(self, tree: ttk.Treeview, df: pd.DataFrame, order_cols, rename_map=None):
        for i in tree.get_children(): tree.delete(i)
        if df is None or df.empty: return
        d = df.rename(columns=rename_map or {})
        cols = [c for c in order_cols if c in d.columns]
        for _, row in d[cols].iterrows():
            tree.insert("", "end", values=[row[c] for c in cols])
    def export_csv(self):
        if getattr(self, "_ev", None) is None or getattr(self, "_loc", None) is None:
            return
        if self.out_dir is None:
            messagebox.showwarning("Exportar", "Primero elige una carpeta de guardado (arriba)."); return
        try:
            p1 = os.path.join(self.out_dir, "eventos_por_persona.csv")
            p2 = os.path.join(self.out_dir, "ubicaciones_por_persona.csv")
            self._ev.to_csv(p1, index=False, encoding="utf-8")
            self._loc.to_csv(p2, index=False, encoding="utf-8")
            self.lbl_status.config(text=f"CSV exportados: {os.path.basename(p1)}, {os.path.basename(p2)}")
        except Exception as e:
            messagebox.showerror("Error al exportar CSV", str(e))
    def export_xlsx(self):
        if getattr(self, "_ev", None) is None or getattr(self, "_loc", None) is None or getattr(self, "_summary", None) is None:
            return
        if self.out_dir is None:
            messagebox.showwarning("Exportar", "Primero elige una carpeta de guardado (arriba)."); return
        try:
            path = os.path.join(self.out_dir, "resumen_estadisticos.xlsx")
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                self._summary.to_excel(xw, sheet_name="Resumen", index=False)
                self._ev.to_excel(xw, sheet_name="Eventos_por_persona", index=False)
                self._loc.to_excel(xw, sheet_name="Ubicaciones_por_persona", index=False)
            self.lbl_status.config(text=f"XLSX exportado: {path}")
        except Exception as e:
            messagebox.showerror("Error al exportar XLSX", str(e))

# ---------- Tab: Desempeño ----------
class PerfTab(BaseTab):
    def __init__(self, parent, title):
        super().__init__(parent, title)

        extras = ttk.Frame(self, style="Card.TFrame"); extras.pack(fill="x", padx=12, pady=(0,4))

        ttk.Label(extras, text="Operador:").pack(side="left", padx=6)
        self.cmb_person = ttk.Combobox(extras, state="readonly", width=40, values=())
        self.cmb_person.pack(side="left", padx=6)
        self.cmb_person.bind("<<ComboboxSelected>>", lambda e: self.update_chart())

        ttk.Label(extras, text="Agrupar por:").pack(side="left", padx=(18,6))
        self.cmb_group = ttk.Combobox(extras, state="readonly", width=10, values=("Día","Semana","Mes"))
        self.cmb_group.current(0)
        self.cmb_group.bind("<<ComboboxSelected>>", lambda e: self.update_chart())
        self.cmb_group.pack(side="left")

        ttk.Label(extras, text="Métrica:").pack(side="left", padx=(18,6))
        self.cmb_metric = ttk.Combobox(extras, state="readonly", width=20, values=("Monitoreos","Ubicaciones únicas"))
        self.cmb_metric.current(0)
        self.cmb_metric.bind("<<ComboboxSelected>>", lambda e: self.update_chart())
        self.cmb_metric.pack(side="left")

        self.var_both = tk.BooleanVar(value=False)
        ttk.Checkbutton(extras, text="Mostrar ambas", variable=self.var_both, command=self.update_chart)\
            .pack(side="left", padx=6)

        ttk.Label(extras, text="Redondeo XY:").pack(side="left", padx=(18,6))
        self.decimals = tk.IntVar(value=3)
        ttk.Spinbox(extras, from_=0, to=8, textvariable=self.decimals, width=5, command=self.update_chart)\
            .pack(side="left", padx=6)

        self.cmb_type["values"] = ("Barras", "Línea")
        self.cmb_type.current(0)

    def set_data(self, df, dmin, dmax):
        super().set_data(df, dmin, dmax)
        if df is not None and PERSON_COL in df.columns:
            vals = sorted([v for v in df[PERSON_COL].dropna().astype(str).unique()])
            self.cmb_person["values"] = vals
            if vals:
                self.cmb_person.current(0)

    def get_default_png_name(self):
        return "desempeno_operador_doble.png" if self.var_both.get() else "desempeno_operador.png"

    def _group_label(self, idx_series):
        if self.cmb_group.get() == "Semana":
            return idx_series.dt.strftime("Sem %U (%Y-%m-%d)")
        elif self.cmb_group.get() == "Mes":
            return idx_series.dt.strftime("%Y-%m")
        else:
            return idx_series.dt.strftime("%Y-%m-%d")

    def _series_monitoreos(self, dd):
        dd = dd.copy()
        dd["_idx"] = pd.to_datetime(dd["__FECHA"])
        if self.cmb_group.get() == "Semana":
            g = dd.groupby(pd.Grouper(key="_idx", freq="W-MON")).size().reset_index(name="valor")
        elif self.cmb_group.get() == "Mes":
            g = dd.groupby(pd.Grouper(key="_idx", freq="MS")).size().reset_index(name="valor")
        else:
            g = dd.groupby("_idx").size().reset_index(name="valor")
        g["_lab"] = self._group_label(g["_idx"])
        return g

    def _series_ubicaciones(self, dd, decimals=3):
        d2 = dd.dropna(subset=[X_COL, Y_COL]).copy()
        if d2.empty:
            return pd.DataFrame(columns=["_idx","valor","_lab"])
        d2[X_COL] = pd.to_numeric(d2[X_COL], errors="coerce")
        d2[Y_COL] = pd.to_numeric(d2[Y_COL], errors="coerce")
        d2 = d2.dropna(subset=[X_COL, Y_COL])
        d2["Xr"] = d2[X_COL].round(decimals)
        d2["Yr"] = d2[Y_COL].round(decimals)
        d2["_idx"] = pd.to_datetime(d2["__FECHA"])

        if self.cmb_group.get() == "Semana":
            d2["PER"] = d2["_idx"].dt.to_period("W-MON")
        elif self.cmb_group.get() == "Mes":
            d2["PER"] = d2["_idx"].dt.to_period("M")
        else:
            d2["PER"] = d2["_idx"].dt.to_period("D")

        d2 = d2.drop_duplicates(subset=["PER","Xr","Yr"])
        g = d2.groupby("PER").size().reset_index(name="valor")
        idx = g["PER"].dt.start_time
        g["_idx"] = pd.to_datetime(idx)
        g["_lab"] = self._group_label(g["_idx"])
        return g[["_idx","valor","_lab"]].sort_values("_idx")

    def update_chart(self):
        if self.df is None:
            return

        start, end = self.get_date_range()
        dff = filter_by_dates(self.df, start, end)

        person = (self.cmb_person.get() or "").strip()
        if not person:
            self._recreate_canvas(fig_w=11, fig_h=5)
            self.ax.set_title("Selecciona un operador")
            self.canvas.draw()
            return

        dd = dff[dff[PERSON_COL].astype(str) == person].dropna(subset=["__FECHA"]).copy()
        if dd.empty:
            self._recreate_canvas(fig_w=11, fig_h=5)
            self.ax.set_title(f"Sin datos para {person}")
            self.canvas.draw()
            return

        show_both = self.var_both.get()
        if show_both:
            self._recreate_canvas(fig_w=11, fig_h=7.5, n_axes=2)
            axes = self.ax_list
            metrics = ["Monitoreos", "Ubicaciones únicas"]
        else:
            self._recreate_canvas(fig_w=11, fig_h=5.6, n_axes=1)
            axes = [self.ax]
            metrics = [self.cmb_metric.get()]

        payloads = []
        for ax, metric in zip(axes, metrics):
            if metric == "Monitoreos":
                g = self._series_monitoreos(dd)
            else:
                g = self._series_ubicaciones(dd, decimals=int(self.decimals.get()))

            xlabels = g["_lab"].tolist()
            y = g["valor"].values

            if self.cmb_type.get() == "Línea":
                ax.plot(np.arange(len(y)), y, marker="o")
                payloads.append({"kind":"line","x":xlabels,"y":y.tolist(),"title":f"{metric} — {person}","ax":ax})
            else:
                bars = ax.bar(np.arange(len(y)), y)
                if self.var_labels.get():
                    for b, v in zip(bars, y):
                        ax.text(b.get_x()+b.get_width()/2, v, f"{int(v)}", va="bottom", ha="center")
                payloads.append({"kind":"bar","bars":[(b, xl, int(v)) for b, xl, v in zip(bars, xlabels, y)],
                                 "title":f"{metric} — {person}"})

            ax.set_xticks(np.arange(len(xlabels)))
            ax.set_xticklabels(xlabels, rotation=45, ha="right")
            ax.set_xlabel("Período")
            ax.set_ylabel(metric)
            ax.set_title(f"{metric} — {person}")
            self._apply_common_to(ax)

            # MA corta y tendencia (igual que otras pestañas)
            if len(y) >= 3:
                s = pd.Series(y).rolling(3, min_periods=1).mean()
                ax.plot(np.arange(len(y)), s.values, linestyle="--")
            if len(y) >= 2:
                t = np.arange(len(y))
                m, b = np.polyfit(t, y, 1)
                ax.plot(t, m*t + b, linestyle=":")

        self.fig.tight_layout()
        self.fig.subplots_adjust(bottom=0.28)
        self.canvas.draw()

        for p in payloads:
            self._connect_hover(p)

# === Pestaña Por Localidad ===
class LocalidadTab(BaseTab):
    """
    - Filtro por localidad (Todas / una)
    - Métrica: Monitoreos / Ubicaciones únicas
    - Gráficas: Barras por localidad (todas), Serie temporal, Heatmap op×loc (Top N), Pie (Ubicaciones)
    - Exportar resumen y exportar/mostrar tabla actual.
    - Incluye: toggles MA/Tendencia en Serie temporal. Heatmap con tamaño dinámico.
    """
    def __init__(self, parent, title: str):
        super().__init__(parent, title)

        extras = ttk.Frame(self, style="Card.TFrame"); extras.pack(fill="x", padx=12, pady=(0,4))

        ttk.Label(extras, text="Localidad:").pack(side="left", padx=6)
        self.cmb_loc = ttk.Combobox(extras, state="readonly", width=80, values=())
        self.cmb_loc.pack(side="left", padx=6)
        self.cmb_loc.bind("<<ComboboxSelected>>", lambda e: self.update_chart())

        ttk.Label(extras, text="Métrica:").pack(side="left", padx=(18,6))
        self.cmb_metric = ttk.Combobox(extras, state="readonly", width=20, values=("Monitoreos","Ubicaciones únicas"))
        self.cmb_metric.current(0); self.cmb_metric.bind("<<ComboboxSelected>>", lambda e: self.update_chart()); self.cmb_metric.pack(side="left")

        ttk.Label(extras, text="Gráfica:").pack(side="left", padx=(18,6))
        self.cmb_chart = ttk.Combobox(extras, state="readonly", width=24,
                                      values=("Barras por localidad", "Serie temporal", "Heatmap op×loc"))
        self.cmb_chart.current(0); self.cmb_chart.bind("<<ComboboxSelected>>", lambda e: self.update_chart()); self.cmb_chart.pack(side="left")

        ttk.Label(extras, text="Top N (Heatmap):").pack(side="left", padx=(18,6))
        self.topn = tk.IntVar(value=15)
        ttk.Spinbox(extras, from_=3, to=100, textvariable=self.topn, width=6, command=self.update_chart).pack(side="left", padx=6)

        ttk.Label(extras, text="Agrupar (serie):").pack(side="left", padx=(18,6))
        self.cmb_group = ttk.Combobox(extras, state="readonly", width=10, values=("Día","Semana","Mes"))
        self.cmb_group.current(0); self.cmb_group.bind("<<ComboboxSelected>>", lambda e: self.update_chart()); self.cmb_group.pack(side="left")

        # toggles para la serie temporal de localidad
        self.var_ma_loc = tk.BooleanVar(value=False)
        self.var_trend_loc = tk.BooleanVar(value=False)
        ttk.Checkbutton(extras, text="Media móvil", variable=self.var_ma_loc, command=self.update_chart).pack(side="left", padx=6)
        ttk.Checkbutton(extras, text="Tendencia", variable=self.var_trend_loc, command=self.update_chart).pack(side="left", padx=6)

        # Acciones a la derecha
        ttk.Button(extras, text="👁 Ver tabla", command=self.show_table).pack(side="right", padx=6)
        ttk.Button(extras, text="⬇ Exportar tabla XLSX", command=self.export_table_xlsx).pack(side="right", padx=6)
        ttk.Button(extras, text="⬇ Exportar tabla CSV", command=self.export_table_csv).pack(side="right", padx=6)
        ttk.Button(extras, text="⬇ Exportar XLSX (resumen)", command=self.export_xlsx).pack(side="right", padx=6)

        self.cmb_type["values"] = ("Barras", "Línea")
        self.cmb_type.current(0)

        self._last_table = None

    def set_data(self, df, dmin, dmax):
        super().set_data(df, dmin, dmax)
        locs = ["Todas"]
        if df is not None:
            locs += list_localidades(df)
        self.cmb_loc["values"] = locs
        if locs: self.cmb_loc.current(0)

    def get_default_png_name(self):
        mode = self.cmb_chart.get()
        metric = self.cmb_metric.get()
        base = "monit" if metric=="Monitoreos" else "ubics"
        return {
            "Barras por localidad": f"localidad_barras_{base}.png",
            "Serie temporal": f"localidad_serie_{base}.png",
            "Heatmap op×loc": f"localidad_heatmap_{base}.png"
        }.get(mode, f"localidad_{base}.png")

    # ----- Exportar tabla visible -----
    def _table_filename_base(self):
        sel = (self.cmb_loc.get() or "Todas").replace("/", "-").replace("\\", "-")
        metric = "monitoreos" if self.cmb_metric.get()=="Monitoreos" else "ubicaciones"
        mode = self.cmb_chart.get().split()[0].lower()
        return f"tabla_{mode}_{metric}_{sel}"

    def export_table_csv(self):
        if self._last_table is None or self._last_table.empty:
            messagebox.showinfo("Exportar", "No hay tabla visible para exportar.")
            return
        if self.out_dir is None:
            messagebox.showwarning("Exportar", "Primero elige una carpeta de guardado (arriba).")
            return
        name = self._table_filename_base() + ".csv"
        path = os.path.join(self.out_dir, name)
        try:
            self._last_table.to_csv(path, index=False, encoding="utf-8")
            self.lbl_status.config(text=f"Tabla exportada: {path}")
        except Exception as e:
            messagebox.showerror("Error al exportar CSV", str(e))

    def export_table_xlsx(self):
        if self._last_table is None or self._last_table.empty:
            messagebox.showinfo("Exportar", "No hay tabla visible para exportar.")
            return
        if self.out_dir is None:
            messagebox.showwarning("Exportar", "Primero elige una carpeta de guardado (arriba).")
            return
        name = self._table_filename_base() + ".xlsx"
        path = os.path.join(self.out_dir, name)
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                self._last_table.to_excel(xw, sheet_name="Datos", index=False)
            self.lbl_status.config(text=f"Tabla exportada: {path}")
        except Exception as e:
            messagebox.showerror("Error al exportar XLSX", str(e))

    def export_xlsx(self):
        if self.df is None: return
        if self.out_dir is None:
            messagebox.showwarning("Exportar", "Primero elige una carpeta de guardado (arriba)."); return
        start, end = self.get_date_range()
        dff = filter_by_dates(self.df, start, end)
        summ = compute_summary_by_localidad(dff, round_decimals=3)
        sel = (self.cmb_loc.get() or "").strip()
        if sel and sel != "Todas":
            summ = summ[summ["Localidad"] == sel]
        if summ.empty:
            messagebox.showinfo("Exportar", "No hay datos para exportar en el rango/selección.")
            return
        path = os.path.join(self.out_dir, "resumen_por_localidad.xlsx")
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                summ.to_excel(xw, sheet_name="Resumen_Localidades", index=False)
            self.lbl_status.config(text=f"XLSX exportado: {path}")
        except Exception as e:
            messagebox.showerror("Error al exportar XLSX", str(e))

    def show_table(self):
        if self._last_table is None or self._last_table.empty:
            messagebox.showinfo("Tabla", "No hay datos para mostrar. Actualiza la gráfica primero.")
            return
        win = tk.Toplevel(self)
        win.title("Datos graficados")
        win.geometry("1000x600")
        frm = ttk.Frame(win, style="Card.TFrame"); frm.pack(fill="both", expand=True, padx=10, pady=10)
        cols = list(self._last_table.columns)
        tree = ttk.Treeview(frm, columns=cols, show="headings")
        vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        hsb.pack(side="bottom", fill="x")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, anchor="w", stretch=True, width=max(120, int(900/len(cols))))
        for _, row in self._last_table.iterrows():
            tree.insert("", "end", values=[row[c] for c in cols])

    def update_chart(self):
        if self.df is None: return
        start, end = self.get_date_range()
        dff_all = filter_by_dates(self.df, start, end)
        if dff_all.empty:
            self._recreate_canvas(fig_w=11, fig_h=5.6)
            self.ax.set_title("Sin datos en el rango seleccionado")
            self.canvas.draw()
            self._last_table = None
            return

        sel = (self.cmb_loc.get() or "").strip()
        mode = self.cmb_chart.get()
        N = int(self.topn.get())
        metric = self.cmb_metric.get()

        # dataset filtrado por localidad cuando aplica
        if LOC_COL in dff_all.columns and sel and sel != "Todas":
            dff_sel = dff_all[dff_all[LOC_COL].astype(str).str.strip() == sel].copy()
        else:
            dff_sel = dff_all.copy()

        # Barras por localidad (todas)
        if mode == "Barras por localidad":
            g = compute_metric_by_localidad(dff_all, metric=metric, round_decimals=3)
            if g.empty:
                self._recreate_canvas(); self.ax.set_title("Sin datos para la selección"); self.canvas.draw(); self._last_table=None; return
            g = g.rename(columns={"valor": metric})
            self._last_table = g.copy()

            names = g["Localidad"].tolist()
            vals = g[metric].astype(int).to_numpy()

            self._recreate_canvas(fig_w=12.5, fig_h=5.8)
            x = np.arange(len(names)); bars = self.ax.bar(x, vals)
            self.ax.set_xticks(x); self.ax.set_xticklabels(names, rotation=45, ha="right", fontsize=9)
            self.ax.set_xlabel("Localidad"); self.ax.set_ylabel(metric)
            self.ax.set_title(f"{metric} por localidad (todas)")
            if self.var_labels.get():
                for b, v in zip(bars, vals):
                    self.ax.text(b.get_x()+b.get_width()/2, v, f"{int(v)}", va="bottom", ha="center", fontsize=8)
            self._apply_common_to(self.ax)
            self.fig.tight_layout(rect=[0.03,0.28,0.98,0.95])
            self.canvas.draw()
            payload = {"kind":"bar","bars":[(b, n, int(v)) for b, n, v in zip(bars, names, vals)],"title":metric}
            self._connect_hover(payload); 
            return

        # Serie temporal
        if mode == "Serie temporal":
            g = series_metric_localidad(dff_sel, metric=metric, group=self.cmb_group.get(), round_decimals=3)
            if g.empty:
                self._recreate_canvas(); self.ax.set_title("Sin datos para la selección"); self.canvas.draw(); self._last_table=None; return
            table = g.rename(columns={"_lab":"Periodo","valor":metric})[["Periodo", metric]].copy()
            self._last_table = table

            xlabels = g["_lab"].tolist(); y = g["valor"].values
            self._recreate_canvas(fig_w=11.8, fig_h=min(7.6, max(5.6, 0.06*len(xlabels))))
            if self.cmb_type.get() == "Línea":
                self.ax.plot(np.arange(len(y)), y, marker="o")
            else:
                bars = self.ax.bar(np.arange(len(y)), y)
                if self.var_labels.get():
                    for i, v in enumerate(y): self.ax.text(i, v, f"{int(v)}", va="bottom", ha="center", fontsize=8)
            self.ax.set_xticks(np.arange(len(xlabels))); self.ax.set_xticklabels(xlabels, rotation=45, ha="right")
            self.ax.set_xlabel("Período"); self.ax.set_ylabel(metric)
            tit = f"Serie temporal — {sel}" if sel and sel != "Todas" else "Serie temporal — Todas las localidades"
            self.ax.set_title(f"{tit} ({metric})")
            self._apply_common_to(self.ax)

            if self.var_ma_loc.get() and len(y) >= 3:
                s = pd.Series(y).rolling(3, min_periods=1).mean(); self.ax.plot(np.arange(len(y)), s.values, linestyle="--")
            if self.var_trend_loc.get() and len(y) >= 2:
                t = np.arange(len(y)); m, b = np.polyfit(t, y, 1); self.ax.plot(t, m*t+b, linestyle=":")

            self.fig.tight_layout(rect=[0.03,0.28,0.98,0.95])
            self.canvas.draw()
            payload = {"kind":"line","x":xlabels,"y":y.tolist(),"title":metric,"ax":self.ax}
            self._connect_hover(payload); 
            return

        # Heatmap operador × localidad (Top N) — tamaño dinámico
        if mode == "Heatmap op×loc":
            if LOC_COL not in self.df.columns:
                self._recreate_canvas(); self.ax.set_title("No existe columna LOCALIDAD"); self.canvas.draw(); self._last_table=None; return
            d = dff_sel.dropna(subset=[PERSON_COL, LOC_COL]).copy()
            if d.empty:
                self._recreate_canvas(); self.ax.set_title("Sin datos para la selección"); self.canvas.draw(); self._last_table=None; return

            d[PERSON_COL] = d[PERSON_COL].astype(str).str.strip()
            d[LOC_COL] = d[LOC_COL].astype(str).str.strip()

            # Top-N por columnas (localidades) según total de la métrica
            if metric == "Monitoreos":
                top = (d.groupby(LOC_COL).size().sort_values(ascending=False).head(N)).index.tolist()
                pivot = d[d[LOC_COL].isin(top)].pivot_table(index=PERSON_COL, columns=LOC_COL, values=DATE_COL, aggfunc="count", fill_value=0)
            else:
                # Ubicaciones únicas: contar (Xr,Yr) únicos por operador×localidad
                if not {X_COL, Y_COL}.issubset(d.columns):
                    self._recreate_canvas(); self.ax.set_title("No existen columnas de coordenadas"); self.canvas.draw(); self._last_table=None; return
                d[X_COL] = pd.to_numeric(d[X_COL], errors="coerce")
                d[Y_COL] = pd.to_numeric(d[Y_COL], errors="coerce")
                d = d.dropna(subset=[X_COL, Y_COL])
                d["Xr"] = d[X_COL].round(3); d["Yr"] = d[Y_COL].round(3)
                d = d.drop_duplicates(subset=[PERSON_COL, LOC_COL, "Xr", "Yr"])
                top = (d.groupby(LOC_COL).size().sort_values(ascending=False).head(N)).index.tolist()
                pivot = d[d[LOC_COL].isin(top)].groupby([PERSON_COL, LOC_COL]).size().unstack(fill_value=0)

            # ordenar filas/cols por suma
            pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]
            pivot = pivot[pivot.sum(axis=0).sort_values(ascending=False).index]

            # tabla para exportar/mostrar
            self._last_table = pivot.reset_index().rename(columns={PERSON_COL: "Operador"}).copy()

            rows = pivot.index.tolist(); cols = pivot.columns.tolist()
            data = pivot.to_numpy(dtype=float)

            # tamaño dinámico
            h = min(7.6, max(4.8, 0.35 * len(rows)))
            w = min(13.0, max(8.5, 0.55 * len(cols)))
            self._recreate_canvas(fig_w=w, fig_h=h)
            im = self.ax.imshow(data, aspect="auto", interpolation="nearest")
            self.ax.set_yticks(np.arange(len(rows))); self.ax.set_yticklabels(rows, fontsize=8)
            self.ax.set_xticks(np.arange(len(cols))); self.ax.set_xticklabels(cols, rotation=45, ha="right", fontsize=8)
            self.ax.set_xlabel("Localidad"); self.ax.set_ylabel("Operador")
            self.ax.set_title(f"Heatmap operador × localidad ({metric})")
            cbar = self.fig.colorbar(im, ax=self.ax, fraction=0.022, pad=0.02)
            cbar.ax.set_ylabel(metric, rotation=90, va="bottom")

            self.fig.tight_layout(rect=[0.05,0.12,0.96,0.95])
            self.canvas.draw()
            payload = {"kind":"heatmap","ax":self.ax,"data":data,"rows":rows,"cols":cols,"metric":metric}
            self._connect_hover(payload)
            return

# ---------- Main -----------
if __name__ == "__main__":
    app = App()
    app.mainloop()
