# -*- coding: utf-8 -*-
import arcpy, math, sys, traceback
arcpy.env.overwriteOutput = True

# Tkinter para Python 2.7 (ArcMap 10.x)
import Tkinter as tk
import ttk
import tkMessageBox

# === Utilidades ===
def get_map_layers():
    mxd = arcpy.mapping.MapDocument("CURRENT")
    dfs = arcpy.mapping.ListDataFrames(mxd)
    if not dfs:
        raise RuntimeError("No hay DataFrames en el MXD.")
    df = dfs[0]
    return [lyr for lyr in arcpy.mapping.ListLayers(mxd, "", df) if lyr.isFeatureLayer]

def list_numeric_fields(layer_name):
    num_types = set(["Double", "Single", "Integer", "SmallInteger"])
    return [f.name for f in arcpy.ListFields(layer_name) if f.type in num_types]

def list_all_fields(layer_name):
    return [f.name for f in arcpy.ListFields(layer_name)]

def sr_of(layer_name):
    return arcpy.Describe(layer_name).spatialReference


def crear_desp_2capas(l_start, x_s, y_s, id_s, l_end, x_e, y_e, id_e):
    # --- Validaciones básicas ---
    for lyr in [l_start, l_end]:
        if not arcpy.Exists(lyr):
            raise RuntimeError("No existe la capa: {}".format(lyr))
    for lyr, flds in [(l_start, [x_s, y_s, id_s]), (l_end, [x_e, y_e, id_e])]:
        names = [f.name for f in arcpy.ListFields(lyr)]
        miss = [f for f in flds if f not in names]
        if miss: raise RuntimeError("En '{}' faltan campos: {}".format(lyr, ", ".join(miss)))

    sr_start = sr_of(l_start)
    sr_end   = sr_of(l_end)

    # === Crear capa en memoria usando la salida devuelta ===
    out_name = "Desplazamientos_2L"
    if arcpy.Exists(r"in_memory\{}".format(out_name)):
        arcpy.Delete_management(r"in_memory\{}".format(out_name))

    result = arcpy.CreateFeatureclass_management("in_memory", out_name, "POLYLINE", spatial_reference=sr_start)
    out_fc = result.getOutput(0)  # Esta es la referencia segura

    # === Ahora sí, agregar campos ===
    arcpy.AddField_management(out_fc, "ID_match", "TEXT", field_length=80)
    for fld in ["X_ini","Y_ini","X_fin","Y_fin","dX","dY","Desp_m","Azim_deg"]:
        arcpy.AddField_management(out_fc, fld, "DOUBLE")

    # === Índice finales ===
    end_dict = {}
    with arcpy.da.SearchCursor(l_end, [id_e, x_e, y_e, "SHAPE@"]) as curE:
        for idv, xe, ye, shp in curE:
            if idv is None or xe is None or ye is None:
                continue
            if sr_end.name != sr_start.name and shp:
                try:
                    shp_proj = shp.projectAs(sr_start)
                    xe2, ye2 = shp_proj.firstPoint.X, shp_proj.firstPoint.Y
                    end_dict[unicode(idv)] = (xe2, ye2)
                    continue
                except:
                    pass
            end_dict[unicode(idv)] = (float(xe), float(ye))

    # === Crear líneas ===
    count_ins = 0
    with arcpy.da.SearchCursor(l_start, [id_s, x_s, y_s, "SHAPE@"]) as sCur, \
         arcpy.da.InsertCursor(out_fc, ["SHAPE@", "ID_match", "X_ini", "Y_ini", "X_fin", "Y_fin", "dX", "dY", "Desp_m", "Azim_deg"]) as iCur:
        for idv, xs, ys, shp_s in sCur:
            if idv is None or xs is None or ys is None:
                continue
            key = unicode(idv)
            if key not in end_dict:
                continue
            xe, ye = end_dict[key]
            if sr_end.name != sr_start.name and shp_s:
                try:
                    shp_s = shp_s.projectAs(sr_start)
                    xs, ys = shp_s.firstPoint.X, shp_s.firstPoint.Y
                except:
                    pass
            xs, ys, xe, ye = float(xs), float(ys), float(xe), float(ye)
            dX, dY = (xe - xs), (ye - ys)
            dist   = math.hypot(dX, dY)
            azim   = (math.degrees(math.atan2(dY, dX)) + 360.0) % 360.0
            line = arcpy.Polyline(arcpy.Array([arcpy.Point(xs, ys), arcpy.Point(xe, ye)]), sr_start)
            iCur.insertRow([line, key, xs, ys, xe, ye, dX, dY, dist, azim])
            count_ins += 1

    # === Agregar al mapa ===
    mxd = arcpy.mapping.MapDocument("CURRENT")
    df  = arcpy.mapping.ListDataFrames(mxd)[0]
    arcpy.MakeFeatureLayer_management(out_fc, "lyr_" + out_name)
    arcpy.mapping.AddLayer(df, arcpy.mapping.Layer("lyr_" + out_name), "TOP")
    return out_fc, count_ins


# === GUI ===
class ParamForm(object):
    def __init__(self, master, layers):
        self.master = master
        self.master.title("Desplazamientos: dos capas")
        self.master.resizable(False, False)
        pad = 6
        self.layer_names = [l.name for l in layers]

        self.var_lstart = tk.StringVar()
        self.var_xs = tk.StringVar()
        self.var_ys = tk.StringVar()
        self.var_ids = tk.StringVar()
        self.var_lend = tk.StringVar()
        self.var_xe = tk.StringVar()
        self.var_ye = tk.StringVar()
        self.var_ide = tk.StringVar()

        row=0
        ttk.Label(master, text="Capa INICIO").grid(row=row, column=0, sticky="w", padx=pad, pady=pad)
        self.cb_ls = ttk.Combobox(master, values=self.layer_names, textvariable=self.var_lstart, state="readonly", width=32)
        self.cb_ls.grid(row=row, column=1, columnspan=3, sticky="we", padx=pad, pady=pad)
        self.cb_ls.bind("<<ComboboxSelected>>", self.on_change_start); row+=1

        ttk.Label(master, text="X_ini").grid(row=row, column=0, sticky="w", padx=pad, pady=pad)
        self.cb_xs = ttk.Combobox(master, values=[], textvariable=self.var_xs, state="readonly", width=15)
        self.cb_xs.grid(row=row, column=1, padx=pad, pady=pad)
        ttk.Label(master, text="Y_ini").grid(row=row, column=2, sticky="w", padx=pad, pady=pad)
        self.cb_ys = ttk.Combobox(master, values=[], textvariable=self.var_ys, state="readonly", width=15)
        self.cb_ys.grid(row=row, column=3, padx=pad, pady=pad); row+=1

        ttk.Label(master, text="ID_ini").grid(row=row, column=0, sticky="w", padx=pad, pady=pad)
        self.cb_ids = ttk.Combobox(master, values=[], textvariable=self.var_ids, state="readonly", width=15)
        self.cb_ids.grid(row=row, column=1, padx=pad, pady=pad); row+=1

        ttk.Separator(master, orient="horizontal").grid(row=row, column=0, columnspan=4, sticky="we", padx=pad, pady=pad); row+=1

        ttk.Label(master, text="Capa FINAL").grid(row=row, column=0, sticky="w", padx=pad, pady=pad)
        self.cb_le = ttk.Combobox(master, values=self.layer_names, textvariable=self.var_lend, state="readonly", width=32)
        self.cb_le.grid(row=row, column=1, columnspan=3, sticky="we", padx=pad, pady=pad)
        self.cb_le.bind("<<ComboboxSelected>>", self.on_change_end); row+=1

        ttk.Label(master, text="X_fin").grid(row=row, column=0, sticky="w", padx=pad, pady=pad)
        self.cb_xe = ttk.Combobox(master, values=[], textvariable=self.var_xe, state="readonly", width=15)
        self.cb_xe.grid(row=row, column=1, padx=pad, pady=pad)
        ttk.Label(master, text="Y_fin").grid(row=row, column=2, sticky="w", padx=pad, pady=pad)
        self.cb_ye = ttk.Combobox(master, values=[], textvariable=self.var_ye, state="readonly", width=15)
        self.cb_ye.grid(row=row, column=3, padx=pad, pady=pad); row+=1

        ttk.Label(master, text="ID_fin").grid(row=row, column=0, sticky="w", padx=pad, pady=pad)
        self.cb_ide = ttk.Combobox(master, values=[], textvariable=self.var_ide, state="readonly", width=15)
        self.cb_ide.grid(row=row, column=1, padx=pad, pady=pad)

        # Botones
        ttk.Button(master, text="Crear", command=self.run_create).grid(row=row, column=2, padx=pad, pady=pad, sticky="e")
        ttk.Button(master, text="Cerrar", command=self.master.destroy).grid(row=row, column=3, padx=pad, pady=pad, sticky="w")

    def on_change_start(self, *args):
        lyr = self.var_lstart.get()
        if not lyr: return
        nums = list_numeric_fields(lyr)
        alls = list_all_fields(lyr)
        self.cb_xs["values"] = nums
        self.cb_ys["values"] = nums
        self.cb_ids["values"] = alls

    def on_change_end(self, *args):
        lyr = self.var_lend.get()
        if not lyr: return
        nums = list_numeric_fields(lyr)
        alls = list_all_fields(lyr)
        self.cb_xe["values"] = nums
        self.cb_ye["values"] = nums
        self.cb_ide["values"] = alls

    def run_create(self):
        try:
            if not all([self.var_lstart.get(), self.var_xs.get(), self.var_ys.get(), self.var_ids.get(),
                        self.var_lend.get(), self.var_xe.get(), self.var_ye.get(), self.var_ide.get()]):
                tkMessageBox.showerror("Error", "Completa todas las selecciones.")
                return

            # Deshabilitar botones mientras corre
            for child in self.master.winfo_children():
                try: child.configure(state="disabled")
                except: pass

            out_fc, n = crear_desp_2capas(
                self.var_lstart.get(), self.var_xs.get(), self.var_ys.get(), self.var_ids.get(),
                self.var_lend.get(),   self.var_xe.get(), self.var_ye.get(), self.var_ide.get()
            )
            tkMessageBox.showinfo("Listo", u"Capa creada: {}\nLíneas generadas: {}".format(out_fc, n))
        except Exception as e:
            tb = traceback.format_exc()
            tkMessageBox.showerror("Fallo", u"Se produjo un error:\n{}\n\n{}".format(e, tb))
        finally:
            # Rehabilitar UI
            for child in self.master.winfo_children():
                try: child.configure(state="normal")
                except: pass

# === Lanzar ventana ===
layers = get_map_layers()
root = tk.Tk(); root.withdraw()
top = tk.Toplevel()
ParamForm(top, layers)
root.mainloop()
