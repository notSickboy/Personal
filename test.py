import os
import subprocess

# Ruta donde están los archivos
ruta = r"C:\Users\Usuario_sino\Documents\Python_Monitoreo\lib"

# Nombres de los scripts en orden
scripts = ['lib_1', 'lib_2', 'lib_3', 'lib_4']

for script in scripts:
    ruta_completa = os.path.join(ruta, script)
    ruta_py = ruta_completa + '.py'

    # Renombrar a .py
    os.rename(ruta_completa, ruta_py)

    try:
        print(f"▶️ Ejecutando: {script}.py")
        subprocess.run(['python', ruta_py], check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ Error al ejecutar {script}.py: {e}")
    finally:
        # Restaurar nombre sin extensión
        os.rename(ruta_py, ruta_completa)
        print(f"✅ Restaurado: {script}")
