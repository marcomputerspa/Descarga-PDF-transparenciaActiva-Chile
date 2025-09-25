import tkinter as tk
from tkinter import messagebox
import subprocess
import threading
import os

CONFIG_FILE = 'config.txt'
SCRIPT_VERSION = 'v2.4.1'

def get_default_config():
    """Devuelve un diccionario con los valores por defecto de configuración.

    Se usan cadenas para conservar espacios y formatos de separadores.
    """
    return {
        # Rutas y archivos
        'carpeta_descargas': 'descargas',
        'csv_folder': 'csvs',
        'col_enlace': '',
        'carpeta_logs': 'logs',
        # Comportamiento
        'eliminar_csv_al_final': 'false',
        # Prefijo/sufijo
        'usar_prefijo_columna': 'false',
        'columna_prefijo': '',
        'tipo_prefijo': 'prefijo',
        'Nombre_de_la_carpeta': '',
        'separador_prefijo': '_',
        # Carpeta combinada
        'usar_carpeta_combinada': 'false',
        'columna_carpeta_1': '',
        'columna_carpeta_2': '',
        'separador_carpeta_combinada': '_',
    }

def leer_config():
    """Lee el archivo de configuración `config.txt` y devuelve un diccionario.

    Notas:
    - Se preservan los espacios en los valores (no se hace strip al valor),
      esto es importante para permitir separadores como un espacio en blanco.
    - Se rellenan valores por defecto para nuevas claves si no existen.
    """
    # Si no existe, crear con valores por defecto
    if not os.path.exists(CONFIG_FILE):
        defaults = get_default_config()
        guardar_config(defaults)
        return defaults.copy()

    config = {}
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            # Preservar espacios en el valor (ej. separador = ' ')
            line = line.rstrip('\r\n')
            if '=' in line:
                k, v = line.split('=', 1)
                config[k.strip()] = v  # no strip del valor
    # Completar claves faltantes con valores por defecto
    defaults = get_default_config()
    for dk, dv in defaults.items():
        if dk not in config:
            config[dk] = dv
    return config

def guardar_config(config):
    """Guarda el diccionario de configuración en `config.txt`.

    - Convierte booleanos a 'true'/'false'.
    - Escribe cada par clave=valor en una línea.
    """
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        for k, v in config.items():
            # Escribir todos los pares clave=valor. Convertir booleanos a 'true'/'false'.
            if isinstance(v, bool):
                v = 'true' if v else 'false'
            f.write(f"{k}={v}\n")

def ejecutar_script():
    """Ejecuta este mismo script en modo 'run' en una nueva ventana de PowerShell.

    Esto permite que la descarga se ejecute en una consola separada
    mientras la interfaz se cierra.
    """
    # Ejecuta el script en una nueva ventana de consola
    script_path = os.path.abspath(__file__)
    # Si tienes otro script principal, cámbialo aquí
    comando = f'powershell.exe -NoExit -Command "python \'{script_path}\' run"'
    subprocess.Popen(comando, shell=True)

def iniciar(config_vars):
    """Persiste la configuración editada y lanza la ejecución."""
    # Guardar los valores editados en config.txt
    config = {k: var.get() for k, var in config_vars.items()}
    guardar_config(config)
    ejecutar_script()

def mostrar_interfaz():
    """Crea y muestra la interfaz de configuración con Tkinter.

    La UI soporta:
    - Activar prefijo/sufijo de carpeta a partir de una columna del CSV.
    - Crear una carpeta combinada con dos columnas.
    - Elegir separadores (incluyendo espacio en blanco).
    - Auto-guardado al modificar cualquier control.
    """

    config = leer_config()
    root = tk.Tk()
    root.title(f'Configuración y Ejecución - {SCRIPT_VERSION}')
    root.update_idletasks()
    # Geometría inicial mínima; se recalcula tras construir la UI
    width = 560
    height = 420
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    config_vars = {}
    row = 0
    # Campos normales
    for k, v in config.items():
        if k in ['usar_prefijo_columna', 'columna_prefijo', 'tipo_prefijo', 'Nombre_de_la_carpeta',
                 'usar_carpeta_combinada', 'columna_carpeta_1', 'columna_carpeta_2', 'separador_carpeta_combinada', 'separador_prefijo']:
            continue
        if k == 'eliminar_csv_al_final':
            tk.Label(root, text=k).grid(row=row, column=0, sticky='w', padx=5, pady=5)
            var = tk.BooleanVar(value=(v.strip().lower() == 'true'))
            chk = tk.Checkbutton(root, variable=var, text='Eliminar los CSV al finalizar')
            chk.grid(row=row, column=1, padx=5, pady=5, sticky='w')
            config_vars[k] = var
            row += 1
        else:
            tk.Label(root, text=k).grid(row=row, column=0, sticky='w', padx=5, pady=5)
            var = tk.StringVar(value=v)
            entry = tk.Entry(root, textvariable=var, width=40)
            entry.grid(row=row, column=1, padx=5, pady=5)
            config_vars[k] = var
            row += 1

    # Checkbox para activar prefijo de carpeta
    usar_prefijo_var = tk.BooleanVar(value=(config.get('usar_prefijo_columna', 'false').lower() == 'true'))
    chk_prefijo = tk.Checkbutton(root, variable=usar_prefijo_var, text='Usar valor de otra columna como prefijo o sufijo de carpeta')
    chk_prefijo.grid(row=row, column=0, columnspan=2, padx=5, pady=5, sticky='w')
    config_vars['usar_prefijo_columna'] = usar_prefijo_var
    row += 1


    # --- Bloque de prefijo/sufijo ---
    # Opción para prefijo o sufijo
    tk.Label(root, text='¿Usar como prefijo o sufijo?').grid(row=row, column=0, sticky='w', padx=5, pady=5)
    tipo_prefijo_var = tk.StringVar(value=config.get('tipo_prefijo', 'prefijo'))
    radio_prefijo = tk.Radiobutton(root, text='Prefijo', variable=tipo_prefijo_var, value='prefijo')
    radio_sufijo = tk.Radiobutton(root, text='Sufijo', variable=tipo_prefijo_var, value='sufijo')
    radio_prefijo.grid(row=row, column=1, sticky='w')
    radio_sufijo.grid(row=row, column=1, sticky='e')
    config_vars['tipo_prefijo'] = tipo_prefijo_var
    row += 1

    # Campo para nombre de la columna a usar como prefijo
    tk.Label(root, text='Nombre de la columna').grid(row=row, column=0, sticky='w', padx=5, pady=5)
    columna_prefijo_var = tk.StringVar(value=config.get('columna_prefijo', ''))
    entry_col_prefijo = tk.Entry(root, textvariable=columna_prefijo_var, width=40)
    entry_col_prefijo.grid(row=row, column=1, padx=5, pady=5)
    config_vars['columna_prefijo'] = columna_prefijo_var
    row += 1

    # Campo para el nombre de la carpeta
    tk.Label(root, text='Nombre de la carpeta').grid(row=row, column=0, sticky='w', padx=5, pady=5)
    nombre_prefijo_var = tk.StringVar(value=config.get('Nombre_de_la_carpeta', ''))
    entry_nombre_prefijo = tk.Entry(root, textvariable=nombre_prefijo_var, width=40)
    entry_nombre_prefijo.grid(row=row, column=1, padx=5, pady=5)
    config_vars['Nombre_de_la_carpeta'] = nombre_prefijo_var
    row += 1

    # Separador para prefijo/sufijo (ahora dentro del bloque)
    tk.Label(root, text='Separador prefijo/sufijo').grid(row=row, column=0, sticky='w', padx=5, pady=5)
    # Mapeo de separadores: mostrar 'espacio en blanco' pero guardar/aplicar ' '
    separadores_display = ['_', '-', 'espacio en blanco']
    sep_display_to_value = {'_': '_', '-': '-', 'espacio en blanco': ' '}
    sep_value_to_display = {'_': '_', '-': '-', ' ': 'espacio en blanco', '': 'espacio en blanco'}

    sep_prefijo_val = config.get('separador_prefijo', '_')
    sep_prefijo_display = sep_value_to_display.get(sep_prefijo_val, 'espacio en blanco' if sep_prefijo_val.strip() == '' else sep_prefijo_val)
    separador_prefijo_var = tk.StringVar(value=sep_prefijo_display)
    sep_pref_menu = tk.OptionMenu(root, separador_prefijo_var, *separadores_display)
    sep_pref_menu.grid(row=row, column=1, padx=5, pady=5, sticky='w')
    config_vars['separador_prefijo'] = separador_prefijo_var
    row += 1

    # --- Fin bloque prefijo/sufijo ---

    # Checkbox para activar carpeta combinada
    usar_carpeta_combinada_var = tk.BooleanVar(value=(config.get('usar_carpeta_combinada', 'false').lower() == 'true'))
    chk_carpeta_combinada = tk.Checkbutton(root, variable=usar_carpeta_combinada_var, text='Usar dos columnas para una sola carpeta combinada')
    chk_carpeta_combinada.grid(row=row, column=0, columnspan=2, padx=5, pady=5, sticky='w')
    config_vars['usar_carpeta_combinada'] = usar_carpeta_combinada_var
    row += 1

    # Campos para los nombres de las columnas de la carpeta combinada
    tk.Label(root, text='Columna 1').grid(row=row, column=0, sticky='w', padx=5, pady=5)
    columna_carpeta_1_var = tk.StringVar(value=config.get('columna_carpeta_1', ''))
    entry_col_carpeta_1 = tk.Entry(root, textvariable=columna_carpeta_1_var, width=40)
    entry_col_carpeta_1.grid(row=row, column=1, padx=5, pady=5)
    config_vars['columna_carpeta_1'] = columna_carpeta_1_var
    row += 1

    tk.Label(root, text='Columna 2').grid(row=row, column=0, sticky='w', padx=5, pady=5)
    columna_carpeta_2_var = tk.StringVar(value=config.get('columna_carpeta_2', ''))
    entry_col_carpeta_2 = tk.Entry(root, textvariable=columna_carpeta_2_var, width=40)
    entry_col_carpeta_2.grid(row=row, column=1, padx=5, pady=5)
    config_vars['columna_carpeta_2'] = columna_carpeta_2_var
    row += 1

    # Separador para carpeta combinada
    tk.Label(root, text='Separador carpeta combinada').grid(row=row, column=0, sticky='w', padx=5, pady=5)
    sep_comb_val = config.get('separador_carpeta_combinada', '_')
    sep_comb_display = sep_value_to_display.get(sep_comb_val, 'espacio en blanco' if sep_comb_val.strip() == '' else sep_comb_val)
    separador_carpeta_combinada_var = tk.StringVar(value=sep_comb_display)
    sep_comb_menu = tk.OptionMenu(root, separador_carpeta_combinada_var, *separadores_display)
    sep_comb_menu.grid(row=row, column=1, padx=5, pady=5, sticky='w')
    config_vars['separador_carpeta_combinada'] = separador_carpeta_combinada_var
    row += 1

    # Función para mostrar/ocultar campos según el checkbox
    def actualizar_visibilidad(*args):
        if usar_prefijo_var.get():
            entry_col_prefijo.config(state='normal')
            entry_nombre_prefijo.config(state='normal')
            radio_prefijo.config(state='normal')
            radio_sufijo.config(state='normal')
            sep_pref_menu.config(state='normal')
        else:
            entry_col_prefijo.config(state='disabled')
            entry_nombre_prefijo.config(state='disabled')
            radio_prefijo.config(state='disabled')
            radio_sufijo.config(state='disabled')
            sep_pref_menu.config(state='disabled')
        # Carpeta combinada
        if usar_carpeta_combinada_var.get():
            entry_col_carpeta_1.config(state='normal')
            entry_col_carpeta_2.config(state='normal')
            sep_comb_menu.config(state='normal')
        else:
            entry_col_carpeta_1.config(state='disabled')
            entry_col_carpeta_2.config(state='disabled')
            sep_comb_menu.config(state='disabled')
    usar_prefijo_var.trace_add('write', actualizar_visibilidad)
    usar_carpeta_combinada_var.trace_add('write', actualizar_visibilidad)
    actualizar_visibilidad()

    # --- Autoguardado inmediato al cambiar cualquier campo ---
    def auto_guardar(*_args):
        try:
            config_dict = {}
            for k, var in config_vars.items():
                if isinstance(var, tk.BooleanVar):
                    config_dict[k] = 'true' if var.get() else 'false'
                else:
                    if k in ['separador_prefijo', 'separador_carpeta_combinada']:
                        sel = var.get()
                        # Mapear 'espacio en blanco' a ' '
                        config_dict[k] = sep_display_to_value.get(sel, ' ' if sel == '' else sel)
                    else:
                        config_dict[k] = var.get()
            guardar_config(config_dict)
        except Exception:
            # Evitar que cualquier error interrumpa la edición de la UI
            pass

    def _on_var_change(*_):
        auto_guardar()

    # Conectar todas las variables para que guarden al cambiar
    for _k, _var in config_vars.items():
        try:
            _var.trace_add('write', _on_var_change)
        except Exception:
            pass
    # Guardar una vez el estado inicial (normaliza separadores si aplica)
    auto_guardar()
    # --- Fin Autoguardado ---

    def on_iniciar():
        # Validar campos vacíos
        for k, var in config_vars.items():
            if k in ['eliminar_csv_al_final', 'usar_prefijo_columna', 'usar_carpeta_combinada']:
                continue
            # Solo validar los campos de prefijo/sufijo si está activado
            if k in ['columna_prefijo', 'tipo_prefijo', 'Nombre_de_la_carpeta', 'separador_prefijo'] and not usar_prefijo_var.get():
                continue
            # Solo validar los campos de carpeta combinada si está activado
            if k in ['columna_carpeta_1', 'columna_carpeta_2', 'separador_carpeta_combinada'] and not usar_carpeta_combinada_var.get():
                continue
            if isinstance(var, tk.BooleanVar):
                continue
            # Para los separadores, si están vacíos, no mostrar error y usar espacio en blanco
            if k in ['separador_prefijo', 'separador_carpeta_combinada'] and var.get() == '':
                continue
            if var.get().strip() == '':
                messagebox.showerror('Error', f'El campo "{k}" no puede estar vacío.')
                return
        # Guardar y cerrar interfaz
        config_dict = {}
        for k, var in config_vars.items():
            # Guardar todos los campos, incluyendo los booleanos
            if isinstance(var, tk.BooleanVar):
                config_dict[k] = 'true' if var.get() else 'false'
            else:
                if k in ['separador_prefijo', 'separador_carpeta_combinada']:
                    sel = var.get()
                    # Mapear 'espacio en blanco' a ' '
                    config_dict[k] = sep_display_to_value.get(sel, ' ' if sel == '' else sel)
                else:
                    config_dict[k] = var.get()
        guardar_config(config_dict)
        ejecutar_script()
        root.destroy()

    btn = tk.Button(root, text='INICIAR', command=on_iniciar, bg='green', fg='white', font=('Arial', 12, 'bold'))
    btn.grid(row=row, column=0, columnspan=2, pady=(15, 5))

    def on_reset():
        if not messagebox.askyesno('Confirmar', '¿Restablecer la configuración a los valores predeterminados?\nEsto sobrescribirá el archivo config.txt.'):
            return
        defaults = get_default_config()
        # Guardar inmediatamente en disco con defaults
        guardar_config(defaults)
        # Recargar valores en la UI
        for k, var in config_vars.items():
            try:
                val = defaults.get(k, '')
                if isinstance(var, tk.BooleanVar):
                    var.set(str(val).strip().lower() == 'true')
                else:
                    if k in ['separador_prefijo', 'separador_carpeta_combinada']:
                        # Convertir valor a display ('espacio en blanco' si corresponde)
                        display_val = {'_': '_', '-': '-', ' ': 'espacio en blanco', '': 'espacio en blanco'}.get(str(val), 'espacio en blanco' if str(val).strip() == '' else str(val))
                        var.set(display_val)
                    else:
                        var.set(val)
            except Exception:
                pass
        # Refrescar visibilidad acorde a los toggles
        try:
            actualizar_visibilidad()
        except Exception:
            pass
        messagebox.showinfo('Restablecido', 'La configuración ha sido restablecida y guardada en config.txt.')

    # Botón de restablecer
    row += 1
    tk.Button(root, text='Restablecer configuración', command=on_reset, bg='#cc3333', fg='white').grid(row=row, column=0, columnspan=2, pady=(0, 8))
    # Etiqueta de versión al pie
    row += 1
    tk.Label(root, text=f"Versión del script: {SCRIPT_VERSION}", fg='gray').grid(row=row, column=0, columnspan=2, pady=(0, 10))

    def on_close():
        # Guardar configuración antes de cerrar
        auto_guardar()
        root.destroy()
        os._exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)
    # Recalcular tamaño final para que quepan todos los campos
    root.update_idletasks()
    req_w = max(560, root.winfo_reqwidth() + 40)
    max_h = int(root.winfo_screenheight() * 0.9)
    req_h = min(max(420, root.winfo_reqheight() + 40), max_h)
    x = (root.winfo_screenwidth() // 2) - (req_w // 2)
    y = (root.winfo_screenheight() // 2) - (req_h // 2)
    root.geometry(f"{req_w}x{req_h}+{x}+{y}")
    root.mainloop()

# Si se ejecuta con argumento 'run', no mostrar la interfaz, solo ejecutar el script real
def procesar_csvs():
    r"""Procesa los CSV y descarga los archivos enlazados.

     Flujo principal:
     1) Lee `config.txt` para saber dónde están los CSV y cómo construir la
         estructura de carpetas destino (prefijo/sufijo y/o carpeta combinada).
     2) Abre cada CSV (separador ';', encoding latin-1) y por cada fila extrae
         un enlace (URL directa o HTML con un href).
     3) Descarga el recurso. Si es exitoso (HTTP 200 con contenido), guarda PDF;
         en caso contrario, genera un TXT con el enlace.

     Robustez añadida:
     - Se sanea el "stem" del nombre de archivo derivado de la URL para evitar
        caracteres inválidos en Windows (p.ej. ?, *, :, \, /, ", <, >, |, &).
     - Se descarta la parte de consulta (query) de la URL para el nombre visible,
        añadiendo un hash corto del query cuando exista, para diferenciar recursos.
     - Se manejan nombres reservados (CON, PRN, AUX, NUL, COM1.., LPT1..).
    - Se garantiza que los separadores vacíos equivalen a un espacio.
    """
    import os
    import re
    import requests
    import pandas as pd
    import urllib3
    import glob
    from urllib.parse import urlparse, unquote
    import hashlib
    import time
    from collections import deque
    # Desactivar advertencias SSL
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    # Leer configuración desde config.txt (crea defaults si no existe)
    config = leer_config()

    # Asignar variables desde config
    carpeta_descargas = config.get("carpeta_descargas", "descargas")
    csv_folder = config.get("csv_folder", "csvs")
    col_enlace = config.get("col_enlace")
    eliminar_csv_al_final = config.get("eliminar_csv_al_final", "false").lower() == "true"
    usar_prefijo_columna = config.get("usar_prefijo_columna", "false").lower() == "true"
    columna_prefijo = config.get("columna_prefijo", "")
    tipo_prefijo = config.get("tipo_prefijo", "prefijo").strip().lower()
    nombre_carpeta = config.get("Nombre_de_la_carpeta", "")
    usar_carpeta_combinada = config.get("usar_carpeta_combinada", "false").lower() == "true"
    columna_carpeta_1 = config.get("columna_carpeta_1", "")
    columna_carpeta_2 = config.get("columna_carpeta_2", "")
    separador_carpeta_combinada = config.get("separador_carpeta_combinada", "_")
    separador_prefijo = config.get("separador_prefijo", "_")
    # Si algún separador está vacío, usar espacio en blanco por defecto
    if separador_carpeta_combinada == "":
        separador_carpeta_combinada = " "
    if separador_prefijo == "":
        separador_prefijo = " "
    # Soportar valor textual desde configuraciones antiguas
    if str(separador_carpeta_combinada).strip().lower() == 'espacio en blanco':
        separador_carpeta_combinada = ' '
    if str(separador_prefijo).strip().lower() == 'espacio en blanco':
        separador_prefijo = ' '

    # --- Configurar logging a archivo .log en carpeta definida por el usuario ---
    # Se hace un "tee" de stdout y stderr hacia un archivo de log para registrar toda la salida.
    try:
        import io
        from datetime import datetime
        import atexit
        # Carpeta de logs desde config (por defecto 'logs')
        carpeta_logs = config.get('carpeta_logs', 'logs') or 'logs'
        try:
            os.makedirs(carpeta_logs, exist_ok=True)
        except Exception:
            # Si no se puede crear, caer a carpeta actual
            carpeta_logs = '.'
        nombre_log = f"TransparenciaActiva_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        ruta_log = os.path.join(carpeta_logs, nombre_log)

        class _StreamTee(io.TextIOBase):
            def __init__(self, stream, log_handle):
                self._stream = stream
                self._log = log_handle
            def write(self, s):
                # Garantiza que se escriba en consola y en archivo; ignora fallos individuales
                try:
                    self._stream.write(s)
                except Exception:
                    pass
                try:
                    self._log.write(s)
                except Exception:
                    pass
                return len(s)
            def flush(self):
                try:
                    self._stream.flush()
                except Exception:
                    pass
                try:
                    self._log.flush()
                except Exception:
                    pass
            def isatty(self):
                try:
                    return self._stream.isatty()
                except Exception:
                    return False

        _log_handle = open(ruta_log, 'a', encoding='utf-8', buffering=1)  # line buffered
        sys.stdout = _StreamTee(sys.stdout, _log_handle)
        sys.stderr = _StreamTee(sys.stderr, _log_handle)
        atexit.register(lambda: (_log_handle.flush(), _log_handle.close()))
        print(f"Registro de ejecución: {os.path.abspath(ruta_log)}")
    except Exception:
        # Si algo falla al configurar el log, continuar sin interrumpir la ejecución
        pass

    # Mostrar banner de ms.marcomputer y versión del script (no interferir con ejecución)
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        banner_path = os.path.join(base_dir, 'ms.marcomputer')
        banner_text = ''
        try:
            with open(banner_path, 'r', encoding='utf-8', errors='replace') as bf:
                banner_text = bf.read()
        except Exception:
            banner_text = ''
        if banner_text:
            print("\n" + banner_text.rstrip('\n'))
        print(f"Versión del script: {SCRIPT_VERSION}")
        print("#" * 60)
    except Exception:
        # No permitir que el banner o su impresión corten la ejecución
        pass

    # Crear carpeta de descarga si no existe
    os.makedirs(carpeta_descargas, exist_ok=True)

    # Expresión regular para extraer href
    href_pattern = re.compile(r"href=['\"](.*?)['\"]", re.IGNORECASE)

    # Obtener lista de archivos CSV desde la carpeta configurada
    csv_files = glob.glob(os.path.join(csv_folder, "*.csv"))

    # --- Utilidades ---
    def human_size(num_bytes: int) -> str:
        """Devuelve el tamaño en MB si >= 1MB; si no, en KB (2 decimales)."""
        try:
            if num_bytes is None:
                return "N/A"
            if num_bytes >= 1024 * 1024:
                return f"{num_bytes / (1024*1024):.2f} MB"
            return f"{num_bytes / 1024:.2f} KB"
        except Exception:
            return "N/A"

    def human_size_summary(num_bytes: int) -> str:
        """Devuelve un tamaño legible para resumen: GB, MB o KB (2 decimales)."""
        try:
            if num_bytes is None:
                return "N/A"
            if num_bytes >= 1024 * 1024 * 1024:  # >= 1 GB
                return f"{num_bytes / (1024*1024*1024):.2f} GB"
            if num_bytes >= 1024 * 1024:  # >= 1 MB
                return f"{num_bytes / (1024*1024):.2f} MB"
            return f"{num_bytes / 1024:.2f} KB"
        except Exception:
            return "N/A"

    def _format_seconds(seg: int) -> str:
        """Formatea segundos a HH:MM:SS o MM:SS si < 1 hora."""
        try:
            seg = int(max(0, seg))
            h, rem = divmod(seg, 3600)
            m, s = divmod(rem, 60)
            if h > 0:
                return f"{h:02d}:{m:02d}:{s:02d}"
            return f"{m:02d}:{s:02d}"
        except Exception:
            return "--:--"

    def _progress_bar(current: int, total: int, width: int = 30) -> str:
        """Barra de progreso ASCII.

        Ej: [##########--------------------]
        """
        try:
            if total <= 0:
                return "[" + ("-" * width) + "]"
            ratio = min(max(current / total, 0.0), 1.0)
            filled = int(round(width * ratio))
            filled = min(filled, width)
            return "[" + ("#" * filled) + ("-" * (width - filled)) + "]"
        except Exception:
            return "[" + ("-" * width) + "]"

    # --- Análisis inicial: contar descargas a intentar ---
    total_intentos = 0
    try:
        for _csv in csv_files:
            try:
                _df = pd.read_csv(_csv, encoding='latin-1', sep=';', quotechar='"')
                try:
                    _df.columns = _df.columns.str.strip()
                except Exception:
                    pass
                _col_enlace_key = (col_enlace or '').strip()
                if _col_enlace_key in _df.columns:
                    _col_enlace_resolved = _col_enlace_key
                else:
                    _lower_map = {c.lower(): c for c in _df.columns}
                    _col_enlace_resolved = _lower_map.get(_col_enlace_key.lower(), '')
                if not _col_enlace_resolved:
                    continue
                for __, _row in _df.iterrows():
                    _val = str(_row.get(_col_enlace_resolved, ''))
                    if _val.startswith('http://') or _val.startswith('https://') or href_pattern.search(_val):
                        total_intentos += 1
            except Exception:
                # Si un CSV falla, continuar con el siguiente
                pass
        print("\n" + "#"*60)
        print(f"ANÁLISIS INICIAL: Se intentarán {total_intentos} descargas en {len(csv_files)} archivo(s) CSV.")
        print("#"*60 + "\n")
    except Exception:
        # Si algo falla, continuar sin bloquear la ejecución
        pass

    def obtener_nombre_unico(ruta_base, extension):
        """Dada una ruta base sin extensión y una extensión, devuelve una ruta
        única que no exista en disco, añadiendo sufijos -2, -3, ... si hace falta."""
        contador = 1
        ruta_final = f"{ruta_base}.{extension}"
        while os.path.exists(ruta_final):
            contador += 1
            ruta_final = f"{ruta_base}-{contador}.{extension}"
        return ruta_final

    def _safe_folder_name(name: str) -> str:
        r"""Devuelve un nombre de carpeta seguro para Windows.

        - Reemplaza caracteres inválidos (< > : " / \ | ? *) por '_'.
        - Elimina espacios al inicio/fin.
        - No intenta recortar longitud por ser parte de la ruta, pero al
          combinarse con otras partes, Windows impone un límite de 260 chars.
    """
        # Reemplazar caracteres inválidos en Windows y limpiar espacios extremos
        invalid = '<>:"/\\|?*'
        for ch in invalid:
            name = name.replace(ch, '_')
        return name.strip()

    def _safe_file_stem(name: str) -> str:
        """Devuelve un nombre de archivo (sin extensión) seguro para Windows.

        - Reemplaza caracteres inválidos por '_'.
        - Quita caracteres de control y puntos finales.
        - Evita nombres reservados (CON, PRN, AUX, NUL, COM1.., LPT1..).
        - Limita la longitud para evitar rutas excesivas.
        """
        # Reemplazar caracteres inválidos explícitos
        invalid = '<>:"/\\|?*'
        name = ''.join('_' if ch in invalid else ch for ch in name)
        # Quitar caracteres de control (0-31) y normalizar otros no ASCII visibles
        name = ''.join(ch if 32 <= ord(ch) < 127 else '_' for ch in name)
        # Quitar espacios extremos y puntos al final (Windows no permite)
        name = name.strip().rstrip('.')
        # Evitar nombres reservados de Windows (case-insensitive)
        reserved = {"CON","PRN","AUX","NUL"} | {f"COM{i}" for i in range(1,10)} | {f"LPT{i}" for i in range(1,10)}
        if not name or name.upper() in reserved:
            name = f"archivo_{int(time.time())}"
        # Limitar longitud del "stem" para prevenir rutas demasiado largas
        return name[:150]

    def _url_to_safe_stem(url: str) -> str:
        """Convierte una URL en un "stem" de nombre de archivo seguro.

        - Usa sólo el último segmento de la ruta (sin la extensión).
        - Descarta el query visible pero añade un hash corto del query (8 hex)
          para evitar colisiones cuando distintas URLs comparten el mismo path.
        - Sanea el resultado con `_safe_file_stem`.
        """
        try:
            parsed = urlparse(url)
            # Último segmento de la ruta (decodificado), sin extensión
            last_seg = os.path.basename(unquote(parsed.path))
            stem, _ext = os.path.splitext(last_seg)
        except Exception:
            stem = ''
            parsed = None
        if not stem:
            stem = 'archivo'
        # Si hay query, agregar huella breve para distinguir
        try:
            q = (parsed.query if parsed else '') or ''
            if q:
                digest = hashlib.sha1(q.encode('utf-8')).hexdigest()[:8]
                stem = f"{stem}_{digest}"
        except Exception:
            pass
        return _safe_file_stem(stem)

    # Normalizar nombres de columnas (el usuario puede haber dejado espacios al final)
    columna_prefijo_key = columna_prefijo.strip()
    columna_carpeta_1_key = columna_carpeta_1.strip()
    columna_carpeta_2_key = columna_carpeta_2.strip()

    # También normalizar clave de enlace
    col_enlace_key = (col_enlace or '').strip()

    def resolve_col(df, configured_name: str) -> str:
        """Dada una columna configurada, devolver el nombre real de la columna en df
        aplicando strip y case-insensitive. Si no se encuentra, retornar ''."""
        name = (configured_name or '').strip()
        if not name:
            return ''
        if name in df.columns:
            return name
        lower_map = {c.lower(): c for c in df.columns}
        return lower_map.get(name.lower(), '')

    # Contadores de resumen
    total_archivos = 0
    total_pdfs = 0
    total_txts = 0
    errores = 0
    total_bytes_descargados = 0

    # Índice de progreso (solo filas con URL válida)
    descarga_idx = 0
    # Ventana móvil para ETA (últimas 10 descargas)
    _times_window = deque(maxlen=10)

    for csv_file in csv_files:
        print(f"Procesando: {csv_file}")
        df = pd.read_csv(csv_file, encoding='latin-1', sep=';', quotechar='"')
        # Normalizar encabezados para evitar fallos por espacios
        try:
            df.columns = df.columns.str.strip()
        except Exception:
            pass
        # Resolver nombres de columnas según df
        col_enlace_resolved = resolve_col(df, col_enlace_key)
        prefijo_col_resolved = resolve_col(df, columna_prefijo_key)
        comb1_col_resolved = resolve_col(df, columna_carpeta_1_key)
        comb2_col_resolved = resolve_col(df, columna_carpeta_2_key)

        for _, row in df.iterrows():
            total_archivos += 1
            if not col_enlace_resolved:
                print(f"ERROR: La columna de enlace configurada ('{col_enlace_key}') no existe en el CSV.")
                errores += 1
                break
            html = str(row[col_enlace_resolved])
            # Determinar carpeta destino
            carpeta_destino = carpeta_descargas
            # Construcción de carpetas: si ambas opciones están activadas, anidar combinada dentro de la de prefijo/sufijo
            if usar_prefijo_columna:
                raw_val = row.get(prefijo_col_resolved, '') if prefijo_col_resolved else ''
                # Tratar NaN como vacío
                try:
                    import pandas as _pd
                    if _pd.isna(raw_val):
                        raw_val = ''
                except Exception:
                    pass
                valor_col_prefijo = str(raw_val).strip() if raw_val is not None else ''

                base_nombre = nombre_carpeta.strip()
                nombre_carpeta_prefijo = ''
                if valor_col_prefijo and base_nombre:
                    if tipo_prefijo == 'prefijo':
                        nombre_carpeta_prefijo = f"{valor_col_prefijo}{separador_prefijo}{base_nombre}"
                    else:
                        nombre_carpeta_prefijo = f"{base_nombre}{separador_prefijo}{valor_col_prefijo}"
                elif valor_col_prefijo:
                    # Solo valor de la columna
                    nombre_carpeta_prefijo = valor_col_prefijo
                elif base_nombre:
                    # Fallback: solo 'Nombre de la carpeta'
                    nombre_carpeta_prefijo = base_nombre

                if nombre_carpeta_prefijo:
                    nombre_carpeta_prefijo = _safe_folder_name(nombre_carpeta_prefijo)
                    carpeta_prefijo_path = os.path.join(carpeta_descargas, nombre_carpeta_prefijo)
                    os.makedirs(carpeta_prefijo_path, exist_ok=True)
                else:
                    carpeta_prefijo_path = None
            else:
                carpeta_prefijo_path = None

            if usar_carpeta_combinada:
                val1 = str(row.get(comb1_col_resolved, '')).strip() if comb1_col_resolved else ''
                val2 = str(row.get(comb2_col_resolved, '')).strip() if comb2_col_resolved else ''
                if val1 and val2:
                    nombre_carpeta_combinada = _safe_folder_name(f"{val1}{separador_carpeta_combinada}{val2}")
                    parent = carpeta_prefijo_path if carpeta_prefijo_path else carpeta_descargas
                    carpeta_destino = os.path.join(parent, nombre_carpeta_combinada)
                    os.makedirs(carpeta_destino, exist_ok=True)
                else:
                    carpeta_destino = carpeta_prefijo_path if carpeta_prefijo_path else carpeta_descargas
            else:
                # Solo prefijo/sufijo
                carpeta_destino = carpeta_prefijo_path if carpeta_prefijo_path else carpeta_descargas
            # Obtener URL
            if html.startswith('http://') or html.startswith('https://'):
                url = html
            else:
                match = href_pattern.search(html)
                if match:
                    url = match.group(1)
                else:
                    url = None
            if url:
                # Actualizar progreso
                descarga_idx += 1
                percent_done = int(round((descarga_idx / total_intentos) * 100)) if total_intentos else 0
                quedan = max(total_intentos - descarga_idx, 0)
                # Obtener un nombre base seguro derivado de la URL
                nombre_base_seguro = _url_to_safe_stem(url)
                ruta_base = os.path.join(carpeta_destino, nombre_base_seguro)
                ruta_archivo = obtener_nombre_unico(ruta_base, "pdf")
                ruta_txt = obtener_nombre_unico(ruta_base, "txt")
                try:
                    _t0 = time.time()
                    response = requests.get(url, verify=False)
                    _elapsed = time.time() - _t0
                    try:
                        _times_window.append(_elapsed)
                    except Exception:
                        pass
                    _avg = (sum(_times_window) / len(_times_window)) if len(_times_window) else 0
                    _eta_secs = int(round(quedan * _avg)) if _avg > 0 else 0
                    _bar = _progress_bar(descarga_idx, total_intentos, 30)
                    # Mostrar como tabla con progreso y tamaños legibles
                    tam_str = human_size(len(response.content) if response.content is not None else 0)
                    encabezado = f" {descarga_idx} / {total_intentos} - {percent_done}% {_bar} Tiempo restante: {_format_seconds(_eta_secs)}\n"
                    tabla = "\n" + encabezado
                    tabla += "="*60 + "\n"
                    tabla += f"| {'Campo':<20} | {'Valor':<35} |\n"
                    tabla += f"|{'-'*20}|{'-'*35}|\n"
                    tabla += f"| {'Archivo':<20} | {ruta_archivo:<35} |\n"
                    tabla += f"| {'Enlace':<20} | {url:<35} |\n"
                    tabla += f"| {'Status HTTP':<20} | {response.status_code:<35} |\n"
                    tabla += f"| {'Tamaño recibido':<20} | {tam_str:<35} |\n"
                    tabla += f"| {'Tipo de contenido':<20} | {response.headers.get('Content-Type', 'N/A'):<35} |\n"
                    tabla += f"| {'Quedan':<20} | {quedan:<35} |\n"
                    if response.status_code == 200 and response.content:
                        with open(ruta_archivo, "wb") as f:
                            f.write(response.content)
                        tabla += f"| {'Resultado':<20} | {'PDF descargado correctamente.':<35} |\n"
                        try:
                            total_bytes_descargados += len(response.content)
                        except Exception:
                            pass
                        total_pdfs += 1
                    else:
                        with open(ruta_txt, "w", encoding="utf-8") as f:
                            f.write(url)
                        tabla += f"| {'Resultado':<20} | {'No se pudo descargar el PDF. Se creó el TXT con el enlace.':<35} |\n"
                        total_txts += 1
                    tabla += "="*60 + "\n"
                    print(tabla)
                except Exception as e:
                    # Intentar guardar TXT con el enlace
                    txt_creado = False
                    _elapsed = time.time() - _t0 if '_t0' in locals() else 0
                    try:
                        _times_window.append(_elapsed)
                    except Exception:
                        pass
                    _avg = (sum(_times_window) / len(_times_window)) if len(_times_window) else 0
                    _eta_secs = int(round(quedan * _avg)) if _avg > 0 else 0
                    _bar = _progress_bar(descarga_idx, total_intentos, 30)
                    try:
                        with open(ruta_txt, "w", encoding="utf-8") as f:
                            f.write(url or '')
                        txt_creado = True
                    except Exception:
                        pass
                    # Formatear error en el mismo cuadro
                    encabezado = f" {descarga_idx} / {total_intentos} - {percent_done}% {_bar} Tiempo restante: {_format_seconds(_eta_secs)}\n"
                    tabla = "\n" + encabezado
                    tabla += "="*60 + "\n"
                    tabla += f"| {'Campo':<20} | {'Valor':<35} |\n"
                    tabla += f"|{'-'*20}|{'-'*35}|\n"
                    tabla += f"| {'Archivo TXT':<20} | {ruta_txt if txt_creado else 'No creado':<35} |\n"
                    tabla += f"| {'Enlace':<20} | {url or 'N/A':<35} |\n"
                    tabla += f"| {'Tamaño recibido':<20} | {'0 KB':<35} |\n"
                    tabla += f"| {'Quedan':<20} | {quedan:<35} |\n"
                    tabla += f"| {'Resultado':<20} | {'ERROR en la descarga':<35} |\n"
                    tabla += f"| {'Detalle error':<20} | {str(e):<35} |\n"
                    tabla += "="*60 + "\n"
                    print(tabla)
                    total_txts += 1
                    errores += 1
            else:
                print(f"No se pudo obtener una URL para descargar en la fila: {html}")
                errores += 1

    # Mostrar resumen final
    print("\n" + "#"*60)
    print("RESUMEN DE DESCARGA")
    # Métricas de avance
    try:
        percent_complete = int(round((descarga_idx / total_intentos) * 100)) if total_intentos else 0
    except Exception:
        percent_complete = 0
    # Barra y ETA finales
    try:
        _avg_final = (sum(_times_window) / len(_times_window)) if len(_times_window) else 0
    except Exception:
        _avg_final = 0
    _quedan_final = max(total_intentos - descarga_idx, 0)
    _eta_final = int(round(_quedan_final * _avg_final)) if _avg_final > 0 else 0
    try:
        _bar_final = _progress_bar(descarga_idx, total_intentos, 30)
    except Exception:
        _bar_final = "[------------------------------]"
    print(f"Descargas detectadas (análisis): {total_intentos}")
    print(f"Descargas intentadas: {descarga_idx}")
    print(f"Completado: {percent_complete}%")
    print(f"Progreso final: {descarga_idx} / {total_intentos} - {percent_complete}% {_bar_final} Tiempo restante: {_format_seconds(_eta_final)}")
    print(f"Total descargado: {human_size_summary(total_bytes_descargados)}")
    print(f"Total de archivos procesados: {total_archivos}")
    print(f"PDFs descargados correctamente: {total_pdfs}")
    print(f"Archivos TXT generados: {total_txts}")
    print(f"Errores encontrados: {errores}")
    print("#"*60 + "\n")

    input("Presiona ENTER para cerrar la ventana...")

    # Eliminar los archivos CSV si la opción está activada
    if eliminar_csv_al_final:
        for csv_file in csv_files:
            try:
                os.remove(csv_file)
                print(f"Eliminado CSV: {csv_file}")
            except Exception as e:
                print(f"Error al eliminar {csv_file}: {e}")

import sys
if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'run':
        print('Ejecutando el script principal...')
        procesar_csvs()
    else:
        mostrar_interfaz()
