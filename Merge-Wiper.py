import os
import sys
import logging
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from colorama import init, Fore, Style
# Importa módulos para el navegador de archivos (interfaz gráfica)
import tkinter as tk
from tkinter import filedialog
# Importa módulos para reportes detallados (medición de recursos)
import time
import psutil

# Inicializa colorama para colorear la salida en consola automáticamente
init(autoreset=True)

# Configuración del sistema de logging para mostrar mensajes de depuración, información y errores
LOG_LEVEL = logging.DEBUG  # Puedes cambiar a logging.INFO o logging.ERROR según lo que necesites ver
logging.basicConfig(
    level=LOG_LEVEL,
    format='[%(levelname)s] %(message)s'
)

def print_separator():
    """
    Imprime una línea separadora amarilla en la consola para mejorar la legibilidad.
    """
    print(Fore.YELLOW + '-' * 60)

def print_menu_title(title):
    """
    Imprime un título de menú resaltado con color y separadores.
    """
    print_separator()
    print(Fore.CYAN + Style.BRIGHT + title)
    print_separator()

def interactive_folder_selection(start_path=None):
    """
    Permite al usuario navegar por las carpetas de manera interactiva desde la consola.
    Devuelve la ruta seleccionada por el usuario.
    """
    if start_path is None:
        start_path = os.getcwd()  # Usa la carpeta actual si no se especifica otra
    current_path = os.path.abspath(start_path)
    while True:
        print(Fore.CYAN + f"\nCarpeta actual: {current_path}")
        # Lista solo las subcarpetas de la carpeta actual
        entries = [e for e in os.listdir(current_path) if os.path.isdir(os.path.join(current_path, e))]
        print(Fore.YELLOW + "Subcarpetas:")
        for idx, entry in enumerate(entries):
            print(f"  {idx+1}. {entry}")
        print(Fore.MAGENTA + "0. Seleccionar esta carpeta")
        print(Fore.MAGENTA + "u. Subir un nivel")
        # Solicita al usuario que seleccione una opción
        sel = input(Fore.WHITE + "Seleccione una opción (número, 0 para seleccionar, u para subir): ").strip()
        if sel == '0':
            return current_path  # Devuelve la carpeta actual
        elif sel.lower() == 'u':
            parent = os.path.dirname(current_path)
            if parent == current_path:
                print(Fore.RED + "Ya está en la carpeta raíz.")
            else:
                current_path = parent  # Sube un nivel en la jerarquía de carpetas
        elif sel.isdigit():
            idx = int(sel) - 1
            if 0 <= idx < len(entries):
                current_path = os.path.join(current_path, entries[idx])  # Entra en la subcarpeta seleccionada
            else:
                print(Fore.RED + "Índice fuera de rango.")
        else:
            print(Fore.RED + "Opción inválida.")

def ask_file_paths(prompt_msg, multiple=False, allow_file_dialog=False):
    """
    Solicita al usuario las rutas de los archivos XLSX a procesar.
    Permite seleccionar uno o varios archivos, ya sea por consola o usando un navegador de archivos gráfico.
    """
    print(Fore.GREEN + prompt_msg)
    files = []
    if allow_file_dialog:
        print(Fore.YELLOW + "¿Cómo desea seleccionar los archivos?")
        print(Fore.MAGENTA + "1. Seleccionar desde una carpeta (navegación interactiva)")
        print(Fore.MAGENTA + "2. Seleccionar usando el navegador de archivos (permite selección múltiple)")
        while True:
            sel = input(Fore.WHITE + "Seleccione una opción (1 o 2): ").strip()
            if sel == '1':
                # Permite navegar por carpetas desde la consola
                folder = interactive_folder_selection()
                # Busca archivos .xlsx en la carpeta seleccionada
                files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.xlsx') and os.path.isfile(os.path.join(folder, f))]
                if not files:
                    print(Fore.RED + "No se encontraron archivos .xlsx en la carpeta.")
                    return []
                break
            elif sel == '2':
                # Abre un navegador de archivos gráfico para seleccionar archivos
                root = tk.Tk()
                root.withdraw()  # Oculta la ventana principal de Tkinter
                if multiple:
                    files = filedialog.askopenfilenames(
                        title="Selecciona archivos XLSX",
                        filetypes=[("Archivos Excel", "*.xlsx")]
                    )
                    files = list(files)
                else:
                    file = filedialog.askopenfilename(
                        title="Selecciona un archivo XLSX",
                        filetypes=[("Archivos Excel", "*.xlsx")]
                    )
                    files = [file] if file else []
                root.destroy()
                if not files:
                    print(Fore.RED + "No se seleccionaron archivos.")
                    return []
                break
            else:
                print(Fore.RED + "Opción inválida.")
        return files
    # Selección manual por consola
    while True:
        path = input(Fore.WHITE + "Ruta de archivo (deja vacío para terminar): ").strip()
        if not path:
            if multiple and files:
                break
            elif not multiple:
                print(Fore.RED + "Debes ingresar al menos un archivo.")
                continue
            else:
                print(Fore.RED + "No se seleccionó ningún archivo.")
                return []
        if not os.path.isfile(path):
            print(Fore.RED + f"Archivo no encontrado: {path}")
            continue
        if not path.lower().endswith('.xlsx'):
            print(Fore.RED + "Solo se permiten archivos .xlsx")
            continue
        files.append(path)
        if not multiple:
            break
    return files

def ask_output_path(default_name):
    """
    Solicita al usuario la carpeta y el nombre del archivo de salida.
    Si el usuario no ingresa nada, usa la carpeta actual y un nombre por defecto.
    """
    print(Fore.GREEN + "¿Dónde desea guardar el archivo de salida?")
    folder = input(Fore.WHITE + "Ruta de carpeta (deja vacío para usar la carpeta actual): ").strip()
    if not folder:
        folder = os.getcwd()
    if not os.path.isdir(folder):
        print(Fore.RED + "La carpeta no existe. Usando carpeta actual.")
        folder = os.getcwd()
    name = input(Fore.WHITE + f"Nombre del archivo de salida (sin extensión, por defecto '{default_name}'): ").strip()
    if not name:
        name = default_name
    return os.path.join(folder, name + '.xlsx')

def get_headers(ws):
    """
    Obtiene los encabezados (primera fila) de una hoja de Excel.
    """
    return [cell.value for cell in ws[1]]

def check_same_headers(file_paths):
    """
    Verifica que todos los archivos tengan los mismos encabezados.
    Si encuentra diferencias, muestra un error y retorna False.
    """
    headers = None
    for path in file_paths:
        wb = load_workbook(path, read_only=True)
        ws = wb.active
        current_headers = get_headers(ws)
        if headers is None:
            headers = current_headers
        elif headers != current_headers:
            logging.error(f"Encabezados diferentes en el archivo: {path}")
            return False
        wb.close()
    return True

# --- SECCIÓN DE REPORTES DETALLADOS ---
def print_report(report_data):
    """
    Imprime un reporte detallado de la operación realizada, incluyendo estadísticas y uso de recursos.
    """
    print_separator()
    print(Fore.GREEN + Style.BRIGHT + "REPORTE DE OPERACIÓN")
    print_separator()
    print(Fore.CYAN + f"Archivos procesados: {report_data.get('files_processed', '-')}")
    # Mostrar líneas por archivo si está disponible
    file_lines = report_data.get('file_lines', None)
    if file_lines:
        print(Fore.BLUE + "Líneas procesadas por archivo:")
        for fname, lines in file_lines.items():
            print(Fore.BLUE + f"  {os.path.basename(fname)} - {lines} líneas")
    # Mostrar líneas al inicio y al final con textos personalizados
    print(Fore.CYAN + f"Líneas al inicio: {report_data.get('lines_in_text', '-')}")
    print(Fore.CYAN + f"Líneas al final: {report_data.get('lines_out_text', '-')}")
    print(Fore.CYAN + f"Tiempo de operación: {report_data.get('duration', '-')} segundos")
    print(Fore.CYAN + f"Tamaño archivo de salida: {report_data.get('output_size_kb', '-')} KB")
    print(Fore.CYAN + f"Archivo de salida: {report_data.get('output_path', '-')}")
    print(Fore.CYAN + f"Carpeta de destino: {report_data.get('output_folder', '-')}")
    print(Fore.CYAN + f"RAM usada: {report_data.get('ram_used_mb', '-')} MB")
    print(Fore.CYAN + f"CPU usada: {report_data.get('cpu_percent', '-')} %")
    print_separator()
# --- FIN SECCIÓN DE REPORTES DETALLADOS ---

def merge_xlsx():
    """
    Función principal para combinar (merge) varios archivos XLSX en uno solo.
    - Solicita al usuario los archivos a combinar.
    - Verifica que tengan la misma estructura.
    - Copia los datos de todos los archivos en uno nuevo.
    - Mide recursos y tiempo, y muestra un reporte detallado.
    """
    print_menu_title("Función MERGE - Consolidar archivos XLSX")
    files = ask_file_paths(
        "Selecciona los archivos XLSX a combinar (misma estructura):",
        multiple=True,
        allow_file_dialog=True
    )
    if not files:
        print(Fore.RED + "No se seleccionaron archivos.")
        return
    if not check_same_headers(files):
        print(Fore.RED + "Los archivos no tienen los mismos encabezados. Operación cancelada.")
        return
    output_path = ask_output_path("merge_result")
    # --- INICIO MEDICIÓN DE RECURSOS Y TIEMPO ---
    start_time = time.time()
    process = psutil.Process(os.getpid())
    ram_before = process.memory_info().rss
    cpu_before = psutil.cpu_percent(interval=None)
    # ---
    try:
        wb_out = Workbook()  # Crea un nuevo libro de Excel para el resultado
        ws_out = wb_out.active
        # Copia el encabezado del primer archivo
        wb_first = load_workbook(files[0], read_only=True)
        ws_first = wb_first.active
        ws_out.append(get_headers(ws_first))
        wb_first.close()
        # Copia los datos de todos los archivos (ignorando la primera fila)
        total_lines_in = 0
        file_lines = {}
        for file in files:
            logging.debug(f"Procesando archivo: {file}")
            wb = load_workbook(file, read_only=True)
            ws = wb.active
            file_line_count = sum(1 for _ in ws.iter_rows(min_row=2, values_only=True))
            file_lines[file] = file_line_count
            total_lines_in += file_line_count
            # Volver a iterar para copiar las filas (no se puede reutilizar el iterador anterior)
            for row in ws.iter_rows(min_row=2, values_only=True):
                ws_out.append(row)
            wb.close()
        wb_out.save(output_path)  # Guarda el archivo combinado
        # --- FIN MEDICIÓN DE RECURSOS Y TIEMPO ---
        end_time = time.time()
        ram_after = process.memory_info().rss
        cpu_after = psutil.cpu_percent(interval=None)
        output_size_kb = os.path.getsize(output_path) // 1024
        output_folder = os.path.dirname(output_path)
        # Cuenta las filas finales (sin encabezado)
        wb_check = load_workbook(output_path, read_only=True)
        ws_check = wb_check.active
        lines_out = ws_check.max_row - 1  # sin encabezado
        wb_check.close()
        # Prepara los datos para el reporte detallado
        report_data = {
            'files_processed': len(files),
            'file_lines': file_lines,
            'lines_in': total_lines_in,
            'lines_in_text': f"{total_lines_in} Líneas Total",
            'lines_out': lines_out,
            'lines_out_text': f"{lines_out} Líneas Mergeadas",
            'duration': round(end_time - start_time, 2),
            'output_size_kb': output_size_kb,
            'output_path': output_path,
            'output_folder': output_folder,
            'ram_used_mb': round((ram_after - ram_before) / (1024*1024), 2),
            'cpu_percent': cpu_after
        }
        print(Fore.GREEN + f"Archivos combinados exitosamente en: {output_path}")
        print_report(report_data)
    except Exception as e:
        logging.error(f"Error al combinar archivos: {e}")

def wipe_xlsx():
    """
    Función principal para eliminar duplicados de un archivo XLSX.
    - Solicita al usuario el archivo a purgar.
    - Pide la columna que contiene los valores únicos.
    - Elimina filas duplicadas según esa columna.
    - Mide recursos y tiempo, y muestra un reporte detallado.
    """
    print_menu_title("Función WIPE - Eliminar duplicados en archivo XLSX")
    files = ask_file_paths(
        "Selecciona el archivo XLSX a purgar (solo uno):",
        multiple=False,
        allow_file_dialog=True
    )
    if not files:
        print(Fore.RED + "No se seleccionó archivo.")
        return
    file = files[0]
    # --- INICIO MEDICIÓN DE RECURSOS Y TIEMPO ---
    start_time = time.time()
    process = psutil.Process(os.getpid())
    ram_before = process.memory_info().rss
    cpu_before = psutil.cpu_percent(interval=None)
    # ---
    try:
        wb = load_workbook(file)
        ws = wb.active
        headers = get_headers(ws)
        # Muestra las columnas disponibles y sus encabezados
        print(Fore.BLUE + "Columnas disponibles:")
        for idx, header in enumerate(headers):
            col_letter = get_column_letter(idx + 1)
            print(f"  Columna {col_letter} - {header}")
        # Solicita al usuario la columna que contiene los valores únicos
        while True:
            col_input = input(Fore.WHITE + "¿Qué columna contiene los valores únicos? (Letra de columna): ").strip().upper()
            if not col_input:
                print(Fore.RED + "Debes ingresar una letra de columna.")
                continue
            try:
                col_idx = ord(col_input) - ord('A')
                if 0 <= col_idx < len(headers):
                    break
                else:
                    print(Fore.RED + "Columna fuera de rango.")
            except Exception:
                print(Fore.RED + "Entrada inválida.")
        unique_col = col_idx
        seen = set()  # Conjunto para almacenar los valores únicos encontrados
        rows_to_keep = []
        total_lines_in = ws.max_row - 1  # sin encabezado
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                rows_to_keep.append(row)  # Siempre guarda el encabezado
                continue
            val = row[unique_col]
            if val not in seen:
                seen.add(val)
                rows_to_keep.append(row)
        output_path = ask_output_path("wipe_result")
        wb_out = Workbook()
        ws_out = wb_out.active
        for row in rows_to_keep:
            ws_out.append(row)
        wb_out.save(output_path)
        # --- FIN MEDICIÓN DE RECURSOS Y TIEMPO ---
        end_time = time.time()
        ram_after = process.memory_info().rss
        cpu_after = psutil.cpu_percent(interval=None)
        output_size_kb = os.path.getsize(output_path) // 1024
        output_folder = os.path.dirname(output_path)
        lines_out = len(rows_to_keep) - 1  # sin encabezado
        lines_removed = total_lines_in - lines_out
        # Prepara los datos para el reporte detallado
        report_data = {
            'files_processed': 1,
            'file_lines': {file: total_lines_in},
            'lines_in': total_lines_in,
            'lines_in_text': f"{total_lines_in} Líneas Total",
            'lines_out': lines_out,
            'lines_out_text': f"{lines_out} Líneas después del Wipe - {lines_removed} líneas duplicadas han sido eliminadas/purgadas",
            'duration': round(end_time - start_time, 2),
            'output_size_kb': output_size_kb,
            'output_path': output_path,
            'output_folder': output_folder,
            'ram_used_mb': round((ram_after - ram_before) / (1024*1024), 2),
            'cpu_percent': cpu_after
        }
        print(Fore.GREEN + f"Archivo purgado exitosamente en: {output_path}")
        print_report(report_data)
    except Exception as e:
        logging.error(f"Error al purgar archivo: {e}")

def main_menu():
    """
    Muestra el menú principal del programa y gestiona la selección del usuario.
    Permite elegir entre combinar archivos, eliminar duplicados o salir.
    """
    while True:
        print_menu_title("MERGE-WIPER - Menú Principal")
        print(Fore.MAGENTA + "1. Merge (Combinar archivos XLSX)")
        print(Fore.MAGENTA + "2. Wipe (Eliminar duplicados en archivo XLSX)")
        print(Fore.MAGENTA + "3. Salir")
        choice = input(Fore.WHITE + "Seleccione una opción: ").strip()
        if choice == '1':
            merge_xlsx()
        elif choice == '2':
            wipe_xlsx()
        elif choice == '3':
            print(Fore.YELLOW + "¡Hasta luego!")
            break
        else:
            print(Fore.RED + "Opción inválida. Intente de nuevo.")

if __name__ == "__main__":
    # Punto de entrada principal del programa
    try:
        main_menu()
    except KeyboardInterrupt:
        print(Fore.YELLOW + "\nOperación cancelada por el usuario.")
