import pandas as pd
import openpyxl
import sys
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk  # Usar PIL para manejar imágenes
sys.setrecursionlimit(10000)

# Crear la ventana principal
window = tk.Tk()
window.title("Auto-Depurador")
window.geometry("500x350")
window.resizable(False, False)  # Evita que la ventana cambie de tamaño

# Declarar la barra de progreso globalmente
progress_bar = ttk.Progressbar(window, orient="horizontal", length=400, mode="determinate")
progress_bar.pack(pady=20)

# Crear la etiqueta para mostrar el contador de archivos procesados
counter_label = tk.Label(window, text="0 de 0")
counter_label.pack()

# Función para limpiar los datos
def clean_data(file_path):
    # Cargar y limpiar el archivo Excel
    if file_path.endswith(".XLS"):  # Si es un archivo .XLS
        df = pd.read_csv(file_path, delimiter="\t", encoding="utf-16", on_bad_lines="skip", skiprows=8)
        df = df.drop(df.columns[[0]], axis=1)
    else:  # Si es un archivo .XLSX
        df = pd.read_excel(file_path, engine="openpyxl")

    # Agrupar por 'Doc.ventas' y aplicar forward-fill y back-fill para rellenar valores nulos
    df1 = df.groupby('Doc.ventas').apply(lambda group: group.fillna(method='ffill').fillna(method='bfill'))
    
    # Formatear 'Val.neto factura' eliminando puntos y cambiando comas por puntos
    df1['Val.neto factura'] = df1['Val.neto factura'].str.replace('.', '', regex=False)
    df1['Val.neto factura'] = df1['Val.neto factura'].str.replace(',', '.', regex=False)
    
    # Convertir 'Val.neto factura' a numérico
    df1['Val.neto factura'] = pd.to_numeric(df1['Val.neto factura'], errors='coerce')
    
    # Eliminar filas con valores nulos en 'Doc.factura jurídico'
    df1 = df1.dropna(subset=['Doc.factura jurídico'])

    # Convertir algunas columnas a tipo string
    df1[['Grupo logístico', 'Grupo comercial', 'Carg']] = df1[['Grupo logístico', 'Grupo comercial', 'Carg']].astype(str)

    # Eliminar ceros a la izquierda en 'Grupo logístico'
    df1['Grupo logístico'] = df1['Grupo logístico'].apply(lambda x: x.lstrip('0') if isinstance(x, str) else x)

    # Eliminar '.0' al final de los valores en 'Grupo comercial'
    df1['Grupo logístico'] = df1['Grupo logístico'].str.rstrip('.0')

    # Reiniciar los índices
    df1 = df1.reset_index(drop=True)

    # Aplicar la función replace_values y fill_logistic_group
    df1 = df1.groupby('Doc.ventas', group_keys=False).apply(replace_values)
    df1 = df1.groupby('Doc.ventas', group_keys=False).apply(fill_logistic_group)

    # Aplicar la función buscar_y_copiar
    df1['Grupo logístico'] = df1.apply(buscar_y_copiar, axis=1)
    
    # Agrupar por 'Doc.factura jurídico' y sumar correctamente
    df2 = df1.groupby('Doc.factura jurídico', as_index=False).agg(lambda x: x.iloc[0] if x.name != 'Val.neto factura' else x.sum())
    
    # Renombrar columna 'Val.neto factura' a 'Valor Factura'
    df2 = df2.rename(columns={'Val.neto factura': 'Valor Factura'})
    
    return df2

# Función para copiar valores de 'Grupo comercial' o 'Carg' en 'Grupo logístico' si tiene menos de 10 caracteres
def buscar_y_copiar(row):
    grupo_logistico = str(row['Grupo logístico'])
    grupo_comercial = str(row['Grupo comercial'])
    carg = str(row['Carg'])
    
    if len(grupo_logistico) != 10:
        if len(grupo_comercial) == 10:
            return grupo_comercial
        elif len(carg) == 10:
            return carg
        else:
            return grupo_logistico
    else:
        return grupo_logistico

# Función para reemplazar valores en las columnas especificadas si tienen 10 caracteres
def replace_values(group):
    for column in ['Grupo logístico', 'Grupo comercial', 'Carg']:
        value_10_char = group[column][group[column].str.len() == 10].unique()
        if len(value_10_char) == 1:
            group[column] = value_10_char[0]
    return group

# Función para rellenar 'Grupo logístico' si tiene menos de 10 caracteres usando 'Grupo comercial' o 'Carg'
def fill_logistic_group(group):
    for i, row in group.iterrows():
        if pd.notnull(row['Grupo logístico']) and len(row['Grupo logístico']) < 10:
            if pd.notnull(row['Grupo comercial']) and len(row['Grupo comercial']) == 10:
                group.at[i, 'Grupo logístico'] = row['Grupo comercial']
            elif pd.notnull(row['Carg']) and len(row['Carg']) == 10:
                group.at[i, 'Grupo logístico'] = row['Carg']
    return group

# Función para seleccionar archivos a procesar
def browse_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xls *.xlsx")])
    if file_paths:
        cleaned_data_list = []
        original_file_names = []

        # Iniciar barra de progreso
        progress_bar['maximum'] = len(file_paths)
        progress_bar['value'] = 0
        counter_label.config(text=f"0 de {len(file_paths)}")
       
        # Procesar cada archivo
        for i, file_path in enumerate(file_paths):
            cleaned_data = clean_data(file_path)
            cleaned_data_list.append(cleaned_data)
            file_title = os.path.basename(file_path)
            original_file_names.append(file_title)
           
            # Actualizar barra de progreso
            progress_bar['value'] = i + 1
            counter_label.config(text=f"{i + 1} de {len(file_paths)} procesados")
            window.update_idletasks()  # Refrescar la ventana

        window.cleaned_data_list = cleaned_data_list
        window.original_file_names = original_file_names

        # Habilitar botón de descarga
        download_button.config(state=tk.NORMAL)
        message_label.config(text="Datos Cargados")

# Función para descargar los archivos procesados
def download_files():
    if hasattr(window, "cleaned_data_list") and window.cleaned_data_list:
        save_folder = filedialog.askdirectory()

        # Iniciar barra de progreso
        progress_bar['maximum'] = len(window.cleaned_data_list)
        progress_bar['value'] = 0
        counter_label.config(text=f"0 de {len(window.cleaned_data_list)}")
        for i, cleaned_data in enumerate(window.cleaned_data_list):
            original_name = window.original_file_names[i]
            base_name, ext = os.path.splitext(original_name)
            save_path = f"{save_folder}/{base_name}_depurada.xlsx"
            cleaned_data.to_excel(save_path, index=False, engine='openpyxl')
           
            # Actualizar barra de progreso
            progress_bar['value'] = i + 1
            counter_label.config(text=f"{i + 1} de {len(window.cleaned_data_list)} descargados")
            window.update_idletasks()  # Refrescar la ventana

        message_label.config(text="Datos Descargados.")
    else:
        message_label.config(text="No files loaded for download.")

# Crear un marco para los botones y mensajes
frame = tk.Frame(window)

# Crear botones y mensajes
browse_button = tk.Button(frame, text="Seleccione el (los) Archivo(s)", command=browse_files)
download_button = tk.Button(frame, text="Descargar Archivo(s)", command=download_files, state=tk.DISABLED)
message_label = tk.Label(window, text="")

# Ubicar los botones
browse_button.pack(pady=10)
download_button.pack(pady=10)

# Ubicar el marco en la ventana
frame.pack(expand=True, pady=50)
message_label.pack()

# Crear un marco inferior
bottom_frame = tk.Frame(window)
bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

# Firma del desarrollador, posicionada abajo a la izquierda
signature_label = tk.Label(bottom_frame, text="Lllamas", font=("Arial", 10))
signature_label.grid(row=0, column=0, padx=10, pady=5, sticky='w')

# Texto "Nestlé", posicionado abajo a la derecha
signature_label2 = tk.Label(bottom_frame, text="Nestlé", font=("Arial", 10))
signature_label2.grid(row=0, column=1, padx=10, pady=5, sticky='e')

# Expandir la columna 1 para empujar el texto "Nestlé" hacia la derecha
bottom_frame.grid_columnconfigure(0, weight=1)

# Iniciar el bucle principal de la ventana
window.mainloop()
