import pandas as pd
from datetime import datetime
import os.path
from pandas import ExcelWriter
import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
from tkinter import ttk


# Verificar si el archivo Excel existe
excel_path = r"C:\Users\aguod\Desktop\registradorFinanciero\registros_financieros.xlsx"
assert os.path.isfile(excel_path)
with open(excel_path, "r") as f:
    pass

if os.path.exists(excel_path):
    # Cargar datos existentes
    entregas_df = pd.read_excel(excel_path, sheet_name='Entregas')
    pagos_df = pd.read_excel(excel_path, sheet_name='Pagos')
    gastos_carreras_df = pd.read_excel(excel_path, sheet_name='Gastos_Carreras')
    gastos_extras_df = pd.read_excel(excel_path, sheet_name='Gastos_Extras')
else:
    # Crear DataFrames si el archivo no existe
    entregas_df = pd.DataFrame(columns=['Fecha', 'Hora', 'Monto', 'Moneda'])
    pagos_df = pd.DataFrame(columns=['Fecha', 'Hora', 'Monto', 'Moneda'])
    gastos_carreras_df = pd.DataFrame(columns=['Fecha', 'Hora', 'Monto', 'Moneda'])
    gastos_extras_df = pd.DataFrame(columns=['Fecha', 'Hora', 'Monto', 'Moneda'])


def encontrar_primera_fila_vacia(excel_path):
    print("Funcion encontrar primera fila ejecutando..")
    try:
        # Leer el archivo Excel
        df = pd.read_excel(excel_path, header=None)
        
        # Encontrar la primera fila vacía
        fila_vacia = df.index[df.isnull().all(axis=1)].min()
        
        if pd.isna(fila_vacia):
            # Si no se encontró ninguna fila vacía, asignar el índice siguiente
            fila_vacia = len(df)
        
    except FileNotFoundError:
        # Si el archivo no existe, empezar desde la primera fila
        fila_vacia = 0

    return fila_vacia


def limpiar_campos():
    print("Funcion limpiar campos ejecutando..")
    entry_fecha.delete(0, tk.END)
    entry_hora.delete(0, tk.END)
    entry_monto.delete(0, tk.END)
    entry_moneda.delete(0, tk.END)


def salir():
    root.quit()
    root.destroy()

def submit():
    print("Funcion submit ejecutando..")
    n_fecha = entry_fecha.get()
    n_hora = entry_hora.get()
    n_monto = entry_monto.get()
    n_moneda = entry_moneda.get().lower()

    # Validación básica
    try:
        float(n_monto)
    except ValueError:
        messagebox.showerror("Error", "El monto debe ser un número.")
        return
    
    fecha=[]
    hora=[]
    monto=[]
    moneda=[]

    fecha.append(n_fecha)
    hora.append(n_hora)
    monto.append(n_monto)
    moneda.append(n_moneda)

    nueva_fila = {'Fecha': fecha, 
                'Hora': hora, 
                'Monto': monto, 
                'Moneda': moneda}
    
    entregas_df = pd.DataFrame(nueva_fila)

    # Mostrar un mensaje de confirmación
    messagebox.showinfo("Datos guardados", "Los datos han sido guardados exitosamente.")
    limpiar_campos()
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay' ) as writer:
    # Encontrar la primera fila vacía
        fila_vacia = encontrar_primera_fila_vacia(excel_path)
        entregas_df.to_excel(writer, sheet_name='Entregas', index=False, header=not os.path.exists(excel_path), startrow=fila_vacia)
    
        pagos_df.to_excel(writer, sheet_name='Pagos', index=False, header=not os.path.exists(excel_path))
        gastos_carreras_df.to_excel(writer, sheet_name='Gastos_Carreras', index=False, header=not os.path.exists(excel_path))
        gastos_extras_df.to_excel(writer, sheet_name='Gastos_Extras', index=False, header=not os.path.exists(excel_path))
   

# Crear la ventana principal
root = tk.Tk()
root.title("Ingreso de Datos")

# Crear y colocar los widgets (etiquetas y campos de entrada)
tk.Label(root, text="Fecha:").grid(row=0, column=0, padx=10, pady=5)
entry_fecha = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
entry_fecha.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Hora:").grid(row=1, column=0, padx=10, pady=5)

# Generar una lista de horas en intervalos de 30 minutos
horas = [datetime.strftime(datetime(1900, 1, 1, h, m), '%H:%M') 
         for h in range(24) 
         for m in (0, 30)]

entry_hora = ttk.Combobox(root, values=horas, state="readonly")
entry_hora.set('00:00')  # Valor predeterminado
entry_hora.grid(row=1, column=1, padx=10, pady=5)



# tk.Label(root, text="Hora (HH:MM):").grid(row=1, column=0, padx=10, pady=5)
# entry_hora = tk.Entry(root)
# entry_hora.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Monto:").grid(row=2, column=0, padx=10, pady=5)
entry_monto = tk.Entry(root)
entry_monto.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Moneda:").grid(row=3, column=0, padx=10, pady=5)

# Combobox para seleccionar la moneda
entry_moneda = ttk.Combobox(root, values=["Dolar", "Peso arg"], state="readonly")
entry_moneda.grid(row=3, column=1, padx=10, pady=5)

# tk.Label(root, text="Moneda (Dólares o Pesos):").grid(row=3, column=0, padx=10, pady=5)
# entry_moneda = tk.Entry(root)
# entry_moneda.grid(row=3, column=1, padx=10, pady=5)

# Botón para enviar los datos
submit_button = tk.Button(root, text="Enviar", command=submit)
submit_button.grid(row=4, column=0, columnspan=2, pady=10)

# Botón para salir
submit_button = tk.Button(root, text="Salir", command=salir)
submit_button.grid(row=4, column=1, columnspan=2, pady=10)

# Ejecutar la aplicación
root.mainloop()