from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
from datetime import datetime


app = Flask(__name__)

# Verificar si el archivo Excel existe
excel_path = r"C:\Users\grios\Desktop\myapp\registradorFinanciero\registros_financieros.xlsx"

if os.path.exists(excel_path):
    # Cargar datos existentes
    entregas_df = pd.read_excel(excel_path, sheet_name='Entregas')
    pagos_df = pd.read_excel(excel_path, sheet_name='Pagos')
    gastos_carreras_df = pd.read_excel(excel_path, sheet_name='Gastos_Carreras')
    gastos_extras_df = pd.read_excel(excel_path, sheet_name='Gastos_Extras')
else:
    # Crear DataFrames si el archivo no existe
    entregas_df = pd.DataFrame(columns=['Fecha', 'Hora', 'Monto', 'Moneda', 'Descripcion', 'Registro'])
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

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Obtener los datos del formulario
        n_fecha = request.form['fecha']
        n_hora = request.form['hora']
        n_monto = request.form['monto']
        n_moneda = request.form['moneda'].lower()
        n_descripcion = request.form['descripcion']
        n_registro = datetime.now()

        # Validación básica
        try:
            float(n_monto)
        except ValueError:
            return "El monto debe ser un número."

        # Crear una nueva fila en el DataFrame
        nueva_fila = pd.DataFrame({
            'Fecha': [n_fecha],
            'Hora': [n_hora],
            'Monto': [n_monto],
            'Moneda': [n_moneda],
            'Descripcion': [n_descripcion],
            'Registro': [n_registro]
        })
        print(nueva_fila)
        entregas_df = pd.DataFrame(nueva_fila)

        # Guardar los datos en el archivo Excel
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            fila_vacia = encontrar_primera_fila_vacia(excel_path)
            print(fila_vacia)
            entregas_df.to_excel(writer, sheet_name='Entregas', index=False, header=not os.path.exists(excel_path), startrow=fila_vacia)

        return redirect(url_for('index'))
    
    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
