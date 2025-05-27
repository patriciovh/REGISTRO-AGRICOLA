from flask import Flask, render_template, request, redirect
import pandas as pd
import os
import smtplib
from email.message import EmailMessage

app = Flask(__name__)

EXCEL_FILE = 'registro.xlsx'

# Crear archivo Excel si no existe
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=["Agricultor", "Labor", "Fecha", "Cultivo"])
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/', methods=['GET', 'POST'])
def formulario():
    if request.method == 'POST':
        agricultor = request.form['agricultor']
        labor = request.form['labor']
        fecha = request.form['fecha']
        cultivo = request.form['cultivo']

        # Leer archivo Excel
        df = pd.read_excel(EXCEL_FILE)

        # Buscar si ya existe fila con agricultor y cultivo iguales
        filtro = (df['Agricultor'] == agricultor) & (df['Cultivo'] == cultivo)
        if filtro.any():
            # Si existe, actualizar esa fila concatenando la fecha y labor
            idx = df[filtro].index[0]

            # Concatenar fechas separados por coma, sin repetir
            fechas_existentes = df.at[idx, 'Fecha']
            fechas_lista = set([f.strip() for f in str(fechas_existentes).split(',')])
            fechas_lista.add(fecha)
            df.at[idx, 'Fecha'] = ', '.join(sorted(fechas_lista))

            # Concatenar labores separados por coma, sin repetir
            labores_existentes = df.at[idx, 'Labor']
            labores_lista = set([l.strip() for l in str(labores_existentes).split(',')])
            labores_lista.add(labor)
            df.at[idx, 'Labor'] = ', '.join(sorted(labores_lista))

        else:
            # Si no existe, agregar fila nueva
            nuevo_registro = pd.DataFrame([[agricultor, labor, fecha, cultivo]], columns=df.columns)
            df = pd.concat([df, nuevo_registro], ignore_index=True)

        # Guardar Excel actualizado
        df.to_excel(EXCEL_FILE, index=False)

        # Enviar correo con Excel actualizado
        enviar_correo(EXCEL_FILE)

        return redirect('/')

    return render_template('formulario.html')

def enviar_correo(archivo):
    remitente = 'ing.patriciovaldes@gmail.com' 
    destinatario = 'patricio.valdes17@inacapmail.cl'  
    contraseña = 'stee xpyp clng bbnv'  

    mensaje = EmailMessage()
    mensaje['Subject'] = 'Nuevo registro de agricultor'
    mensaje['From'] = remitente
    mensaje['To'] = destinatario
    mensaje.set_content('Se ha añadido un nuevo registro. Revisa el archivo adjunto.')

    with open(archivo, 'rb') as f:
        mensaje.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=archivo)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(remitente, contraseña)
        smtp.send_message(mensaje)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
