import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
from datetime import datetime, timedelta
import schedule
import time
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import os

gmail_user = "mail"
gmail_password = "key"


# Replace 'path/to/your/credentials.json' with the actual path to your JSON key file
credentials_file = 'autoemailer-396122-fac1b5274034.json'

# Define the OAuth2 scopes for Google Sheets
scopes = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Authenticate using the JSON key file and specified scopes
credentials = Credentials.from_service_account_file(credentials_file, scopes=scopes)
client = gspread.authorize(credentials)

# Input Curso
curso = "Data Translator"
codigo_curso = "Q3S23DTdummy"

# Open the Google Sheets document by title
document = client.open("Fechas e Informacion de Cursos")


# Extrac sheet data
def sheet_df(sheet_name):
    sheet = document.worksheet(sheet_name)
    data = sheet.get_all_values()
    header = data[0]
    data = data[1:]
    pd.set_option('display.max_colwidth', None)
    df = pd.DataFrame(data, columns=header)
    return df


df_IPE = sheet_df("Información de Programa Estandar")
df_IF = sheet_df("Info Fechas")
df_II = sheet_df("Info Instructores")
df_CT = sheet_df("Lista Correos")
df_EI = sheet_df("Estudiantes Inscritos")

# Set your Gmail credentials
gmail_user = gmail_user
gmail_password = gmail_password

# Set the recipient email address
to_email = "fausto@datarebels.mx"

# Define the starting and end dates
starting_date = datetime(2023, 8, 20) # El lunes de la primer semana que tienen clase
end_date = datetime(2023, 9, 10)


# Calculate the scheduled dates
fecha_correo_bienvenida = starting_date - timedelta(weeks=2)
fecha_correo_induccion = starting_date - timedelta(weeks=1)
fecha_correo_grupo_comm_ambtec = fecha_correo_induccion + timedelta(days=1)
fecha_semana_1 = starting_date
fecha_viernes_1 = starting_date + timedelta(days=4)
fecha_mitad = starting_date + (end_date - starting_date)/2
fecha_final_2 = end_date - timedelta(weeks=1) - timedelta(days=4)
fecha_final_1 = end_date - timedelta(days=4)
fecha_dia_antes = end_date - timedelta(days=1)




def correo_template(nombre):
    url_p = df_CT[df_CT["Nombre_Correo_"] == nombre]["Link_HTML_"]
    url = str(url_p.iloc[0])
    response = requests.get(url)
    content = response.text
    return content

# Correo Bienvenida
Programa_ = curso
# nombre = Varia iterando sobre los nombres y correos al mandarse
Dia_Onboarding_ = df_IF[df_IF["Codigo_Curso"] == codigo_curso]["Dia_Onboarding_"]
Horario_Onboarding_ = df_IF[df_IF["Codigo_Curso"] == codigo_curso]["Horario_Onboarding_"]
Descripcion_Gen_Programa_ = df_IPE[df_IPE["Programa_"] == curso]["Descripcion_Gen_Programa_"]
Instructor_ = df_IF[df_IF["Codigo_Curso"] == codigo_curso]["Instructor_"]
Rol_ = df_II[df_II["Instructor_"] == str(Instructor_.iloc[0])]["Rol_"]
Instructor_2 = df_IF[df_IF["Codigo_Curso"] == codigo_curso]["Instructor_2"]
Rol_2 = df_II[df_II["Instructor_"] == str(Instructor_2.iloc[0])]["Rol_"]
Short_Bio_1 = df_II[df_II["Instructor_"] == str(Instructor_.iloc[0])]["Short_Bio_"]
Short_Bio_2 = df_II[df_II["Instructor_"] == str(Instructor_2.iloc[0])]["Short_Bio_"]
Fecha_fin_ = df_IF[df_IF["Codigo_Curso"]== codigo_curso]["Fecha_Fin_"]
Imagen_Estructura_Semana = df_IF[df_IF["Codigo_Curso"]==codigo_curso]["Imagen_Estructura_Semana"]
Horario_Sesiones = df_IF[df_IF["Codigo_Curso"] == codigo_curso]["Horarios_Sesiones"]
Dia_Sesiones = df_IF[df_IF["Codigo_Curso"] == codigo_curso]["Dias_Sesiones_"]
Fechas_Todas_ = df_IF[df_IF["Codigo_Curso"] == codigo_curso]["Fechas_Todas_"]



# Lista nombres
Lista_nombres = list(df_EI[df_EI["Programa_"]==curso]["Nombre_"])
Lista_correos = list(df_EI[df_EI["Programa_"]==curso]["Correo_"])

# Replace specific words
# replacement_mapping = {
#    "Programa_": Programa_,
#    "Dia_Onboarding_": str(Dia_Onboarding_.iloc[0]),
#    "Horario_Onboarding_": str(Horario_Onboarding_.iloc[0]),
#    "Descripcion_Gen_Programa_": str(Descripcion_Gen_Programa_.iloc[0]),
#    "Instructor_": str(Instructor_.iloc[0]),
#    "Rol_": str(Rol_.iloc[0]),
#    "Instructor_2": str(Instructor_2.iloc[0]),
#    "Rol_2": str(Rol_2.iloc[0]),
#    "Short_Bio_1": str(Short_Bio_1.iloc[0]),
#    "Short_Bio_2": str(Short_Bio_2.iloc[0])
#}

# Template correo
correo_b = correo_template("correo_bienvenida")


#for old_word, new_word in replacement_mapping.items():
#    content_correo_b = correo_b.replace(old_word, str(new_word))


# Create a dictionary where names are keys and emails are values
name_email_dict = {nombre: correo for nombre, correo in zip(Lista_nombres, Lista_correos)}

print(name_email_dict)

correo_b= correo_b.replace("Descripcion_Gen_Programa_", str(Descripcion_Gen_Programa_.iloc[0]))
correo_b = correo_b.replace("Programa_", Programa_)
correo_b = correo_b.replace("Dia_Onboarding_", str(Dia_Onboarding_.iloc[0]))
correo_b = correo_b.replace("Horario_Onboarding_", str(Horario_Onboarding_.iloc[0]))
correo_b = correo_b.replace("Instructor_2", str(Instructor_2.iloc[0]))
correo_b = correo_b.replace("Instructor_", str(Instructor_.iloc[0]))
correo_b = correo_b.replace("Rol_2", str(Rol_2.iloc[0]))
correo_b = correo_b.replace("Rol_", str(Rol_.iloc[0]))
correo_b = correo_b.replace("Short_Bio_1", str(Short_Bio_1.iloc[0]))
correo_b = correo_b.replace("Short_Bio_2", str(Short_Bio_2.iloc[0]))
correo_b = correo_b.replace("Fecha_Fin_", str(Fecha_fin_.iloc[0]))
correo_b = correo_b.replace("Imagen_Estructura_Semana", f'<img src="{str(Imagen_Estructura_Semana.iloc[0])}" alt="Image">')
correo_b =correo_b.replace("Horarios_Sesiones", str(Horario_Sesiones.iloc[0]))
correo_b =correo_b.replace("Dias_Sesiones_", str(Dia_Sesiones.iloc[0]))
correo_b = correo_b.replace("Fechas_Todas_", str(Fechas_Todas_.iloc[0]))

for nombre, correo in name_email_dict.items():
    correo_b_1 = correo_b.replace("Nombre_", nombre)
    # Create the MIME object
    message = MIMEMultipart()
    message["From"] = gmail_user
    message["To"] = correo
    message["Subject"] = "Bienvenido al curso!!"
    # Attach the fetched HTML content to the email
    message.attach(MIMEText(correo_b_1, "html"))

    # Connect to Gmail's SMTP server
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        # Log in to your Gmail account
        server.login(gmail_user, gmail_password)

        # Send the email
        server.sendmail(gmail_user, correo, message.as_string())
        

print("Email Bienvenida Enviado!")

# Comienza procesamiento de Email2
correo_e = correo_template("correo_tarea_aviso")

# Cargamos la sheet
df_PE = sheet_df("Progreso Estudiantes")

# Convert "Cuenta_Missings" column to numerical format
df_PE["Cuenta_Missings"] = pd.to_numeric(df_PE["Cuenta_Missings"], errors="coerce")

# Lista nombres
Lista_nombres_reprobados = list(df_PE[df_PE["Cuenta_Missings"] >=5]["user"])
Lista_correos_reprobados = list(df_PE[df_PE["Cuenta_Missings"] >=5]["email"])

reprobados_dict = {nombre: correo for nombre, correo in zip(Lista_nombres_reprobados, Lista_correos_reprobados)}
print(reprobados_dict)

for nombre, correo in reprobados_dict.items():
    correo_e_1 = correo_e.replace("user", nombre)
    # Create the MIME object
    message = MIMEMultipart()
    message["From"] = gmail_user
    message["To"] = correo
    message["Subject"] = "AVISO IMPORTANTE! (Urgente)"
    # Attach the fetched HTML content to the email
    message.attach(MIMEText(correo_e_1, "html"))

    # Connect to Gmail's SMTP server
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        # Log in to your Gmail account
        server.login(gmail_user, gmail_password)

        # Send the email
        server.sendmail(gmail_user, correo, message.as_string())
        

print("Email Aviso Enviado!")