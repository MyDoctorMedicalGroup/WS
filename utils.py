# Misceláneas
import re
import base64
import glob
from sys import exit
import pandas as pd
import os
import shutil
import json
from datetime import datetime
import matplotlib.pyplot as plt
from google.colab import files

# Correo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from os.path import basename

# Sharepoint
from google.colab import userdata
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Web Scraping y paralelización
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
import time
import threading

def espera(driver,tiempo,com):
    """Make the driver wait for specified seconds (at most) while executing an action

    Parameters
    ----------
    driver : WebDriver
        The driver what is being used
    tiempo : int
        The time (in seconds) multiplied by 5 is the time will the driver wait at most (checking each 5 seconds)
    com : str
        The action the driver will try

    Returns
    -------
    None
    """
    a = True
    p = 0
    while a==True and p<=tiempo:
        try:
            exec(com)
            a = False
            #print("al fin...")
        except:
            time.sleep(5)
            #print("esperando")
            p += 1
            pass

def dividir_lista(lista, n):
    k, m = divmod(len(lista), n)
    return (lista[i * k + min(i, m):(i + 1) * k + min(i + 1, m)] for i in range(n))

def dividir_diccionario(dic, n):
    keys = list(dic.keys())
    division = len(keys) // n
    if division==0:
      division=1
    print(division)
    return [dict((k, dic[k]) for k in keys[i:i + division]) for i in range(0, len(keys), division)]

def send_email(files_list=[]):
  sender_address = 'mdmgdata18@gmail.com'
  sender_pass = 'zlhx ulzd ojtx sfkv'

  #Iniciamos la conexión al servidor de correo electrónico
  session = smtplib.SMTP('smtp.gmail.com',587)
  #session = smtplib.SMTP('smtp.office365.com',587)
  session.starttls()
  session.login(sender_address, sender_pass)
  files=files_list

  #Obtenemos una dirección de correo electrónico
  for correo in ['nathobi@mydoctortampa.com']:
      #Preparamos el correo electrónico
      message=MIMEMultipart()
      message['From'] = sender_address
      if len(files_list)>0:
        message['Subject'] = 'Fallo en script de Athena reports - No total de archivos'
        mail_content = 'Este correo es para indicarle que el script de Athena reports no ha logrado descargar en su totalidad los informes esperados, una posibilidad es volver a ejecutar el script empleando el archivo "no logro.xlsx"'
        message.attach(MIMEText(mail_content,'plain'))
        for f in files or []:
            with open(f, "rb") as fil:
                ext = f.split('.')[-1:]
                attachedfile = MIMEApplication(fil.read(), _subtype = ext)
                attachedfile.add_header(
                    'content-disposition', 'attachment', filename=basename(f) )
            message.attach(attachedfile)
      else:
        message['Subject'] = 'Fallo en script de Athena reports - Zona no válida'
        mail_content = 'Este correo es para indicarle que el script de Athena no completó su labor debido a no pertenecer a una zona IP válida'
        message.attach(MIMEText(mail_content,'plain'))
      #Enviamos el correo
      text = message.as_string()
      session.sendmail(sender_address, correo, text)
      print("Correo enviado")
  session.quit()
