import imaplib
import email
from email.header import decode_header
from datetime import datetime
import pywhatkit

# Datos del usuario


# Crear conexión
imap = imaplib.IMAP4_SSL("imap.gmail.com")
# iniciar sesión
imap.login(username, password)
# Crear conexión
imap = imaplib.IMAP4_SSL("imap.gmail.com")
# iniciar sesión
imap.login(username, password)

status, mensajes = imap.select("INBOX")

mensajes = int(mensajes[0])
i=14
#for i in range (1,mensajes+1):
try:
    res, mensaje = imap.fetch(str(i), "(RFC822)")
except:
    pass
for respuesta in mensaje:
    if isinstance(respuesta, tuple):
        # Obtener el contenido
        mensaje = email.message_from_bytes(respuesta[1])
        # decodificar el contenido
        subject = decode_header(mensaje["Subject"])[0][0]
        if isinstance(subject, bytes):
            # convertir a string
            subject = subject.decode()
        # de donde viene el correo
        from_ = mensaje.get("From")
        spos1 = subject.find("SD ")+3
        ticket = subject[spos1:spos1+8]
        if from_ == "Noc TPC <noctpc@yeapdata.com>":
            to_ = mensaje.get("To")
            posto = to_.find('@')
            to_ = to_[1:posto].replace("."," ")
            Cc_ = mensaje.get("CC")
            areas = ["grabaciones","Facilities","Aplicaciones","networking","RPA","SYSTEM"]
            for area in areas:
                if Cc_.find(area) != -1:
                    areaR = area
                else:
                    areaR = "Validar area del correo"
            if mensaje.is_multipart():
                for part in mensaje.walk():
                    content_type = part.get_content_type()
                    try:
                        body = part.get_payload(decode=True).decode('latin-1')
                    except:
                        pass
                    if content_type == "text/plain":
                        body = str(body).replace(", POR FAVOR REVISAR.","")
                        
                        mens1 = body.find("+-+-+")+5
                        mens2 = body.find("Hora de Evento: ")
                        alerta = body[mens1:mens2]
                        texto = str(alerta)
                    elif content_type == "image/png":
                        nombre_fichero = part.get_filename()
                        if nombre_fichero == "image.png":
                            fp = open("image.jpg",'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()

texto = ("\U0001F4E2 *NOTIFICACION GESTION DE EVENTOS NOC* \U0001F4E2 \n"
+"*SD:* "+ ticket
+"\n*EVENTO:* "+texto
+"*ANALISTA REPORTADO:* "+to_ 
+"\n*AREA:* "+areaR
+"\n*OBSERVACIONES:* Se realiza reporte con analista encargado para su respectiva validación."
+"correo: "+str(i)).replace(", POR FAVOR REVISAR","")

print("Correo: " +str(i)+" ticket: " +ticket)
now = datetime.now()
hora = int(now.hour)
minuntos = int(now.minute+1)
print(str(now.hour) + ": "+str(int(now.minute)+1))
try:
    pywhatkit.sendwhatmsg("+573202922822", texto, hora, minuntos, 15, True, 5)
    texto = ""
except Exception as inst:
    print(type(inst))
    pywhatkit.sendwhatmsg("+573202922822", texto, hora, minuntos+1, 15, True, 5)

#pywhatkit.sendwhatmsg_to_group("FkLYvGFtZZwJV8UT7vdHaJ", texto, 12, 00)
#pywhatkit.sendwhats_image("+573202922822","image.jpg")



#https://unicode.org/emoji/charts/full-emoji-list.html  https://chat.whatsapp.com/FkLYvGFtZZwJV8UT7vdHaJ
#sendwhatmsg(phone_no: str, message: str, time_hour: int, time_min: int, wait_time: int = 15, tab_close: bool = False, close_time: int = 3) -> None
