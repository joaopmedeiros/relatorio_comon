import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def enviaEmailComAnexo(de, para, assunto, corpoEmail, caminhoAnexo, nomeAnexo):
    msg = MIMEMultipart()

    msg['From'] = de
    msg['To'] = para
    msg['Subject'] = assunto

    corpoEmail = corpoEmail

    msg.attach(MIMEText(corpoEmail, 'plain'))

    filename = nomeAnexo
    attachment = open(caminhoAnexo + nomeAnexo, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)
    try:
       smtpObj = smtplib.SMTP('mail.zanc.com.br', 25)
       smtpObj.sendmail(de, para, msg.as_string())
       print ("Email enviado com Sucesso")
    except:
       print ("Erro no envio do Email")