import os
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

def testar_envio():
    try:
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        # Destinatario que voce forneceu
        destino = "laurinds10@gmail.com" 
        
        msg = MIMEText(f"Teste de Seguranca Laurindo - Horario: {datetime.now().strftime('%H:%M:%S')}")
        msg['Subject'] = "Aviso Importante Logistica"
        msg['From'] = meu_email
        msg['To'] = destino
        
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(meu_email, senha)
        server.sendmail(meu_email, destino, msg.as_string())
        server.quit()
        print("TESTE DE TEXTO ENVIADO COM SUCESSO!")
    except Exception as e:
        print(f"ERRO NO TESTE: {e}")

if __name__ == "__main__":
    testar_envio()
