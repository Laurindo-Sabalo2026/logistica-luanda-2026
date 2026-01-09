import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def verificar_e_enviar_alerta():
    ficheiro = "meus_locais (1).xlsx"
    limite = 50000 
    
    if os.path.exists(ficheiro):
        df = pd.read_excel(ficheiro)
        # Filtra destinos acima de 50.000 Kz
        caros = df[df['Custo Total (Kz)'] > limite]
        
        if not caros.empty:
            corpo = f"⚠️ ALERTA LOGÍSTICA LUANDA\n\nDestinos caros encontrados:\n{caros[['Destino', 'Custo Total (Kz)']].to_string(index=False)}"
            enviar_email(corpo)
    else:
        print("Erro: Ficheiro Excel não encontrado no GitHub.")

def enviar_email(conteudo):
    meu_email = "laurindokutala.sabalo@gmail.com"
    minha_senha = os.environ.get('MINHA_SENHA') 
    
    msg = MIMEMultipart()
    msg['From'] = meu_email
    msg['To'] = "laurics10@gmail.com"
    msg['Subject'] = "Relatório de Alerta - Luanda 2026"
    msg.attach(MIMEText(conteudo, 'plain'))
    
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(meu_email, minha_senha)
        server.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())

if __name__ == "__main__":
    verificar_e_enviar_alerta()
