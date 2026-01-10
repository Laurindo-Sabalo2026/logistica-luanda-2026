import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def verificar_e_enviar_alerta():
    ficheiro = "meus_locais (1).xlsx"
    if not os.path.exists(ficheiro):
        return

    try:
        df = pd.read_excel(ficheiro, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        
        coluna_custo = [c for c in df.columns if 'Custo' in c]
        
        if coluna_custo:
            col_nome = coluna_custo[0]
            caros = df[df[col_nome] > 100]
            
            if not caros.empty:
                # CORREÇÃO AQUI: Mudamos 'Destino' para 'Endereço'
                corpo = f"⚠️ ALERTA LOGÍSTICA LUANDA\n\nDestinos caros encontrados:\n{caros[['Endereço', col_nome]].to_string(index=False)}"
                enviar_email(corpo)
                print("!!! E-MAIL ENVIADO COM SUCESSO !!!")
    except Exception as e:
        print(f"Erro ao ler o Excel: {e}")

def enviar_email(conteudo):
    meu_email = "laurindokutala.sabalo@gmail.com"
    minha_senha = os.environ.get('MINHA_SENHA').replace(" ", "")
    msg = MIMEMultipart()
    msg['From'] = meu_email
    msg['To'] = "laurics10@gmail.com"
    msg['Subject'] = "RELATÓRIO: Alerta de Custos Luanda"
    msg.attach(MIMEText(conteudo, 'plain'))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(meu_email, minha_senha)
        server.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())

if __name__ == "__main__":
    verificar_e_enviar_alerta()
