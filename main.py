import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def verificar_e_enviar_alerta():
    # Nome exato como aparece no seu GitHub
    ficheiro = "meus_locais (1).xlsx"
    limite = 50000 
    
    print(f"Tentando abrir o ficheiro: {ficheiro}")
    
    if os.path.exists(ficheiro):
        try:
            df = pd.read_excel(ficheiro, engine='openpyxl')
            # Garante que o nome da coluna não tem espaços extras
            df.columns = df.columns.str.strip()
            
            caros = df[df['Custo Total (Kz)'] > limite]
            
            if not caros.empty:
                corpo = f"⚠️ ALERTA LOGÍSTICA LUANDA\n\nDestinos caros encontrados:\n{caros[['Destino', 'Custo Total (Kz)']].to_string(index=False)}"
                enviar_email(corpo)
                print("E-mail enviado com sucesso!")
            else:
                print("Nenhum destino acima do limite encontrado.")
        except Exception as e:
            print(f"Erro ao ler o Excel: {e}")
    else:
        print(f"ERRO CRÍTICO: O ficheiro '{ficheiro}' não foi encontrado na pasta principal.")

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
