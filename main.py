import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def verificar_e_enviar_alerta():
    ficheiro = "meus_locais (1).xlsx"
    
    if os.path.exists(ficheiro):
        try:
            df = pd.read_excel(ficheiro, engine='openpyxl')
            # Esta linha limpa espaços e caracteres invisíveis nos títulos
            df.columns = [str(c).strip() for c in df.columns]
            
            # Procuramos qualquer coluna que contenha a palavra "Custo"
            coluna_custo = [c for c in df.columns if 'Custo' in c]
            
            if coluna_custo:
                col_nome = coluna_custo[0]
                # Filtra valores acima de 50.000 Kz
                caros = df[df[col_nome] > 100]
                
                if not caros.empty:
                    corpo = f"⚠️ ALERTA LOGÍSTICA LUANDA\n\nDestinos caros encontrados:\n{caros[['Destino', col_nome]].to_string(index=False)}"
                    enviar_email(corpo)
                    print(f"E-mail enviado! Encontrados {len(caros)} destinos.")
                else:
                    print("Nenhum custo acima de 50.000 Kz encontrado.")
            else:
                print(f"ERRO: Não encontrei nenhuma coluna com o nome 'Custo'. Colunas lidas: {list(df.columns)}")
                
        except Exception as e:
            print(f"Erro no processamento: {e}")
    else:
        print("Ficheiro não encontrado no repositório.")

def enviar_email(conteudo):
    meu_email = "laurindokutala.sabalo@gmail.com"
    # Certifique-se de que MINHA_SENHA no GitHub não tem espaços
    minha_senha = os.environ.get('MINHA_SENHA') 
    
    msg = MIMEMultipart()
    msg['From'] = meu_email
    msg['To'] = "laurics10@gmail.com"
    msg['Subject'] = "ALERTA: Custo Logística Luanda"
    msg.attach(MIMEText(conteudo, 'plain'))
    
    try:
        # Usamos a porta 465 que é mais estável para o Gmail
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(meu_email, minha_senha)
        server.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())
        server.close()
        print("Conexão com Gmail realizada com sucesso!")
    except Exception as e:
        print(f"Erro no servidor de e-mail: {e}")
