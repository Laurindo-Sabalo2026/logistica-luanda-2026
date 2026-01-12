import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import random

def executar():
    try:
        # Tenta nomes diferentes caso você não tenha renomeado ainda
        arquivos_possiveis = ["dados.xlsx", "meus_locais (1).xlsx", "meus_locais.xlsx"]
        arquivo_alvo = None
        
        for nome in arquivos_possiveis:
            if os.path.exists(nome):
                arquivo_alvo = nome
                break
        
        if not arquivo_alvo:
            print("ERRO: Arquivo Excel não encontrado no GitHub!")
            return

        df = pd.read_excel(arquivo_alvo)
        
        # Criar PDF Simples para teste rápido
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, "RELATORIO LOGISTICA - TESTE FINAL", ln=True, align='C')
        
        nome_pdf = f"Relatorio_{random.randint(100,999)}.pdf"
        pdf.output(nome_pdf)

        # Configuração de Email
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        destino = "laurinds10@gmail.com"

        msg = MIMEMultipart()
        msg['Subject'] = f"Relatorio de Logistica - Ref {random.randint(10,99)}"
        msg['From'] = meu_email
        msg['To'] = destino
        msg.attach(MIMEText("O robô executou com sucesso. Veja o anexo.", 'plain'))

        with open(nome_pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, destino, msg.as_string())
        
        print("ENVIADO! Verifique agora a sua caixa de entrada.")

    except Exception as e:
        print(f"FALHA: {e}")

if __name__ == "__main__":
    executar()
