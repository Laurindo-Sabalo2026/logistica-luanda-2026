import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime
import random

def executar():
    try:
        # Carregar Excel
        df = pd.read_excel("meus_locais (1).xlsx")
        
        # Gerar PDF
        pdf = FPDF('L', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "RELATORIO DE LOGISTICA - LAURINDO SABALO", ln=True, align='C')
        pdf.ln(10)
        
        # Tabela (Colunas: Destino, Custo, Motorista, Data)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(80, 10, " Destino", 1)
        pdf.cell(40, 10, " Custo", 1)
        pdf.cell(50, 10, " Motorista", 1)
        pdf.cell(40, 10, " Data", 1, 1)
        
        pdf.set_font("Arial", '', 10)
        for i in range(len(df)):
            linha = df.iloc[i]
            pdf.cell(80, 10, str(linha.iloc[0])[:35], 1) # Coluna A
            pdf.cell(40, 10, str(linha.iloc[2]), 1)       # Coluna C
            pdf.cell(50, 10, str(linha.iloc[4]), 1)       # Coluna E
            pdf.cell(40, 10, str(linha.iloc[5]), 1, 1)    # Coluna F

        # Nome aleatorio para evitar bloqueio
        nome_pdf = f"Relatorio_Final_{random.randint(1000,9999)}.pdf"
        pdf.output(nome_pdf)

        # Envio
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        msg['Subject'] = f"Relatorio Logistica - Ref {random.randint(100,999)}"
        msg['From'] = meu_email
        msg['To'] = "laurinds10@gmail.com"
        msg.attach(MIMEText("Segue o relatorio atualizado.", 'plain'))

        with open(nome_pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurinds10@gmail.com", msg.as_string())
        print("ENVIADO COM SUCESSO!")

    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
