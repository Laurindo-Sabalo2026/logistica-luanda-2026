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
        # 1. Carregar os dados do seu Excel
        arquivo = "meus_locais (1).xlsx"
        df = pd.read_excel(arquivo)
        
        # 2. Criar o PDF com os dados reais
        pdf = FPDF('L', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "RELATORIO DE LOGISTICA COMPLETO", ln=True, align='C')
        pdf.ln(10)
        
        # Cabeçalho da Tabela
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(80, 10, " Destino", 1)
        pdf.cell(40, 10, " Custo (Kz)", 1)
        pdf.cell(50, 10, " Motorista", 1)
        pdf.cell(40, 10, " Data", 1, 1)
        
        # Linhas da Tabela (Dados do Excel)
        pdf.set_font("Arial", '', 10)
        for i in range(len(df)):
            linha = df.iloc[i]
            pdf.cell(80, 10, f" {str(linha.iloc[0])[:35]}", 1) # Coluna A
            pdf.cell(40, 10, f" {str(linha.iloc[2])}", 1)       # Coluna C
            pdf.cell(50, 10, f" {str(linha.iloc[4])}", 1)       # Coluna E
            pdf.cell(40, 10, f" {str(linha.iloc[5])}", 1, 1)    # Coluna F

        nome_pdf = f"Relatorio_Final_{random.randint(1000,9999)}.pdf"
        pdf.output(nome_pdf)

        # 3. Enviar por Email
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        msg['Subject'] = f"RELATORIO LOGISTICA - DADOS REAIS - Ref {random.randint(100,999)}"
        msg['From'] = meu_email
        msg['To'] = "laurinds10@gmail.com"
        msg.attach(MIMEText("Ola Laurindo, aqui esta o relatorio com todos os dados do Excel.", 'plain'))

        with open(nome_pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurinds10@gmail.com", msg.as_string())
        
        print("RELATORIO COMPLETO ENVIADO!")

    except Exception as e:
        print(f"Erro ao processar: {e}")

if __name__ == "__main__":
    executar()
