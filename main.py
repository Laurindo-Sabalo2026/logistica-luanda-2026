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
        df = pd.read_excel("dados.xlsx")
        
        pdf = FPDF('L', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "RELATORIO DE LOGISTICA - LAURINDO SABALO", ln=True, align='C')
        pdf.ln(10)
        
        # Cabeçalho formatado
        pdf.set_fill_color(200, 220, 255)
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(80, 10, " Destino", 1, 0, 'L', True)
        pdf.cell(40, 10, " Custo (Kz)", 1, 0, 'C', True)
        pdf.cell(50, 10, " Motorista", 1, 0, 'C', True)
        pdf.cell(45, 10, " Data", 1, 1, 'C', True)
        
        pdf.set_font("Arial", '', 10)
        for i in range(len(df)):
            linha = df.iloc[i]
            # Tenta ler as colunas A, C, E e F do seu Excel original
            destino = str(linha.iloc[0])[:40] if len(linha) > 0 else "---"
            custo = str(linha.iloc[2]) if len(linha) > 2 else "---"
            motorista = str(linha.iloc[4]) if len(linha) > 4 else "---"
            data = str(linha.iloc[5]) if len(linha) > 5 else "---"
            
            pdf.cell(80, 10, f" {destino}", 1)
            pdf.cell(40, 10, f" {custo}", 1, 0, 'C')
            pdf.cell(50, 10, f" {motorista}", 1, 0, 'C')
            pdf.cell(45, 10, f" {data}", 1, 1, 'C')

        nome_pdf = f"Relatorio_Logistica_{random.randint(100,999)}.pdf"
        pdf.output(nome_pdf)

        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        msg['Subject'] = "RELATORIO DE LOGISTICA COMPLETO"
        msg['From'] = meu_email
        msg['To'] = "laurinds10@gmail.com"
        msg.attach(MIMEText("Relatorio atualizado com os dados do sistema.", 'plain'))

        with open(nome_pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurinds10@gmail.com", msg.as_string())
        print("SUCESSO!")

    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
