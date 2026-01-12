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
        # 1. LER O EXCEL (dados.xlsx)
        df = pd.read_excel("dados.xlsx")
        
        # 2. CONFIGURAR O PDF (Horizontal)
        pdf = FPDF('L', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "RELATORIO DE LOGISTICA - LAURINDO SABALO", ln=True, align='C')
        pdf.ln(10)
        
        # Cabeçalho Azul com Letras Brancas
        pdf.set_fill_color(0, 51, 102) 
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(85, 10, " DESTINO", 1, 0, 'L', True)
        pdf.cell(45, 10, " CUSTO (Kz)", 1, 0, 'C', True)
        pdf.cell(60, 10, " MOTORISTA", 1, 0, 'C', True)
        pdf.cell(45, 10, " DATA", 1, 1, 'C', True)
        
        pdf.set_text_color(0, 0, 0) # Volta para preto
        pdf.set_font("Arial", '', 10)
        
        # 3. PREENCHER OS DADOS (Lógica Inteligente)
        for i in range(len(df)):
            linha = df.iloc[i].dropna().tolist() # Remove espaços vazios
            
            # Garante que temos pelo menos as 4 informações necessárias
            destino = str(linha[0])[:40] if len(linha) > 0 else "---"
            custo = str(linha[2]) if len(linha) > 2 else "---"
            motorista = str(linha[4]) if len(linha) > 4 else "---"
            data = str(linha[5]) if len(linha) > 5 else "---"
            
            pdf.cell(85, 10, f" {destino}", 1)
            pdf.cell(45, 10, f" {custo}", 1, 0, 'C')
            pdf.cell(60, 10, f" {motorista}", 1, 0, 'C')
            pdf.cell(45, 10, f" {data}", 1, 1, 'C')

        nome_pdf = "Relatorio_Final_Concluido.pdf"
        pdf.output(nome_pdf)

        # 4. ENVIAR POR EMAIL
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        msg['Subject'] = "RELATORIO DE LOGISTICA - ENTREGA FINAL"
        msg['From'] = meu_email
        msg['To'] = "laurinds10@gmail.com"
        msg.attach(MIMEText("Ola Laurindo, o seu robo terminou o processamento dos dados.", 'plain'))

        with open(nome_pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurinds10@gmail.com", msg.as_string())
        
        print("MISSÃO CUMPRIDA! E-mail enviado.")

    except Exception as e:
        print(f"Erro no processamento: {e}")

if __name__ == "__main__":
    executar()
