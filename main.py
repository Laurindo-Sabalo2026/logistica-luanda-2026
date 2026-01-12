import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime

def gerar_pdf_simples(df):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 15, "LAURINDO LOGISTICA - RELATORIO DE EMERGENCIA", ln=True, align='C')
    pdf.ln(10)
    
    # Tabela com mapeamento fixo para evitar que nomes caiam no motorista
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(80, 10, " Destino", 1)
    pdf.cell(40, 10, " Custo (Kz)", 1, 0, 'C')
    pdf.cell(50, 10, " Motorista", 1, 0, 'C')
    pdf.cell(40, 10, " Data", 1, 1, 'C')
    
    pdf.set_font("Arial", '', 10)
    for i in range(len(df)):
        linha = df.iloc[i]
        # Pega as colunas pela posição exata (0, 2, 4, 5)
        pdf.cell(80, 10, f" {str(linha.iloc[0])[:35]}", 1)
        pdf.cell(40, 10, f" {str(linha.iloc[2])}", 1, 0, 'C')
        pdf.cell(50, 10, f" {str(linha.iloc[4])}", 1, 0, 'C')
        pdf.cell(40, 10, f" {str(linha.iloc[5])}", 1, 1, 'C')

    pdf.ln(20)
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, f"Relatorio gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='R')
    pdf.output("Relatorio_Urgente.pdf")

def enviar():
    try:
        df = pd.read_excel("meus_locais (1).xlsx")
        gerar_pdf_simples(df)
        
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        # Assunto totalmente diferente para furar o bloqueio do Gmail
        msg['Subject'] = f"ARQUIVO LOGISTICA PRIORITARIO {datetime.now().strftime('%M%S')}"
        msg['From'] = meu_email
        msg['To'] = "laurics10@gmail.com"
        msg.attach(MIMEText("Envio de seguranca para desbloqueio do sistema.", 'plain'))
        
        with open("Relatorio_Urgente.pdf", "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename="Relatorio_Laurindo.pdf")
            msg.attach(anexo)
            
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())
        print("FOI ENVIADO!")
    except Exception as e:
        print(f"ERRO: {e}")

if __name__ == "__main__":
    enviar()
