import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime

def gerar_pdf_limpo(df):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(0, 15, "LAURINDO LOGISTICA - RELATORIO DE CONTROLE", ln=True, align='C')
    pdf.ln(5)
    
    # Tabela simplificada para garantir o alinhamento
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(70, 10, " Destino", 1, 0, 'L', True)
    pdf.cell(40, 10, " Custo (Kz)", 1, 0, 'C', True)
    pdf.cell(50, 10, " Motorista", 1, 0, 'C', True)
    pdf.cell(40, 10, " Data Entr.", 1, 1, 'C', True)
    
    pdf.set_font("Arial", '', 9)
    for i in range(len(df)):
        linha = df.iloc[i]
        pdf.cell(70, 9, f" {str(linha.iloc[0])[:35]}", 1)
        pdf.cell(40, 9, f" {str(linha.iloc[2])}", 1, 0, 'C')
        # AQUI ESTA O SEGREDO: Forçamos a coluna 4 (Motorista)
        pdf.cell(50, 9, f" {str(linha.iloc[4])[:20]}", 1, 0, 'C')
        pdf.cell(40, 9, f" {str(linha.iloc[5])}", 1, 1, 'C')

    pdf.ln(20)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 10, "________________________________", ln=True, align='R')
    pdf.cell(0, 5, "Laurindo Sabalo    ", ln=True, align='R')
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, f"Extraido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='R')
    pdf.output("Relatorio_Laurindo.pdf")

def enviar():
    try:
        df = pd.read_excel("meus_locais (1).xlsx")
        gerar_pdf_limpo(df)
        
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        msg['Subject'] = f"🚀 REENVIO: Relatorio Laurindo {datetime.now().strftime('%H:%M')}"
        msg['From'] = meu_email
        msg['To'] = "laurics10@gmail.com"
        msg.attach(MIMEText("Relatorio gerado com mapeamento fixo de colunas.", 'plain'))
        
        with open("Relatorio_Laurindo.pdf", "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename="Relatorio_Laurindo.pdf")
            msg.attach(part)
            
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(meu_email, senha)
        server.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())
        server.quit()
        print("EMAIL ENVIADO!")
    except Exception as e:
        print(f"ERRO: {e}")

if __name__ == "__main__":
    enviar()
