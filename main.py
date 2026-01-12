import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime

def gerar_pdf_profissional(df):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.add_page()
    
    # --- TOPO PROFISSIONAL ---
    pdf.set_fill_color(0, 51, 102) 
    pdf.rect(0, 0, 297, 30, 'F')
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 15, "LAURINDO LOGISTICA & SERVICOS", ln=True, align='C')
    pdf.set_font("Arial", 'I', 10)
    pdf.cell(0, 5, "Gestao de Custos e Operacoes - Luanda", ln=True, align='C')
    
    pdf.ln(15)
    
    # --- TABELA DE DADOS ---
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.set_fill_color(200, 220, 255)
    
    # Cabecalhos
    pdf.cell(80, 10, " Destino", 1, 0, 'L', True)
    pdf.cell(40, 10, " Custo (Kz)", 1, 0, 'C', True)
    pdf.cell(40, 10, " Status", 1, 0, 'C', True)
    pdf.cell(50, 10, " Motorista", 1, 0, 'C', True)
    pdf.cell(40, 10, " Data Entr.", 1, 1, 'C', True)
    
    pdf.set_font("Arial", '', 10)
    # Preenchimento garantindo a ordem das colunas do seu Excel
    for i in range(len(df)):
        linha = df.iloc[i]
        pdf.cell(80, 10, f" {str(linha.iloc[0])[:35]}", 1) # Destino
        pdf.cell(40, 10, f" {str(linha.iloc[2])}", 1, 0, 'C') # Custo
        pdf.cell(40, 10, f" {str(linha.iloc[3])}", 1, 0, 'C') # Status
        pdf.cell(50, 10, f" {str(linha.iloc[4])}", 1, 0, 'C') # Motorista
        pdf.cell(40, 10, f" {str(linha.iloc[5])}", 1, 1, 'C') # Data

    # --- RODAPE ---
    pdf.ln(20)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 10, "__________________________", ln=True, align='R')
    pdf.cell(0, 5, "Laurindo Sabalo    ", ln=True, align='R')
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='R')
    
    nome_arq = "Relatorio_Logistica_Final.pdf"
    pdf.output(nome_arq)
    return nome_arq

def enviar():
    try:
        df = pd.read_excel("meus_locais (1).xlsx")
        pdf_anexo = gerar_pdf_profissional(df)
        
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        destino = "laurinds10@gmail.com"
        
        msg = MIMEMultipart()
        msg['Subject'] = f"RELATORIO LOGISTICA: {datetime.now().strftime('%d/%m/%Y')}"
        msg['From'] = remetente = meu_email
        msg['To'] = destino
        
        msg.attach(MIMEText("Bom dia Laurindo, segue o relatorio processado.", 'plain'))
        
        with open(pdf_anexo, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=pdf_anexo)
            msg.attach(part)
            
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, destino, msg.as_string())
        print("RELATORIO ENVIADO!")
    except Exception as e:
        print(f"ERRO: {e}")

if __name__ == "__main__":
    enviar()
