import pandas as pd
import os
import smtplib
import matplotlib.pyplot as plt
from fpdf import FPDF
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

def criar_pdf_final(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    pdf = FPDF()
    pdf.add_page()
    
    # --- CABEÇALHO ---
    pdf.set_fill_color(0, 102, 204) 
    pdf.rect(15, 15, 15, 15, 'F')  
    pdf.set_text_color(255, 255, 255) 
    pdf.set_font("Arial", 'B', 12)
    pdf.text(18, 25, "LL") 

    pdf.set_xy(35, 15)
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
    
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.set_xy(35, 22)
    pdf.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
    
    pdf.ln(8)
    pdf.set_draw_color(200, 200, 200) 
    pdf.line(15, 35, 195, 35) 
    pdf.ln(5)

    # --- TABELA COM 3 COLUNAS (AUMENTADO) ---
    pdf.set_font("Arial", 'B', 11)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    # Definindo larguras: Destino(75), Custo(35), Status(40) = 150mm total
    pdf.cell(75, 10, " Destino", border=1, fill=True)
    pdf.cell(35, 10, "Custo (Kz)", border=1, fill=True, align='C')
    pdf.cell(40, 10, "Status", border=1, ln=True, fill=True, align='C')
    
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        # Coluna 1: Destino
        pdf.cell(75, 10, f" {str(row[col_nome])[:30]}", border=1)
        # Coluna 2: Custo
        pdf.cell(35, 10, f"{row[col_custo]:,.2f}", border=1, align='C')
        # Coluna 3: Status (Pega do Excel, se não existir coloca "Pendente")
        status = str(row['Status']) if 'Status' in df.columns else "Pendente"
        pdf.cell(40, 10, f" {status}", border=1, ln=True, align='C')
    
    # --- GRÁFICO ---
    if os.path.exists(caminho_grafico):
        pdf.ln(10)
        pdf.image(caminho_grafico, x=35, w=130)
    
    # --- ASSINATURA E DATA NO FUNDO ---
    pdf.set_y(-35) 
    pdf.set_draw_color(0, 0, 0)
    pdf.line(70, pdf.get_y(), 140, pdf.get_y())
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 5, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Gerado em Luanda - Data: {datetime.now().strftime('%d/%m/%Y')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO LOGISTICA COMPLETO: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Segue o relatorio oficial com 3 colunas (Destino, Custo e Status).", 'plain'))
    if os.path.exists(pdf_nome
