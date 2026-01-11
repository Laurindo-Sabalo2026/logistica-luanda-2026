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
    
    # --- LOGOTIPO AZUL ---
    pdf.set_fill_color(0, 102, 204) 
    pdf.rect(15, 15, 15, 15, 'F')  
    pdf.set_text_color(255, 255, 255) 
    pdf.set_font("Arial", 'B', 12)
    pdf.text(18, 25, "LL") 

    # --- TEXTO DO CABEÇALHO ---
    pdf.set_xy(35, 15)
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
    
    # --- RECUPERANDO O SLOGAN (O QUE TINHA SUMIDO) ---
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.set_xy(35, 22)
    pdf.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
    
    # --- LINHA QUE ATRAVESSA A PÁGINA ---
    pdf.ln(8)
    pdf.set_draw_color(200, 200, 200) 
    pdf.line(15, 35, 195, 35) 
    pdf.ln(10)

    # --- TABELA COM TRÊS COLUNAS ---
    pdf.set_font("Arial", 'B', 11)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    pdf.cell(80, 10, " Destino", border=1, fill=True)
    pdf.cell(35, 10, "Custo (Kz)", border=1, fill=True, align='C')
    pdf.cell(35, 10, "Status", border=1, ln=True, fill=True, align='C')
    
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        pdf.cell(80, 10, f" {str(row[col_nome])[:35]}", border=1)
        pdf.cell(35, 10, f"{row[col_custo]:,.2f}", border=1, align='C')
        status_texto = str(row['Status']) if 'Status' in df.columns else "N/A"
        pdf.cell(35
