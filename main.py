import pandas as pd
import os
import smtplib
import matplotlib.pyplot as plt
from fpdf import FPDF
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

class PDF_Logistica(FPDF):
    def header(self):
        self.set_fill_color(0, 102, 204)
        self.rect(10, 8, 12, 12, 'F')
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", 'B', 10)
        self.text(12, 16, "LL")
        self.set_xy(25, 8)
        self.set_font("Arial", 'B', 14)
        self.set_text_color(0, 102, 204)
        self.cell(100, 8, "LAURINDO LOGISTICA & SERVICOS", ln=True)
        self.set_font("Arial", 'I', 8)
        self.set_text_color(100, 100, 100)
        self.set_xy(25, 14)
        self.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
        self.line(10, 22, 287, 22)
        self.ln(3)

def criar_pdf_compacto(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # --- TABELA SUPER COMPACTA ---
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    larguras = [("Destino", 65), ("Custo (Kz)", 30), ("Status", 30), ("Motorista", 40), ("Data Entr.", 30), ("Obs", 82)]
    for nome, largura in larguras:
        pdf.cell(largura, 7, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    pdf.set_text_color(0, 0, 0)
    for _, row in df.head(8).iterrows(): # Limitado a 8 linhas para garantir espaço
        pdf.cell(65, 6, f" {str(row[col_nome])[:35]}", border=1)
        pdf.cell(30, 6, f"{row[col_custo]:,.2f}", border=1, align='C')
        pdf.cell(30, 6, f" {str(row.get('Status', 'Pendente'))}", border=1, align='C')
        pdf.cell(40, 6, f" {str(row.get('Motorista', 'N/A'))}", border=1, align='C')
        pdf.cell(30, 6, f" {str(row.get('Data_Entrega', '---'))}", border=1, align='C')
        pdf.cell(82, 6, f" {str(row.get('Obs', 'Sem notas'))[:55]}", border=1, ln=True)

    # --- GRÁFICO PEQUENO (O SEGREDO PARA A PÁGINA ÚNICA) ---
    if os.path.exists(caminho_grafico):
        pdf.ln(2)
        # Reduzido para w=110 para sobrar espaço para a assinatura
        pdf.image(caminho_grafico, x=90, w=110)
    
    # --- ASSINATURA SUBIDA ---
    pdf.set_y(-35) # Subi de -25 para -35 para garantir que não pula de página
    pdf.set_draw_color(0, 0, 0)
    pdf.line(110, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(1)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(0, 5, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    pdf.set_font("Arial", 'I', 7)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 4, f"Relatorio Final - Luanda: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = f"✅ RELATORIO PAGINA UNICA: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Relatorio corrigido para uma unica pagina.", 'plain'))
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, destinatario, msg.as_string())

def executar():
    excel = "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_custo = [c for c in df.columns if 'Custo' in c][0]
        
        # Gráfico com altura muito reduzida (figsize 10x2.5)
        plt.figure(figsize=(10, 2.5))
        plt.bar(df['Endereço'].str[:10], df[col_custo], color='#2E8B57') 
        plt.title('Resumo de Custos')
        plt.tight_layout()
        plt.savefig('grafico_fixo.png')
        plt.close()
        
        nome_pdf = "Relatorio_Laurindo_Final.pdf"
        criar_pdf_compacto(df, 'Endereço', col_custo, nome_pdf, 'grafico_fixo.png')
        enviar_email(nome_pdf)
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
