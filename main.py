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
        self.rect(10, 10, 15, 15, 'F')
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", 'B', 12)
        self.text(13, 20, "LL")
        self.set_xy(30, 10)
        self.set_font("Arial", 'B', 16)
        self.set_text_color(0, 102, 204)
        self.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
        self.set_font("Arial", 'I', 9)
        self.set_text_color(100, 100, 100)
        self.set_xy(30, 17)
        self.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf_analitico(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # --- CÁLCULO DE PERCENTAGEM ---
    total_geral = df[col_custo].sum()
    
    # --- TABELA ---
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    # Ajuste de larguras para caber a nova coluna %
    larguras = [("Destino", 55), ("Custo (Kz)", 30), ("(%)", 20), ("Status", 25), ("Motorista", 35), ("Data", 25), ("Obs", 82)]
    for nome, largura in larguras:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.head(10).iterrows():
        custo_v = float(row[col_custo])
        perc = (custo_v / total_geral) * 100 if total_geral > 0 else 0
        
        pdf.cell(55, 8, f" {str(row[col_nome])[:28]}", border=1)
        pdf.cell(30, 8, f"{custo_v:,.2f}", border=1, align='C')
        pdf.cell(20, 8, f"{perc:.1f}%", border=1, align='C') # Nova coluna
        pdf.cell(25, 8, f" {str(row.get('Status', 'Ok'))}", border=1, align='C')
        pdf.cell(35, 8, f" {str(row.get('Motorista', 'N/A'))}", border=1, align='C')
        pdf.cell(25, 8, f" {str(row.get('Data_Entrega', '--'))}", border=1, align='C')
        pdf.cell(82, 8, f" {str(row.get('Obs', ''))[:45]}", border=1, ln=True)

    # --- LINHA DE TOTAL ---
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(55, 10, " TOTAL GERAL:", border=1, fill=True, align='R')
    pdf.set_text_color(200, 0, 0)
    pdf.cell(30, 10, f"{total_geral:,.2f}", border=1, fill=True, align='C')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(20, 10, "100%", border=1, fill=True, align='C')
    pdf.cell(167, 10, " Kwanzas (Analise de Custos)", border=1, ln=True, fill=True)

    # --- ÁREA INFERIOR ---
    pdf.ln(5)
    y_pos = pdf.get_y()
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=15, y=y_pos, w=140)
    
    pdf.set_xy(180, y_pos + 15)
    pdf.set_draw_color(0, 0, 0)
    pdf.line(180, y_pos + 20, 270, y_pos + 20)
    pdf.set_xy(180, y_pos + 22)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 5, "Laurindo Sabalo", ln=True, align='C')
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(90, 5, f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = f"📈 ANALISE LOGISTICA (%): {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Relatorio com calculo de impacto percentual por destino.", 'plain'))
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
        
        plt.figure(figsize=(8, 3.5))
        plt.bar(df['Endereço'].str[:12], df[col_custo], color='#2E8B57') 
        plt.title('Custos de Transporte por Destino')
        plt.tight_layout()
        plt.savefig('grafico_analise.png')
        plt.close()
        
        nome_pdf = "Relatorio_Analitico_Laurindo.pdf"
        criar_pdf_analitico(df, 'Endereço', col_custo, nome_pdf, 'grafico_analise.png')
        enviar_email(nome_pdf)
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
