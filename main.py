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
        # Logotipo LL
        self.set_fill_color(0, 102, 204)
        self.rect(10, 10, 15, 15, 'F')
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", 'B', 12)
        self.text(13, 20, "LL")
        
        # Título
        self.set_xy(30, 10)
        self.set_font("Arial", 'B', 16)
        self.set_text_color(0, 102, 204)
        self.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
        
        # Slogan e Linha
        self.set_font("Arial", 'I', 9)
        self.set_text_color(100, 100, 100)
        self.set_xy(30, 17)
        self.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf_premium(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    # 'L' para Landscape (Paisagem) para caber mais colunas
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # --- TABELA EXPANDIDA ---
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    # Cabeçalhos personalizados
    colunas = [
        ("Destino", 65), ("Custo (Kz)", 30), ("Status", 35), 
        ("Motorista", 45), ("Data Entr.", 35), ("Obs", 65)
    ]
    
    for nome, largura in colunas:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    # Dados
    pdf.set_font("Arial", '', 9)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        pdf.cell(65, 8, f" {str(row[col_nome])[:35]}", border=1)
        pdf.cell(30, 8, f"{row[col_custo]:,.2f}", border=1, align='C')
        pdf.cell(35, 8, f" {str(row.get('Status', 'Pendente'))}", border=1, align='C')
        pdf.cell(45, 8, f" {str(row.get('Motorista', 'N/A'))}", border=1, align='C')
        pdf.cell(35, 8, f" {str(row.get('Data_Entrega', '---'))}", border=1, align='C')
        pdf.cell(65, 8, f" {str(row.get('Obs', 'Sem notas'))[:40]}", border=1, ln=True)

    # --- GRÁFICO ---
    if os.path.exists(caminho_grafico):
        pdf.ln(10)
        # Centralizado na página paisagem
        pdf.image(caminho_grafico, x=60, w=170)
    
    # --- ASSINATURA E DATA NO FUNDO ---
    pdf.set_y(-30)
    pdf.set_draw_color(0, 0, 0)
    pdf.line(110, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 5, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 5, f"Relatorio Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    
    msg = MIMEMultipart()
    msg['Subject'] = f"🚚 RELATORIO DE CARGA PREMIUM: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Olá Laurindo, segue o novo relatório detalhado com campos de motorista e prazos.", 'plain'))
    
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, destinatario, msg.as_string())
        print("Relatório enviado!")
    except Exception as e:
        print(f"Erro no e-mail: {e}")

def executar():
    excel = "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_custo = [c for c in df.columns if 'Custo' in c][0]
        
        # Gráfico Verde Profissional
        plt.figure(figsize=(12, 5))
        plt.bar(df['Endereço'].str[:12], df[col_custo], color='#2E8B57') 
        plt.title('Custos por Destino - Laurindo Logistica')
        plt.ylabel('Kz')
        plt.savefig('grafico_premium.png')
        plt.close()
        
        nome_pdf = "Relatorio_Logistica_Premium.pdf"
        criar_pdf_premium(df, 'Endereço', col_custo, nome_pdf, 'grafico_premium.png')
        enviar_email(nome_pdf)
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
