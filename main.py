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
    pdf.ln(15)
    pdf.line(10, 40, 200, 40)
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 12)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(110, 10, " Destino / Localizacao", border=1, fill=True)
    pdf.cell(40, 10, "Custo (Kz)", border=1, ln=True, fill=True, align='C')
    pdf.set_font("Arial", '', 11)
    pdf.set_text_color(0, 0, 0)
    for _, row in df.iterrows():
        pdf.cell(110, 10, f" {str(row[col_nome])}", border=1)
        pdf.cell(40, 10, f"{row[col_custo]:,.2f}", border=1, ln=True, align='C')
    if os.path.exists(caminho_grafico):
        pdf.ln(10)
        pdf.image(caminho_grafico, x=10, w=180)
    pdf.ln(10)
    pdf.line(60, pdf.get_y() + 10, 150, pdf.get_y() + 10)
    pdf.ln(12)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(200, 8, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    pdf.set_font("Arial", 'I', 9)
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    pdf.cell(200, 8, f"Gerado em Luanda - Data: {data_hoje}", ln=True, align='C')
    pdf.output(nome_pdf)

def enviar_email_limpo(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO LOGISTICA: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Prezado Laurindo, segue o relatorio consolidado em anexo unico.", 'plain'))
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, destinatario, msg.as_string())

def executar_sistema_limpo():
    excel = "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_custo = [c for c in df.columns if 'Custo' in c][0]
        caros = df[df[col_custo] > 100]
        if not caros.empty:
            plt.figure(figsize=(10, 6))
            plt.bar(caros['Endereço'].str[:15], caros[col_custo], color='royalblue')
            plt.title('Custos de Transporte - Luanda')
            plt.savefig('temp_grafico.png')
            plt.close()
            nome_relatorio = "Relatorio_Executivo_Laurindo.pdf"
            criar_pdf_final(caros, 'Endereço', col_custo, nome_relatorio, 'temp_grafico.png')
            enviar_email_limpo(nome_relatorio)
            print("Sucesso!")
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar_sistema_limpo()
