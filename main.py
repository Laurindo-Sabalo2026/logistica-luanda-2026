import pandas as pd
import os
import smtplib
import matplotlib.pyplot as plt
from fpdf import FPDF
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from datetime import datetime

def criar_pdf_personalizado(df, col_nome, col_custo, nome_pdf):
    pdf = FPDF()
    pdf.add_page()
    
    # CABEÇALHO PROFISSIONAL
    pdf.set_font("Arial", 'B', 22)
    pdf.set_text_color(0, 51, 102) # Azul Escuro
    pdf.cell(200, 15, "LAURINDO LOGISTICA & SERVICOS", ln=True, align='C')
    
    pdf.set_font("Arial", 'I', 10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(200, 5, "Relatorio Automatico de Controlo de Custos - Luanda", ln=True, align='C')
    pdf.ln(10)
    
    # TABELA COM CORES
    pdf.set_font("Arial", 'B', 12)
    pdf.set_fill_color(200, 220, 255) # Azul Claro
    pdf.set_text_color(0, 0, 0)
    pdf.cell(110, 10, "Destino / Localizacao", border=1, fill=True)
    pdf.cell(40, 10, "Custo (Kz)", border=1, ln=True, fill=True, align='C')
    
    pdf.set_font("Arial", '', 11)
    for _, row in df.iterrows():
        pdf.cell(110, 10, str(row[col_nome]), border=1)
        pdf.cell(40, 10, f"{row[col_custo]:,.2f}", border=1, ln=True, align='C')
    
    pdf.ln(20)
    
    # RODAPE E ASSINATURA
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    pdf.set_font("Arial", 'I', 9)
    pdf.cell(200, 10, f"Gerado em: {data_hoje}", ln=True, align='R')
    pdf.ln(10)
    pdf.line(70, pdf.get_y(), 140, pdf.get_y())
    pdf.cell(200, 8, "Responsavel de Logistica", ln=True, align='C')
    pdf.output(nome_pdf)

def enviar_email_v3(img_nome, pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = "📊 RELATORIO EXECUTIVO: Logistica Luanda"
    msg.attach(MIMEText("Ola Laurindo, o relatorio personalizado da sua empresa ja esta pronto em anexo.", 'plain'))
    
    for nome, tipo in [(pdf_nome, "pdf"), (img_nome, "png")]:
        if os.path.exists(nome):
            with open(nome, "rb") as f:
                if tipo == "pdf":
                    anexo = MIMEApplication(f.read(), _subtype="pdf")
                else:
                    anexo = MIMEImage(f.read())
                anexo.add_header('Content-Disposition', 'attachment', filename=nome)
                msg.attach(anexo)
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, destinatario, msg.as_string())

def verificar_e_enviar_tudo():
    excel = "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_custo = [c for c in df.columns if 'Custo' in c][0]
        caros = df[df[col_custo] > 100]
        if not caros.empty:
            # Grafico
            plt.figure(figsize=(10, 6), facecolor='white')
            plt.bar(caros['Endereço'].str[:15], caros[col_custo], color='orange')
            plt.title('Analise de Custos Elevados - Luanda')
            plt.tight_layout()
            plt.savefig('grafico_empresa.png')
            
            # PDF Personalizado
            criar_pdf_personalizado(caros, 'Endereço', col_custo, "relatorio_executivo.pdf")
            
            # Enviar
            enviar_email_v3('grafico_empresa.png', 'relatorio_executivo.pdf')
            print("!!! RELATORIO EMPRESARIAL ENVIADO !!!")
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    verificar_e_enviar_tudo()
