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

def criar_pdf_executivo(df, col_nome, col_custo, nome_pdf):
    pdf = FPDF()
    pdf.add_page()
    
    # --- LOGO PROFISSIONAL ---
    pdf.set_fill_color(0, 51, 102) # Azul Marinho
    pdf.rect(15, 15, 15, 15, 'F')  # Quadrado do Logo
    
    pdf.set_text_color(255, 255, 255) # Branco
    pdf.set_font("Arial", 'B', 12)
    pdf.text(18, 25, "LL") # Iniciais Laurindo Logística

    # CABEÇALHO DA EMPRESA
    pdf.set_xy(35, 15)
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
    
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(255, 128, 0) # Laranja
    pdf.set_xy(35, 22)
    pdf.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
    
    pdf.ln(15)
    pdf.line(10, 40, 200, 40) # Linha divisória
    pdf.ln(10)

    # --- TABELA DE DADOS ---
    pdf.set_font("Arial", 'B', 12)
    pdf.set_fill_color(0, 51, 102) # Topo da Tabela Azul
    pdf.set_text_color(255, 255, 255) # Texto Branco
    pdf.cell(110, 10, " Destino / Localizacao", border=1, fill=True)
    pdf.cell(40, 10, "Custo (Kz)", border=1, ln=True, fill=True, align='C')
    
    pdf.set_font("Arial", '', 11)
    pdf.set_text_color(0, 0, 0) # Texto Preto para os dados
    for _, row in df.iterrows():
        pdf.cell(110, 10, f" {str(row[col_nome])}", border=1)
        pdf.cell(40, 10, f"{row[col_custo]:,.2f}", border=1, ln=True, align='C')
    
    # --- ESPAÇO PARA ASSINATURA ---
    pdf.ln(30) # Espaço para a assinatura
    pdf.set_draw_color(0, 0, 0)
    pdf.line(60, pdf.get_y(), 150, pdf.get_y()) # A Linha de assinatura
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(200, 10, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    
    pdf.set_font("Arial", 'I', 8)
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    pdf.cell(200, 5, f"Documento Gerado em Luanda - {data_hoje}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(img_nome, pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO OFICIAL: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Ola Laurindo, segue o relatorio logistico com assinatura e logo oficial.", 'plain'))
    
    for arq in [pdf_nome, img_nome]:
        if os.path.exists(arq):
            with open(arq, "rb") as f:
                if arq.endswith('.pdf'): anexo = MIMEApplication(f.read(), _subtype="pdf")
                else: anexo = MIMEImage(f.read())
                anexo.add_header('Content-Disposition', 'attachment', filename=arq)
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
            plt.figure(figsize=(10, 6))
            plt.bar(caros['Endereço'].str[:15], caros[col_custo], color='orange')
            plt.savefig('grafico.png')
            
            # Chama a função que cria o PDF com tudo o que pediste
            criar_pdf_executivo(caros, 'Endereço', col_custo, "Relatorio_Oficial_Laurindo.pdf")
            enviar_email('grafico.png', "Relatorio_Oficial_Laurindo.pdf")
            print("Sucesso: Relatorio Completo Enviado!")
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    verificar_e_enviar_tudo()
