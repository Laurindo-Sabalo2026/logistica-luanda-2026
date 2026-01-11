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

    # --- CABEÇALHO ---
    pdf.set_xy(35, 15)
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
    
    # --- SLOGAN (RECUPERADO) ---
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.set_xy(35, 22)
    pdf.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
    
    # --- LINHA COMPLETA DIREITA ---
    pdf.ln(8)
    pdf.set_draw_color(200, 200, 200) 
    pdf.line(15, 35, 195, 35) 
    pdf.ln(10)

    # --- DETECTAR COLUNA STATUS ---
    tem_status = 'Status' in df.columns
    larg_destino = 80 if tem_status else 115 # Ajusta largura se houver status

    # --- CABEÇALHO TABELA ---
    pdf.set_font("Arial", 'B', 11)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    pdf.cell(larg_destino, 10, " Destino", border=1, fill=True)
    pdf.cell(35, 10, "Custo (Kz)", border=1, fill=True, align='C')
    if tem_status:
        pdf.cell(35, 10, "Status", border=1, ln=True, fill=True, align='C')
    else:
        pdf.ln(10)
    
    # --- DADOS ---
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    for _, row in df.iterrows():
        pdf.cell(larg_destino, 10, f" {str(row[col_nome])[:35]}", border=1)
        pdf.cell(35, 10, f"{row[col_custo]:,.2f}", border=1, align='C')
        if tem_status:
            pdf.cell(35, 10, f"{str(row['Status'])}", border=1, ln=True, align='C')
        else:
            pdf.ln(10)
    
    # --- GRÁFICO VERDE ---
    if os.path.exists(caminho_grafico):
        pdf.ln(5)
        pdf.image(caminho_grafico, x=20, w=160)
    
    # --- ASSINATURA E DATA (REPOSICIONADAS) ---
    pdf.set_y(-35) 
    pdf.set_draw_color(0, 0, 0)
    pdf.line(60, pdf.get_y(), 150, pdf.get_y())
    pdf.ln(2)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 8, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    pdf.set_font("Arial", 'I', 9)
    pdf.cell(0, 8, f"Gerado em Luanda - Data: {datetime.now().strftime('%d/%m/%Y')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO LOGISTICA: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Relatorio corrigido com sucesso.", 'plain'))
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
        caros = df[df[col_custo] > 100]
        if not caros.empty:
            plt.figure(figsize=(10, 5))
            plt.bar(caros['Endereço'].str[:15], caros[col_custo], color='seagreen') 
            plt.title('Analise de Custos de Logistica')
            plt.savefig('grafico.png')
            plt.close()
            nome_pdf = "Relatorio_Oficial.pdf"
            criar_pdf_final(caros, 'Endereço', col_custo, nome_pdf, 'grafico.png')
            enviar_email(nome_pdf)
            print("Sucesso!")
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
