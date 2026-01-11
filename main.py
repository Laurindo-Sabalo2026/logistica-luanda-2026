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
    
    # --- LOGOTIPO ---
    pdf.set_fill_color(0, 102, 204) 
    pdf.rect(15, 15, 15, 15, 'F')  
    pdf.set_text_color(255, 255, 255) 
    pdf.set_font("Arial", 'B', 12)
    pdf.text(18, 25, "LL") 

    pdf.set_xy(35, 15)
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
    pdf.ln(15)
    pdf.line(10, 40, 200, 40)
    pdf.ln(10)

    # --- TABELA COM NOVA COLUNA 'STATUS' ---
    pdf.set_font("Arial", 'B', 11)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    # Cabeçalho (Aumentamos para 3 colunas)
    pdf.cell(80, 10, " Destino", border=1, fill=True)
    pdf.cell(35, 10, "Custo (Kz)", border=1, fill=True, align='C')
    pdf.cell(35, 10, "Status", border=1, ln=True, fill=True, align='C')
    
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        pdf.cell(80, 10, f" {str(row[col_nome])[:35]}", border=1)
        pdf.cell(35, 10, f"{row[col_custo]:,.2f}", border=1, align='C')
        # Adicionando o Status da nova coluna do Excel
        status_texto = str(row['Status']) if 'Status' in df.columns else "N/A"
        pdf.cell(35, 10, f"{status_texto}", border=1, ln=True, align='C')
    
    # --- GRÁFICO (AGORA EM VERDE) ---
    if os.path.exists(caminho_grafico):
        pdf.ln(10)
        pdf.image(caminho_grafico, x=10, w=180)
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO ATUALIZADO: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Olá Laurindo, segue o relatório com a nova coluna de Status e gráfico verde.", 'plain'))
    
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
            # --- MUDANÇA DE COR PARA VERDE (color='seagreen') ---
            plt.figure(figsize=(10, 5))
            plt.bar(caros['Endereço'].str[:15], caros[col_custo], color='seagreen') 
            plt.title('Análise de Custos de Logística')
            plt.ylabel('Kwanza (Kz)')
            plt.savefig('grafico_verde.png')
            plt.close()
            
            nome_pdf = "Relatorio_Laurindo_Premium.pdf"
            criar_pdf_final(caros, 'Endereço', col_custo, nome_pdf, 'grafico_verde.png')
            enviar_email(nome_pdf)
            print("Enviado com sucesso!")
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
