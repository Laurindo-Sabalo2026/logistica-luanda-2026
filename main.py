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
    
    # --- CABEÇALHO (AZUL ROYAL) ---
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

    # --- TABELA COM 3 COLUNAS ---
    pdf.set_font("Arial", 'B', 11)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    # Larguras: Destino(70), Custo(40), Status(40)
    pdf.cell(70, 10, " Destino / Localizacao", border=1, fill=True)
    pdf.cell(40, 10, "Custo (Kz)", border=1, fill=True, align='C')
    pdf.cell(40, 10, "Status", border=1, ln=True, fill=True, align='C')
    
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        # Destino
        pdf.cell(70, 10, f" {str(row[col_nome])[:30]}", border=1)
        # Custo
        pdf.cell(40, 10, f"{row[col_custo]:,.2f}", border=1, align='C')
        # Status (Lógica para não falhar se a coluna não existir no Excel)
        val_status = str(row['Status']) if 'Status' in df.columns else "Pendente"
        pdf.cell(40, 10, f" {val_status}", border=1, ln=True, align='C')
    
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
    msg['Subject'] = f"📊 RELATORIO LOGISTICA FINAL: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Relatorio oficial com 3 colunas e layout corrigido.", 'plain'))
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
    if not os.path.exists(excel):
        print("Erro: Arquivo Excel nao encontrado.")
        return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Identifica a coluna de custo
        cols_custo = [c for c in df.columns if 'Custo' in c]
        if not cols_custo:
            print("Erro: Coluna de Custo nao encontrada.")
            return
        col_custo = cols_custo[0]
        
        # Filtra e gera gráfico
        caros = df[df[col_custo] > 100]
        if not caros.empty:
            plt.figure(figsize=(8, 4))
            plt.bar(caros['Endereço'].str[:15], caros[col_custo], color='royalblue') 
            plt.title('Custos de Transporte - Luanda')
            plt.savefig('grafico.png')
            plt.close()
            
            nome_pdf = "Relatorio_Final_3Colunas.pdf"
            criar_pdf_final(caros, 'Endereço', col_custo, nome_pdf, 'grafico.png')
            enviar_email(nome_pdf)
            print("Sucesso!")
    except Exception as e:
        print(f"Erro no processamento: {e}")

if __name__ == "__main__":
    executar()
