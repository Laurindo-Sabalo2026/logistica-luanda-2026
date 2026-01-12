import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime

# Criar PDF sem firulas para não travar
def criar_pdf_simples(df, nome_pdf):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "LAURINDO LOGISTICA - RELATORIO DE EMERGENCIA", ln=True, align='C')
    pdf.ln(10)
    
    # Cabeçalho básico
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(80, 10, " Destino", 1, 0, 'L', True)
    pdf.cell(40, 10, " Custo (Kz)", 1, 0, 'C', True)
    pdf.cell(50, 10, " Motorista", 1, 0, 'C', True)
    pdf.cell(40, 10, " Data", 1, 1, 'C', True)
    
    # Dados puros do Excel (Forçando a posição das colunas)
    pdf.set_font("Arial", '', 10)
    for i in range(len(df)):
        linha = df.iloc[i]
        pdf.cell(80, 10, f" {str(linha.iloc[0])[:35]}", 1)
        pdf.cell(40, 10, f" {str(linha.iloc[2])}", 1, 0, 'C')
        pdf.cell(50, 10, f" {str(linha.iloc[4])}", 1, 0, 'C')
        pdf.cell(40, 10, f" {str(linha.iloc[5])}", 1, 1, 'C')

    # Assinatura e Data de Geração
    pdf.ln(20)
    pdf.cell(0, 10, "__________________________", ln=True, align='R')
    pdf.cell(0, 10, "Laurindo Sabalo    ", ln=True, align='R')
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='R')
    pdf.output(nome_pdf)

def executar():
    try:
        # 1. Ler o Excel
        df = pd.read_excel("meus_locais (1).xlsx")
        nome_pdf = "Relatorio_Laurindo_Urgente.pdf"
        
        # 2. Criar o PDF
        criar_pdf_simples(df, nome_pdf)
        
        # 3. Enviar e-mail (Sem validações complexas)
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        msg['Subject'] = "RELATORIO RECUPERADO - LAURINDO"
        msg['From'] = meu_email
        msg['To'] = "laurics10@gmail.com"
        msg.attach(MIMEText("Segue o relatorio em modo de seguranca.", 'plain'))
        
        with open(nome_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(anexo)
            
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())
        print("Enviado com sucesso!")

    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
