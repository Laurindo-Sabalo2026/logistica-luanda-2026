import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime

def gerar_pdf_final(df):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(0, 15, "LAURINDO LOGISTICA - RELATORIO DE CONTROLE", ln=True, align='C')
    pdf.ln(5)
    
    # Tabela com mapeamento fixo (Coluna 0, 2, 4, 5)
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(220, 230, 241)
    pdf.cell(70, 10, " Destino", 1, 0, 'L', True)
    pdf.cell(40, 10, " Custo (Kz)", 1, 0, 'C', True)
    pdf.cell(50, 10, " Motorista", 1, 0, 'C', True)
    pdf.cell(40, 10, " Data Entr.", 1, 1, 'C', True)
    
    pdf.set_font("Arial", '', 9)
    for i in range(len(df)):
        linha = df.iloc[i]
        pdf.cell(70, 9, f" {str(linha.iloc[0])[:35]}", 1)
        pdf.cell(40, 9, f" {str(linha.iloc[2])}", 1, 0, 'C')
        pdf.cell(50, 9, f" {str(linha.iloc[4])[:20]}", 1, 0, 'C')
        pdf.cell(40, 9, f" {str(linha.iloc[5])}", 1, 1, 'C')

    pdf.ln(20)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 10, "________________________________", ln=True, align='R')
    pdf.cell(0, 5, "Laurindo Sabalo    ", ln=True, align='R')
    pdf.set_font("Arial", 'I', 8)
    # A LINHA DA DATA QUE TINHA DESAPARECIDO
    data_extracao = datetime.now().strftime('%d/%m/%Y %H:%M')
    pdf.cell(0, 10, f"Documento extraido em Luanda: {data_extracao}", ln=True, align='R')
    
    # Nome do arquivo mudado para evitar bloqueio
    nome_arquivo = f"Doc_Logistica_{datetime.now().strftime('%H%M')}.pdf"
    pdf.output(nome_arquivo)
    return nome_arquivo

def executar():
    try:
        df = pd.read_excel("meus_locais (1).xlsx")
        pdf_gerado = gerar_pdf_final(df)
        
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        # NOVO DESTINATARIO (Corrigido para o formato Gmail)
        destinatario_novo = "laurids10@gmail.com" 
        
        msg = MIMEMultipart()
        msg['Subject'] = f"ENVIO PRIORITARIO - REF {datetime.now().strftime('%M%S')}"
        msg['From'] = meu_email
        msg['To'] = destinatario_novo
        
        msg.attach(MIMEText("Segue o documento de logistica atualizado.", 'plain'))
        
        with open(pdf_gerado, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=pdf_gerado)
            msg.attach(part)
            
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, destinatario_novo, msg.as_string())
        print("SUCESSO: Email enviado para o novo destino!")
        
    except Exception as e:
        print(f"ERRO: {e}")

if __name__ == "__main__":
    executar()
