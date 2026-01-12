import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import random

def executar():
    try:
        # 1. LER O ARQUIVO DADOS.XLSX
        nome_arquivo = "dados.xlsx"
        df = pd.read_excel(nome_arquivo)
        print("Arquivo lido! Criando tabela...")

        # 2. CRIAR PDF
        pdf = FPDF('L', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "RELATORIO DE LOGISTICA - LAURINDO SABALO", ln=True, align='C')
        pdf.ln(10)
        
        # Cabeçalho
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(70, 10, " Destino", 1)
        pdf.cell(40, 10, " Custo", 1)
        pdf.cell(50, 10, " Motorista", 1)
        pdf.cell(50, 10, " Data", 1, 1)
        
        # 3. LER COLUNAS PELO NOME (EVITA O ERRO OUT-OF-BOUNDS)
        pdf.set_font("Arial", '', 10)
        for index, linha in df.iterrows():
            # O código pega o que estiver na 1ª, 3ª, 5ª e 6ª colunas do seu Excel
            destino = str(linha.iloc[0])[:30]
            custo = str(linha.iloc[2]) if len(linha) > 2 else "---"
            motorista = str(linha.iloc[4]) if len(linha) > 4 else "---"
            data = str(linha.iloc[5]) if len(linha) > 5 else "---"
            
            pdf.cell(70, 10, f" {destino}", 1)
            pdf.cell(40, 10, f" {custo}", 1)
            pdf.cell(50, 10, f" {motorista}", 1)
            pdf.cell(50, 10, f" {data}", 1, 1)

        nome_pdf = f"Relatorio_Final_{random.randint(100,999)}.pdf"
        pdf.output(nome_pdf)

        # 4. ENVIAR EMAIL
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        
        msg = MIMEMultipart()
        msg['Subject'] = "RELATORIO DE LOGISTICA ATUALIZADO"
        msg['From'] = meu_email
        msg['To'] = "laurinds10@gmail.com"
        msg.attach(MIMEText("Ola Laurindo, o relatorio com os dados do Excel chegou!", 'plain'))

        with open(nome_pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurinds10@gmail.com", msg.as_string())
        
        print("ENVIADO COM SUCESSO!")

    except Exception as e:
        print(f"Erro Final: {e}")

if __name__ == "__main__":
    executar()
