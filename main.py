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
        self.set_fill_color(0, 102, 204)
        self.rect(10, 10, 15, 15, 'F')
        self.set_xy(30, 10)
        self.set_font("Arial", 'B', 16)
        self.set_text_color(0, 102, 204)
        self.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf(df, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # Cabeçalho da Tabela
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(200, 220, 255)
    # Definimos as colunas exatamente como no seu Excel
    cols = ["Destino", "Bairro", "Custo", "Status", "Motorista", "Data"]
    larguras = [60, 40, 30, 30, 40, 30]
    
    for i, col in enumerate(cols):
        pdf.cell(larguras[i], 10, col, border=1, fill=True, align='C')
    pdf.ln()

    # Dados
    pdf.set_font("Arial", '', 9)
    for i in range(len(df)):
        linha = df.iloc[i]
        # Usamos iloc para garantir que pegamos a coluna pela posição (0, 1, 2...)
        pdf.cell(60, 8, str(linha.iloc[0])[:35], border=1) # Coluna A
        pdf.cell(40, 8, str(linha.iloc[1])[:20], border=1) # Coluna B
        pdf.cell(30, 8, str(linha.iloc[2]), border=1, align='C') # Coluna C
        pdf.cell(30, 8, str(linha.iloc[3]), border=1, align='C') # Coluna D
        pdf.cell(40, 8, str(linha.iloc[4]), border=1, align='C') # Coluna E
        pdf.cell(30, 8, str(linha.iloc[5]), border=1, align='C', ln=True) # Coluna F

    # Gráfico
    pdf.ln(10)
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=10, w=150)
    
    # Assinatura
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 10, "Laurindo Sabalo", ln=True, align='R')
    pdf.output(nome_pdf)

def executar():
    try:
        df = pd.read_excel("meus_locais (1).xlsx")
        
        # Gerar Gráfico
        plt.figure(figsize=(10, 5))
        plt.bar(df.iloc[:, 0].str[:10], df.iloc[:, 2], color='blue')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig('grafico.png')
        plt.close()

        nome_pdf = "Relatorio_Laurindo_Recuperado.pdf"
        criar_pdf(df, nome_pdf, 'grafico.png')

        # Envio de Email
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        # Enviando para o novo destino que sabemos que funciona
        destino = "laurinds10@gmail.com"

        msg = MIMEMultipart()
        msg['Subject'] = "Relatorio de Logistica - Versao Recuperada"
        msg['From'] = meu_email
        msg['To'] = destino
        msg.attach(MIMEText("Segue o relatorio no formato que deu certo as 01:09.", 'plain'))

        with open(nome_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(anexo)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, destino, msg.as_string())
        print("Enviado com sucesso!")

    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
