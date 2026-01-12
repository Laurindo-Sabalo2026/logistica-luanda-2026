import pandas as pd
import os
import smtplib
import matplotlib.pyplot as plt
from fpdf import FPDF
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

# 1. Configuração do PDF
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'RELATORIO DE LOGISTICA - LAURINDO', 0, 1, 'C')
        self.ln(5)

def criar_pdf(df, nome_pdf, grafico):
    pdf = PDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 10)
    
    # Cabeçalho da Tabela
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(60, 10, 'Destino', 1, 0, 'C', True)
    pdf.cell(40, 10, 'Custo', 1, 0, 'C', True)
    pdf.cell(40, 10, 'Status', 1, 0, 'C', True)
    pdf.cell(50, 10, 'Motorista', 1, 0, 'C', True)
    pdf.cell(40, 10, 'Data', 1, 1, 'C', True)

    # Linhas da Tabela
    pdf.set_font("Arial", '', 10)
    for i in range(len(df)):
        linha = df.iloc[i]
        pdf.cell(60, 10, str(linha.iloc[0])[:30], 1)
        pdf.cell(40, 10, str(linha.iloc[2]), 1, 0, 'C')
        pdf.cell(40, 10, str(linha.iloc[3]), 1, 0, 'C')
        pdf.cell(50, 10, str(linha.iloc[4]), 1, 0, 'C')
        pdf.cell(40, 10, str(linha.iloc[5]), 1, 1, 'C')

    # Inserir Gráfico
    if os.path.exists(grafico):
        pdf.ln(10)
        pdf.image(grafico, x=10, w=160)
    
    pdf.output(nome_pdf)

def executar():
    try:
        # Carregar Excel
        df = pd.read_excel("meus_locais (1).xlsx")
        
        # Criar Gráfico Simples
        plt.figure(figsize=(8, 4))
        plt.bar(df.iloc[:, 0].astype(str).str[:10], df.iloc[:, 2], color='green')
        plt.title('Custos por Destino')
        plt.savefig('grafico.png')
        plt.close()

        nome_pdf = "Relatorio_Laurindo_Antigo.pdf"
        criar_pdf(df, nome_pdf, 'grafico.png')

        # Configuração de E-mail
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        destino = "laurinds10@gmail.com"

        msg = MIMEMultipart()
        msg['Subject'] = f"Relatorio Logistica - {datetime.now().strftime('%H:%M')}"
        msg['From'] = meu_email
        msg['To'] = destino
        msg.attach(MIMEText("Segue o relatorio na versao anterior que funcionou bem.", 'plain'))

        with open(nome_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(anexo)

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(meu_email, senha)
        server.sendmail(meu_email, destino, msg.as_string())
        server.quit()
        print("Enviado!")

    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
