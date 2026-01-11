import pandas as pd
import os
import smtplib
import matplotlib.pyplot as plt
from fpdf import FPDF
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

def criar_pdf(df, col_nome, col_custo, nome_pdf):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, "Relatorio de Logistica - Luanda", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(110, 10, "Destino", border=1)
    pdf.cell(40, 10, "Custo (Kz)", border=1, ln=True)
    pdf.set_font("Arial", '', 12)
    for _, row in df.iterrows():
        pdf.cell(110, 10, str(row[col_nome]), border=1)
        pdf.cell(40, 10, str(row[col_custo]), border=1, ln=True)
    pdf.output(nome_pdf)

def enviar_email_v2(img_nome, pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA').replace(" ", "")
    destinatario = "laurics10@gmail.com"

    msg = MIMEMultipart()
    msg['Subject'] = "📊 AGORA COM GRAFICO: Logistica Luanda"
    msg['From'] = meu_email
    msg['To'] = destinatario
    msg.attach(MIMEText("Ola Laurindo, seguem os dois anexos separados abaixo.", 'plain'))

    # ANEXAR PDF (Usando formato Application)
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            pdf_anexo = MIMEApplication(f.read(), _subtype="pdf")
            pdf_anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(pdf_anexo)

    # ANEXAR GRAFICO (Usando formato Image para o Gmail aceitar)
    if os.path.exists(img_nome):
        with open(img_nome, "rb") as f:
            img_anexo = MIMEImage(f.read())
            img_anexo.add_header('Content-Disposition', 'attachment', filename=img_nome)
            msg.attach(img_anexo)

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
            # 1. Gerar Grafico (Fundo branco para o Gmail ler melhor)
            plt.figure(figsize=(10, 6), facecolor='white')
            plt.bar(caros['Endereço'].str[:15], caros[col_custo], color='orange')
            plt.title('Custos Elevados de Logistica em Luanda')
            plt.ylabel('Preco (Kz)')
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig('grafico_luanda.png')
            
            # 2. Gerar PDF
            criar_pdf(caros, 'Endereço', col_custo, "relatorio_luanda.pdf")
            
            # 3. Enviar
            enviar_email_v2('grafico_luanda.png', 'relatorio_luanda.pdf')
            print("!!! SUCESSO TOTAL !!!")
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    verificar_e_enviar_tudo()
