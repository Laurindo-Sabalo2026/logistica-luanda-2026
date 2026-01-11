import pandas as pd
import os
import smtplib
import matplotlib.pyplot as plt
from fpdf import FPDF
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

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

def enviar_email_com_anexos(lista_anexos):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA').replace(" ", "")
    msg = MIMEMultipart()
    msg['Subject'] = "📊 RELATORIO COMPLETO: Logistica Luanda"
    msg.attach(MIMEText("Ola Laurindo, seguem em anexo o GRAFICO e o PDF com os dados.", 'plain'))
    
    # CORREÇÃO AQUI: Agora ele percorre cada anexo corretamente
    for nome_ficheiro in lista_anexos:
        if os.path.exists(nome_ficheiro):
            with open(nome_ficheiro, "rb") as anexo:
                p = MIMEBase("application", "octet-stream")
                p.set_payload(anexo.read())
            encoders.encode_base64(p)
            p.add_header("Content-Disposition", f"attachment; filename={nome_ficheiro}")
            msg.attach(p)
            
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())

def verificar_e_enviar_tudo():
    excel = "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_custo = [c for c in df.columns if 'Custo' in c][0]
        caros = df[df[col_custo] > 100]

        if not caros.empty:
            # 1. Gerar Grafico
            plt.figure(figsize=(8, 5))
            plt.bar(caros['Endereço'].str[:12], caros[col_custo], color='orange')
            plt.title('Custos Elevados Luanda')
            plt.tight_layout() # Garante que o gráfico não saia cortado
            plt.savefig('grafico.png')
            
            # 2. Gerar PDF
            criar_pdf(caros, 'Endereço', col_custo, "relatorio.pdf")
            
            # 3. Enviar Email com a lista de ficheiros
            enviar_email_com_anexos(["grafico.png", "relatorio.pdf"])
            print("!!! TUDO ENVIADO COM SUCESSO !!!")
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    verificar_e_enviar_tudo()
