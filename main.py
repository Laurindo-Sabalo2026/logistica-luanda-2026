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
        self.set_font("Arial", 'I', 9)
        self.set_text_color(100, 100, 100)
        self.set_xy(30, 17)
        self.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf_perfeito(df, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # Ordem exata das colunas baseada na tua planilha
    # 0: Endereço, 1: Bairro, 2: Custo, 3: Status, 4: Motorista, 5: Data, 6: Obs
    
    total_kz = df.iloc[:, 2].sum() # Coluna Custo é a terceira (índice 2)
    
    # --- CABEÇALHO DA TABELA ---
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    cols = [("Destino", 55), ("Custo (Kz)", 30), ("(%)", 15), ("Status", 25), ("Motorista", 35), ("Data Entr.", 25), ("Obs", 92)]
    for nome, largura in cols:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    # --- DADOS (MAPEAMENTO MANUAL PARA NÃO FALHAR) ---
    pdf.set_font("Arial", '', 8)
    for i in range(len(df)):
        linha = df.iloc[i]
        custo = float(linha[2])
        perc = (custo / total_kz * 100) if total_kz > 0 else 0
        status = str(linha[3]).strip()
        
        # Cores
        if status.lower() in ['ok', 'concluido', 'concluído']:
            pdf.set_text_color(0, 128, 0) # Verde
        elif status.lower() in ['atrasado', 'atraso']:
            pdf.set_text_color(200, 0, 0) # Vermelho
        else:
            pdf.set_text_color(0, 0, 0) # Preto
            
        pdf.cell(55, 8, f" {str(linha[0])[:32]}", border=1) # Destino
        pdf.cell(30, 8, f"{custo:,.2f}", border=1, align='C') # Custo
        pdf.cell(15, 8, f"{perc:.1f}%", border=1, align='C') # %
        pdf.cell(25, 8, f" {status}", border=1, align='C') # Status
        pdf.cell(35, 8, f" {str(linha[4])[:20]}", border=1, align='C') # Motorista
        pdf.cell(25, 8, f" {str(linha[5])}", border=1, align='C') # Data
        pdf.cell(92, 8, f" {str(linha[6])[:55]}", border=1, ln=True) # Obs

    # --- TOTAL ---
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(55, 10, " TOTAL GERAL:", border=1, fill=True, align='R')
    pdf.set_text_color(200, 0, 0)
    pdf.cell(30, 10, f"{total_kz:,.2f}", border=1, fill=True, align='C')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(15, 10, "100%", border=1, fill=True, align='C')
    pdf.cell(177, 10, " Kwanzas (Relatorio de Gestao)", border=1, ln=True, fill=True)

    # --- RODAPÉ ---
    pdf.ln(15)
    y = pdf.get_y()
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=15, y=y, w=115)
    
    # Assinatura
    pdf.set_xy(180, y + 15)
    pdf.line(180, y + 15, 270, y + 15)
    pdf.set_xy(180, y + 17)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 5, "Laurindo Sabalo", ln=True, align='C')
    
    # Data de Geração (BEM ABAIXO)
    pdf.set_xy(180, y + 30)
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(120, 120, 120)
    agora = datetime.now().strftime('%d/%m/%Y %H:%M')
    pdf.cell(90, 5, f"Relatorio extraido em Luanda: {agora}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    msg = MIMEMultipart()
    msg['Subject'] = f"✅ RELATORIO FINAL: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Relatorio corrigido com mapeamento manual de colunas.", 'plain'))
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())

def executar():
    try:
        df = pd.read_excel("meus_locais (1).xlsx")
        # Gráfico
        plt.figure(figsize=(7, 3.5))
        plt.bar(df.iloc[:,0].str[:10], df.iloc[:,2], color='#2E8B57') 
        plt.tight_layout()
        plt.savefig('grafico.png')
        plt.close()
        
        criar_pdf_perfeito(df, "Relatorio_Final_Laurindo.pdf", "grafico.png")
        enviar_email("Relatorio_Final_Laurindo.pdf")
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
