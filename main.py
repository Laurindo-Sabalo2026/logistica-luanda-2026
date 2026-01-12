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
        self.cell(100, 5, "Gestao de frotas e entregas - Luanda", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf_final(df, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # Forçar conversão de custos para números para não dar erro na soma
    df.iloc[:, 2] = pd.to_numeric(df.iloc[:, 2], errors='coerce').fillna(0)
    total_kz = df.iloc[:, 2].sum()
    
    # --- CABEÇALHO DA TABELA ---
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    cols = [("Destino", 60), ("Custo (Kz)", 30), ("(%)", 15), ("Status", 25), ("Motorista", 35), ("Data Entr.", 30), ("Obs", 82)]
    for nome, largura in cols:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    # --- DADOS (MAPEAMENTO POR POSIÇÃO RÍGIDA) ---
    pdf.set_font("Arial", '', 8)
    for i in range(len(df)):
        linha = df.iloc[i]
        custo = float(linha[2])
        perc = (custo / total_kz * 100) if total_kz > 0 else 0
        status = str(linha[3]).strip()
        
        # Cor do Status
        if status.lower() in ['ok', 'concluido', 'concluído']:
            pdf.set_text_color(0, 128, 0)
        elif status.lower() in ['atrasado', 'atraso']:
            pdf.set_text_color(200, 0, 0)
        else:
            pdf.set_text_color(0, 0, 0)

        # Inserção de dados nas colunas corretas
        pdf.cell(60, 8, f" {str(linha[0])[:35]}", border=1)      # Coluna A (Destino)
        pdf.cell(30, 8, f"{custo:,.2f}", border=1, align='C')   # Coluna C (Custo)
        pdf.cell(15, 8, f"{perc:.1f}%", border=1, align='C')    # Cálculo
        pdf.cell(25, 8, f" {status}", border=1, align='C')      # Coluna D (Status)
        pdf.cell(35, 8, f" {str(linha[4])[:20]}", border=1, align='C') # Coluna E (Motorista)
        pdf.cell(30, 8, f" {str(linha[5])}", border=1, align='C')      # Coluna F (Data)
        pdf.cell(82, 8, f" {str(linha[6])[:45]}", border=1, ln=True)   # Coluna G (Obs)

    # --- TOTAL GERAL ---
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(60, 10, " TOTAL GERAL:", border=1, fill=True, align='R')
    pdf.set_text_color(200, 0, 0)
    pdf.cell(30, 10, f"{total_kz:,.2f}", border=1, fill=True, align='C')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(15, 10, "100%", border=1, fill=True, align='C')
    pdf.cell(172, 10, " Kwanzas (Relatorio Oficial)", border=1, ln=True, fill=True)

    # --- ASSINATURA E DATA DE RODAPÉ ---
    pdf.ln(10)
    y_base = pdf.get_y()
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=15, y=y_base, w=110)
    
    # Assinatura
    pdf.set_xy(180, y_base + 10)
    pdf.line(180, y_base + 15, 270, y_base + 15)
    pdf.set_xy(180, y_base + 17)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 5, "Laurindo Sabalo", ln=True, align='C')
    
    # DATA DE GERAÇÃO (RODAPÉ FINAL)
    pdf.set_xy(180, y_base + 30) # Espaço grande para a data aparecer
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(120, 120, 120)
    data_log = datetime.now().strftime('%d/%m/%Y %H:%M')
    pdf.cell(90, 5, f"Relatorio extraido em: {data_log}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def executar():
    try:
        nome_excel = "meus_locais (1).xlsx"
        if not os.path.exists(nome_excel):
            print("Erro: Arquivo Excel nao encontrado.")
            return

        df = pd.read_excel(nome_excel)
        
        # Gerar Gráfico
        plt.figure(figsize=(7, 3.5))
        plt.bar(df.iloc[:,0].str[:10], pd.to_numeric(df.iloc[:,2], errors='coerce').fillna(0), color='#2E8B57') 
        plt.tight_layout()
        plt.savefig('grafico.png')
        plt.close()
        
        nome_pdf = "Relatorio_Final_Laurindo.pdf"
        criar_pdf_final(df, nome_pdf, "grafico.png")
        
        # Enviar Email
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        destinatario = "laurics10@gmail.com"
        
        msg = MIMEMultipart()
        msg['Subject'] = f"🚀 RELATORIO CONCLUIDO: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        msg.attach(MIMEText("Ola Laurindo, o relatorio foi processado com sucesso.", 'plain'))
        
        with open(nome_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(anexo)
            
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, destinatario, msg.as_string())
        print("Sucesso!")
            
    except Exception as e:
        print(f"Erro no processo: {e}")

if __name__ == "__main__":
    executar()
