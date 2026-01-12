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
        self.cell(100, 5, "Monitoramento de Operacoes - Luanda", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf_blindado(df, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # FORÇAR A LEITURA DAS COLUNAS PELO ÍNDICE (0=A, 1=B, 2=C...)
    # Baseado no teu Excel: A=Destino, C=Custo, D=Status, E=Motorista, F=Data
    
    total_kz = pd.to_numeric(df.iloc[:, 2], errors='coerce').sum()
    
    # --- CABEÇALHO ---
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    cols = [("Destino", 55), ("Custo (Kz)", 30), ("(%)", 15), ("Status", 25), ("Motorista", 35), ("Data Entr.", 30), ("Obs", 87)]
    for nome, largura in cols:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    # --- DADOS ---
    pdf.set_font("Arial", '', 8)
    for i in range(len(df)):
        linha = df.iloc[i]
        custo = pd.to_numeric(linha[2], errors='coerce') or 0
        perc = (custo / total_kz * 100) if total_kz > 0 else 0
        status = str(linha[3]).strip()
        
        # Cor do Status
        if status.lower() in ['ok', 'concluido', 'concluído']:
            pdf.set_text_color(0, 128, 0) # Verde
        elif status.lower() in ['atrasado', 'atraso']:
            pdf.set_text_color(200, 0, 0) # Vermelho
        else:
            pdf.set_text_color(0, 0, 0) # Preto

        # MAPEAMENTO MANUAL PARA EVITAR TROCAS:
        pdf.cell(55, 8, f" {str(linha[0])[:30]}", border=1)      # COLUNA A (Destino)
        pdf.cell(30, 8, f"{custo:,.2f}", border=1, align='C')   # COLUNA C (Custo)
        pdf.cell(15, 8, f"{perc:.1f}%", border=1, align='C')    # Calculado
        pdf.cell(25, 8, f" {status}", border=1, align='C')      # COLUNA D (Status)
        pdf.cell(35, 8, f" {str(linha[4])[:20]}", border=1, align='C') # COLUNA E (Motorista)
        pdf.cell(30, 8, f" {str(linha[5])}", border=1, align='C') # COLUNA F (Data)
        pdf.cell(87, 8, f" {str(linha[6])[:50]}", border=1, ln=True) # COLUNA G (Obs)

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

    # --- RODAPÉ (DATA DE GERAÇÃO) ---
    pdf.ln(10)
    y_final = pdf.get_y()
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=15, y=y_final, w=110)
    
    # Assinatura e Data de Geracao Final
    pdf.set_xy(180, y_final + 10)
    pdf.line(180, y_final + 15, 270, y_final + 15)
    pdf.set_xy(180, y_final + 17)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 5, "Laurindo Sabalo", ln=True, align='C')
    
    pdf.set_xy(180, y_final + 28) # ESPAÇO PARA A LINHA DA DATA
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(120, 120, 120)
    data_agora = datetime.now().strftime('%d/%m/%Y %H:%M')
    pdf.cell(90, 5, f"Relatorio gerado em Luanda: {data_agora}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def executar():
    try:
        df = pd.read_excel("meus_locais (1).xlsx")
        
        # Gráfico seguro
        plt.figure(figsize=(7, 3.5))
        plt.bar(df.iloc[:,0].str[:10], pd.to_numeric(df.iloc[:,2], errors='coerce'), color='#2E8B57') 
        plt.tight_layout()
        plt.savefig('grafico.png')
        plt.close()
        
        criar_pdf_blindado(df, "Relatorio_Final_Laurindo.pdf", "grafico.png")
        
        # Envio de Email
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        msg = MIMEMultipart()
        msg['Subject'] = f"✅ RELATORIO FINALIZADO - {datetime.now().strftime('%d/%m/%Y')}"
        msg.attach(MIMEText("Segue o relatorio com as colunas e data corrigidas.", 'plain'))
        with open("Relatorio_Final_Laurindo.pdf", "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename="Relatorio_Final_Laurindo.pdf")
            msg.attach(anexo)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())
            
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
