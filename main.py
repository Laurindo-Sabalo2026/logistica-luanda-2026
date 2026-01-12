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
        self.cell(100, 5, "Gestao de Operacoes e Custos - Luanda", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf_seguro(df, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # Limpeza de dados para evitar erro de soma
    custos_limpos = pd.to_numeric(df.iloc[:, 2], errors='coerce').fillna(0)
    total_kz = custos_limpos.sum()
    
    # Cabeçalho da Tabela
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    cols = [("Destino", 60), ("Custo (Kz)", 30), ("(%)", 15), ("Status", 25), ("Motorista", 35), ("Data Entr.", 30), ("Obs", 82)]
    for nome, largura in cols:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    # Preenchimento Rigido das Colunas (A, C, D, E, F, G)
    pdf.set_font("Arial", '', 8)
    for i in range(len(df)):
        linha = df.iloc[i]
        valor_custo = float(custos_limpos.iloc[i])
        percentual = (valor_custo / total_kz * 100) if total_kz > 0 else 0
        status_txt = str(linha[3]).strip() if len(linha) > 3 else "N/A"

        # Cor do Status
        if status_txt.lower() in ['ok', 'concluido', 'concluído']:
            pdf.set_text_color(0, 128, 0)
        elif status_txt.lower() in ['atrasado', 'atraso']:
            pdf.set_text_color(200, 0, 0)
        else:
            pdf.set_text_color(0, 0, 0)

        # Células mapeadas para não trocar nomes com destinos
        pdf.cell(60, 8, f" {str(linha[0])[:35]}", border=1)             # Coluna A
        pdf.cell(30, 8, f"{valor_custo:,.2f}", border=1, align='C')    # Coluna C
        pdf.cell(15, 8, f"{percentual:.1f}%", border=1, align='C')     # Calculado
        pdf.cell(25, 8, f" {status_txt}", border=1, align='C')         # Coluna D
        pdf.cell(35, 8, f" {str(linha[4])[:20]}", border=1, align='C') # Coluna E
        pdf.cell(30, 8, f" {str(linha[5])}", border=1, align='C')      # Coluna F
        pdf.cell(82, 8, f" {str(linha[6])[:45]}", border=1, ln=True)   # Coluna G

    # Totalizador
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(60, 10, " TOTAL GERAL:", border=1, fill=True, align='R')
    pdf.set_text_color(200, 0, 0)
    pdf.cell(30, 10, f"{total_kz:,.2f}", border=1, fill=True, align='C')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(15, 10, "100%", border=1, fill=True, align='C')
    pdf.cell(172, 10, " Kwanzas (Relatorio Oficial Auditado)", border=1, ln=True, fill=True)

    # Rodapé com Assinatura e Data de Geração
    pdf.ln(10)
    y_sig = pdf.get_y()
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=15, y=y_sig, w=110)
    
    pdf.set_xy(180, y_sig + 15)
    pdf.line(180, y_sig + 15, 270, y_sig + 15)
    pdf.set_xy(180, y_sig + 17)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 5, "Laurindo Sabalo", ln=True, align='C')
    
    # LINHA DA DATA ISOLADA NO FUNDO
    pdf.set_xy(180, y_sig + 32)
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(90, 5, f"Relatorio gerado automaticamente em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def executar_processo():
    try:
        arquivo = "meus_locais (1).xlsx"
        df = pd.read_excel(arquivo)
        
        # Gráfico
        plt.figure(figsize=(7, 3.5))
        plt.bar(df.iloc[:,0].str[:10], pd.to_numeric(df.iloc[:,2], errors='coerce').fillna(0), color='#2E8B57')
        plt.tight_layout()
        plt.savefig('grafico.png')
        plt.close()
        
        nome_pdf = "Relatorio_Final_Corrigido.pdf"
        criar_pdf_seguro(df, nome_pdf, "grafico.png")
        
        # Envio garantido
        meu_email = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        msg = MIMEMultipart()
        msg['Subject'] = f"✅ RELATORIO LOGISTICA: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        msg.attach(MIMEText("Relatorio oficial com colunas e rodapé corrigidos.", 'plain'))
        with open(nome_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=nome_pdf)
            msg.attach(anexo)
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, "laurics10@gmail.com", msg.as_string())
        print("E-mail enviado com sucesso!")
            
    except Exception as e:
        print(f"Erro fatal detectado: {e}")

if __name__ == "__main__":
    executar_processo()
