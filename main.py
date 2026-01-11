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
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", 'B', 12)
        self.text(13, 20, "LL")
        self.set_xy(30, 10)
        self.set_font("Arial", 'B', 16)
        self.set_text_color(0, 102, 204)
        self.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
        self.set_font("Arial", 'I', 9)
        self.set_text_color(100, 100, 100)
        self.set_xy(30, 17)
        self.cell(100, 5, "Monitoramento de Carga e Custos - Luanda", ln=True)
        self.line(10, 30, 287, 30)
        self.ln(10)

def criar_pdf_versao_final(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    total_geral = df[col_custo].sum()
    
    # --- TABELA ---
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    larguras = [("Destino", 50), ("Custo (Kz)", 30), ("(%)", 15), ("Status", 25), ("Motorista", 35), ("Data Entr.", 30), ("Obs", 92)]
    for nome, largura in larguras:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    for _, row in df.head(10).iterrows():
        custo_v = float(row[col_custo])
        perc = (custo_v / total_geral) * 100 if total_geral > 0 else 0
        status_raw = str(row.get('Status', 'Pendente')).strip().lower()
        
        if status_raw in ['ok', 'concluído', 'concluido']:
            pdf.set_text_color(0, 128, 0)
            txt_status = "Concluído"
        elif status_raw in ['atrasado', 'atraso']:
            pdf.set_text_color(200, 0, 0)
            txt_status = "Atrasado"
        else:
            pdf.set_text_color(0, 0, 0)
            txt_status = "Pendente"

        pdf.cell(50, 8, f" {str(row[col_nome])[:25]}", border=1)
        pdf.cell(30, 8, f"{custo_v:,.2f}", border=1, align='C')
        pdf.cell(15, 8, f"{perc:.1f}%", border=1, align='C')
        pdf.cell(25, 8, f" {txt_status}", border=1, align='C')
        pdf.cell(35, 8, f" {str(row.get('Motorista', 'N/A'))[:18]}", border=1, align='C')
        pdf.cell(30, 8, f" {str(row.get('Data Entr.', '12/01/2026'))}", border=1, align='C')
        pdf.cell(92, 8, f" {str(row.get('Obs', ''))[:55]}", border=1, ln=True)

    # --- LINHA TOTAL ---
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(240, 240, 240)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(50, 10, " TOTAL GERAL:", border=1, fill=True, align='R')
    pdf.set_text_color(200, 0, 0)
    pdf.cell(30, 10, f"{total_geral:,.2f}", border=1, fill=True, align='C')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(15, 10, "100%", border=1, fill=True, align='C')
    pdf.cell(182, 10, " Kwanzas (Relatorio de Luanda)", border=1, ln=True, fill=True)

    # --- ÁREA INFERIOR (REAJUSTADA PARA CABER TUDO) ---
    pdf.ln(5) # Reduzi o salto para ganhar espaço no fundo
    y_pos = pdf.get_y()
    
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=15, y=y_pos, w=125) # Gráfico ligeiramente menor
    
    # Assinatura (Posicionada mais para cima)
    pdf.set_xy(180, y_pos + 10)
    pdf.set_draw_color(0, 0, 0)
    pdf.line(180, y_pos + 15, 270, y_pos + 15)
    
    pdf.set_xy(180, y_pos + 17)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 5, "Laurindo Sabalo", ln=True, align='C')
    
    # --- A LINHA DA DATA (FORÇADA NO RODAPÉ) ---
    pdf.set_xy(180, y_pos + 23) # Nova posição garantida
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(100, 100, 100)
    data_formatada = datetime.now().strftime('%d/%m/%Y %H:%M')
    pdf.cell(90, 5, f"Relatorio gerado em: {data_formatada}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO COMPLETO: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Segue o relatorio com a data de rodapé corrigida.", 'plain'))
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, destinatario, msg.as_string())

def executar():
    excel = "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_custo = [c for c in df.columns if 'Custo' in c][0]
        
        plt.figure(figsize=(7, 3.2)) # Reduzi a altura do gráfico
        plt.bar(df['Endereço'].str[:10], df[col_custo], color='#2E8B57') 
        plt.title('Custos de Transporte')
        plt.tight_layout()
        plt.savefig('grafico_v3.png')
        plt.close()
        
        nome_pdf = "Relatorio_Final_Corrigido.pdf"
        criar_pdf_versao_final(df, 'Endereço', col_custo, nome_pdf, 'grafico_v3.png')
        enviar_email(nome_pdf)
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
