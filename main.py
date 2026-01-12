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

def criar_pdf_final_oficial(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    total_geral = df[col_custo].sum()
    
    # --- CABEÇALHO DA TABELA ---
    pdf.set_font("Arial", 'B', 8)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    larguras = [("Destino", 50), ("Custo (Kz)", 30), ("(%)", 15), ("Status", 25), ("Motorista", 35), ("Data Entr.", 30), ("Obs", 92)]
    for nome, largura in larguras:
        pdf.cell(largura, 10, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    # --- LINHAS DA TABELA COM ALERTAS DE CORES ---
    pdf.set_font("Arial", '', 8)
    for _, row in df.iterrows():
        custo_v = float(row[col_custo]) if pd.notna(row[col_custo]) else 0
        perc = (custo_v / total_geral) * 100 if total_geral > 0 else 0
        
        # Lógica de Cores para o Status
        status_raw = str(row.get('Status', 'Pendente')).strip()
        status_lower = status_raw.lower()
        
        # 1. Primeiro escrevemos as colunas normais (Preto)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(50, 8, f" {str(row[col_nome])[:25]}", border=1)
        pdf.cell(30, 8, f"{custo_v:,.2f}", border=1, align='C')
        pdf.cell(15, 8, f"{perc:.1f}%", border=1, align='C')
        
        # 2. Aplicamos a cor APENAS na célula do Status
        if status_lower in ['ok', 'concluido', 'concluído', 'entregue']:
            pdf.set_text_color(0, 128, 0) # VERDE para sucesso
        elif status_lower in ['atrasado', 'atraso', 'urgente', 'erro']:
            pdf.set_text_color(200, 0, 0) # VERMELHO para alertas
        else:
            pdf.set_text_color(0, 0, 0) # PRETO para o resto
            
        pdf.cell(25, 8, f" {status_raw}", border=1, align='C')
        
        # 3. Voltamos para Preto para o resto da linha
        pdf.set_text_color(0, 0, 0)
        pdf.cell(35, 8, f" {str(row.get('Motorista', 'N/A'))[:18]}", border=1, align='C')
        pdf.cell(30, 8, f" {str(row.get('Data', '12/01/2026'))[:10]}", border=1, align='C')
        pdf.cell(92, 8, f" {str(row.get('Obs', 'Relatorio de Luanda'))[:55]}", border=1, ln=True)

    # --- TOTAL GERAL ---
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(50, 10, " TOTAL GERAL:", border=1, fill=True, align='R')
    pdf.set_text_color(200, 0, 0) 
    pdf.cell(30, 10, f"{total_geral:,.2f}", border=1, fill=True, align='C')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(15, 10, "100%", border=1, fill=True, align='C')
    pdf.cell(182, 10, " Kwanzas (Alerta de Gestão Luanda)", border=1, ln=True, fill=True)

    # --- GRÁFICO E ASSINATURA ---
    pdf.ln(5)
    y_pos = pdf.get_y()
    if os.path.exists(caminho_grafico):
        pdf.image(caminho_grafico, x=15, y=y_pos, w=125)
    
    pdf.set_xy(180, y_pos + 10)
    pdf.line(180, y_pos + 15, 270, y_pos + 15)
    pdf.set_xy(180, y_pos + 17)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(90, 5, "Laurindo Sabalo", ln=True, align='C')
    
    pdf.set_xy(180, y_pos + 25) 
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(90, 5, f"Relatorio gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurinds10@gmail.com"
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO COM ALERTAS DE CORES - {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'] = meu_email
    msg['To'] = destinatario
    msg.attach(MIMEText("Relatorio oficial com sistema de cores (Verde/Vermelho) no Status.", 'plain'))
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, destinatario, msg.as_string())

def executar():
    excel = "dados.xlsx" if os.path.exists("dados.xlsx") else "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_nome = df.columns[0]
        col_custo = next((c for c in df.columns if 'Custo' in c or 'custo' in c), None)
        
        plt.figure(figsize=(7, 3.2))
        plt.bar(df[col_nome].astype(str).str[:10], df[col_custo], color='#2E8B57') 
        plt.title('Custos de Transporte - Luanda')
        plt.tight_layout()
        plt.savefig('grafico_alerta.png')
        plt.close()
        
        nome_pdf = "Relatorio_Laurindo_Alertas.pdf"
        criar_pdf_final_oficial(df, col_nome, col_custo, nome_pdf, 'grafico_alerta.png')
        enviar_email(nome_pdf)
        print("CONCLUÍDO COM ALERTAS!")
    except Exception as e: print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
