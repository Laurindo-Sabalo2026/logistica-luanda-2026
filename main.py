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
        # Logotipo LL
        self.set_fill_color(0, 102, 204)
        self.rect(10, 8, 12, 12, 'F')
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", 'B', 10)
        self.text(12, 16, "LL")
        
        # Título
        self.set_xy(25, 8)
        self.set_font("Arial", 'B', 14)
        self.set_text_color(0, 102, 204)
        self.cell(100, 8, "LAURINDO LOGISTICA & SERVICOS", ln=True)
        
        # Slogan e Linha
        self.set_font("Arial", 'I', 8)
        self.set_text_color(100, 100, 100)
        self.set_xy(25, 14)
        self.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
        self.line(10, 22, 287, 22)
        self.ln(5)

def criar_pdf_premium(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    # 'L' para Landscape (Paisagem)
    pdf = PDF_Logistica('L', 'mm', 'A4')
    pdf.add_page()
    
    # --- TABELA (Ajustada para economizar espaço vertical) ---
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    # Larguras otimizadas
    colunas = [
        ("Destino", 65), ("Custo (Kz)", 30), ("Status", 30), 
        ("Motorista", 40), ("Data Entr.", 30), ("Obs", 82)
    ]
    
    for nome, largura in colunas:
        pdf.cell(largura, 8, f" {nome}", border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", '', 8)
    pdf.set_text_color(0, 0, 0)
    
    # Limitar a 10 linhas para garantir que cabe na página com o gráfico
    df_preview = df.head(10)
    
    for _, row in df_preview.iterrows():
        pdf.cell(65, 7, f" {str(row[col_nome])[:35]}", border=1)
        pdf.cell(30, 7, f"{row[col_custo]:,.2f}", border=1, align='C')
        pdf.cell(30, 7, f" {str(row.get('Status', 'Pendente'))}", border=1, align='C')
        pdf.cell(40, 7, f" {str(row.get('Motorista', 'N/A'))}", border=1, align='C')
        pdf.cell(30, 7, f" {str(row.get('Data_Entrega', '---'))}", border=1, align='C')
        pdf.cell(82, 7, f" {str(row.get('Obs', 'Sem notas'))[:50]}", border=1, ln=True)

    # --- GRÁFICO (Redimensionado para caber no espaço restante) ---
    if os.path.exists(caminho_grafico):
        pdf.ln(3)
        # Largura reduzida para 140 e altura controlada
        pdf.image(caminho_grafico, x=70, w=150)
    
    # --- ASSINATURA E DATA (Posição fixa no rodapé da mesma página) ---
    pdf.set_y(-25)
    pdf.set_draw_color(0, 0, 0)
    pdf.line(110, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(1)
    pdf.set_font("Arial", 'B', 9)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 5, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    pdf.set_font("Arial", 'I', 8)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 5, f"Gerado em Luanda: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurics10@gmail.com"
    
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO LOGISTICA UNICO: {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Segue o relatorio consolidado em uma unica pagina.", 'plain'))
    
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(meu_email, senha)
            s.sendmail(meu_email, destinatario, msg.as_string())
        print("Enviado!")
    except Exception as e:
        print(f"Erro: {e}")

def executar():
    excel = "meus_locais (1).xlsx"
    if not os.path.exists(excel): return
    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        col_custo = [c for c in df.columns if 'Custo' in c][0]
        
        # Gráfico com tamanho otimizado
        plt.figure(figsize=(10, 4))
        plt.bar(df['Endereço'].str[:12], df[col_custo], color='#2E8B57') 
        plt.title('Resumo de Custos')
        plt.tight_layout()
        plt.savefig('grafico_unico.png')
        plt.close()
        
        nome_pdf = "Relatorio_Laurindo_Pagina_Unica.pdf"
        criar_pdf_premium(df, 'Endereço', col_custo, nome_pdf, 'grafico_unico.png')
        enviar_email(nome_pdf)
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar()
