import pandas as pd
import os
import smtplib
import matplotlib.pyplot as plt
from fpdf import FPDF
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

def criar_pdf_final(df, col_nome, col_custo, nome_pdf, caminho_grafico):
    pdf = FPDF()
    pdf.add_page()
    
    # --- CABEÇALHO (AZUL ROYAL) ---
    pdf.set_fill_color(0, 102, 204) 
    pdf.rect(15, 15, 15, 15, 'F')  
    pdf.set_text_color(255, 255, 255) 
    pdf.set_font("Arial", 'B', 12)
    pdf.text(18, 25, "LL") 

    pdf.set_xy(35, 15)
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 102, 204)
    pdf.cell(100, 10, "LAURINDO LOGISTICA & SERVICOS", ln=True)
    
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.set_xy(35, 22)
    pdf.cell(100, 5, "Excelencia e Confianca em Luanda", ln=True)
    
    pdf.ln(8)
    pdf.set_draw_color(200, 200, 200) 
    pdf.line(15, 35, 195, 35) 
    pdf.ln(5)

    # --- TABELA DE DADOS ---
    pdf.set_font("Arial", 'B', 11)
    pdf.set_fill_color(0, 102, 204)
    pdf.set_text_color(255, 255, 255)
    
    pdf.cell(90, 10, " Destino / Localizacao", border=1, fill=True)
    pdf.cell(50, 10, "Custo (Kz)", border=1, fill=True, align='C')
    pdf.cell(40, 10, "Status", border=1, ln=True, fill=True, align='C')
    
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(0, 0, 0)
    
    for _, row in df.iterrows():
        destino = str(row[col_nome])[:45]
        try:
            custo_val = float(row[col_custo])
            custo_texto = f"{custo_val:,.2f}"
        except:
            custo_texto = str(row[col_custo])
            
        pdf.cell(90, 10, f" {destino}", border=1)
        pdf.cell(50, 10, custo_texto, border=1, align='C')
        
        # Procura coluna Status ou define como Pendente
        val_status = "Pendente"
        for c in df.columns:
            if 'Status' in str(c):
                val_status = str(row[c])
        
        pdf.cell(40, 10, f" {val_status}", border=1, ln=True, align='C')
    
    # --- GRÁFICO ---
    if os.path.exists(caminho_grafico):
        pdf.ln(10)
        pdf.image(caminho_grafico, x=35, w=130)
    
    # --- ASSINATURA E DATA (AJUSTADO PARA MAIOR ESPAÇAMENTO) ---
    pdf.set_y(-45) # Sobe a posição da linha para dar mais margem ao fundo
    pdf.set_draw_color(0, 0, 0)
    pdf.line(70, pdf.get_y(), 140, pdf.get_y()) # Linha horizontal
    
    pdf.ln(5) # Espaço entre a linha e o nome
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 5, "Laurindo Sabalo - Direccao de Logistica", ln=True, align='C')
    
    pdf.ln(4) # Espaço extra entre o nome e a data
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Gerado em Luanda - Data: {datetime.now().strftime('%d/%m/%Y')}", ln=True, align='C')
    
    pdf.output(nome_pdf)

def enviar_email(pdf_nome):
    meu_email = "laurindokutala.sabalo@gmail.com"
    # Pega a senha das variáveis de ambiente do GitHub
    senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
    destinatario = "laurinds10@gmail.com"
    
    msg = MIMEMultipart()
    msg['Subject'] = f"📊 RELATORIO LOGISTICA FINAL: {datetime.now().strftime('%d/%m/%Y')}"
    msg['From'] = meu_email
    msg['To'] = destinatario
    msg.attach(MIMEText("Segue em anexo o relatorio oficial de logistica com layout corrigido.", 'plain'))
    
    if os.path.exists(pdf_nome):
        with open(pdf_nome, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_nome)
            msg.attach(anexo)
            
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(meu_email, senha)
        s.sendmail(meu_email, destinatario, msg.as_string())

def executar():
    # Verifica qual arquivo existe
    excel = "dados.xlsx" if os.path.exists("dados.xlsx") else "meus_locais (1).xlsx"
    
    if not os.path.exists(excel):
        print(f"Erro: Arquivo {excel} nao encontrado.")
        return

    try:
        df = pd.read_excel(excel)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Localiza colunas de Destino e Custo
        col_nome = df.columns[0]
        col_custo = next((c for c in df.columns if 'Custo' in c or 'custo' in c), None)
        
        if not col_custo:
            print("Erro: Coluna de Custo nao encontrada.")
            return

        # Gera o Gráfico com as primeiras linhas
        plt.figure(figsize=(8, 4))
        df_plot = df.head(6) # Pega os 6 primeiros para o gráfico
        plt.bar(df_plot[col_nome].astype(str).str[:12], df_plot[col_custo], color='royalblue') 
        plt.title('Resumo de Custos - Logistica Luanda')
        plt.ylabel('Kz')
        plt.tight_layout()
        plt.savefig('grafico.png')
        plt.close()
        
        nome_pdf = "Relatorio_Oficial_Laurindo.pdf"
        criar_pdf_final(df, col_nome, col_custo, nome_pdf, 'grafico.png')
        enviar_email(nome_pdf)
        print("CONCLUIDO: Relatorio enviado para laurinds10@gmail.com")

    except Exception as e:
        print(f"Erro no processamento: {e}")

if __name__ == "__main__":
    executar()
