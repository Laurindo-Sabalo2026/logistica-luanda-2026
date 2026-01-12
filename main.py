import pandas as pd
import os
import smtplib
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime

def gerar_pdf_oficial(df):
    # Configuração da página (L = Paisagem/A4)
    pdf = FPDF('L', 'mm', 'A4')
    pdf.add_page()
    
    # --- CABEÇALHO PROFISSIONAL ---
    pdf.set_fill_color(0, 51, 102) # Azul Escuro
    pdf.rect(0, 0, 297, 35, 'F')
    pdf.set_font("Arial", 'B', 20)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 15, "LAURINDO LOGISTICA & SERVICOS", ln=True, align='C')
    pdf.set_font("Arial", 'I', 11)
    pdf.cell(0, 5, "Relatorio Oficial de Monitoramento - Luanda", ln=True, align='C')
    
    pdf.ln(20)
    
    # --- TABELA DE DADOS (MAPEAMENTO RIGIDO) ---
    # Larguras: Destino(80), Custo(40), Status(30), Motorista(50), Data(40)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.set_fill_color(220, 230, 241) # Azul claro para o topo da tabela
    
    pdf.cell(80, 10, " Destino", 1, 0, 'L', True)
    pdf.cell(40, 10, " Custo (Kz)", 1, 0, 'C', True)
    pdf.cell(30, 10, " Status", 1, 0, 'C', True)
    pdf.cell(50, 10, " Motorista", 1, 0, 'C', True)
    pdf.cell(40, 10, " Data Entr.", 1, 1, 'C', True)
    
    pdf.set_font("Arial", '', 10)
    # Loop para preencher a tabela usando a posição das colunas do Excel
    for i in range(len(df)):
        linha = df.iloc[i]
        pdf.cell(80, 10, f" {str(linha.iloc[0])[:35]}", 1)      # Coluna A
        pdf.cell(40, 10, f" {str(linha.iloc[2])}", 1, 0, 'C')   # Coluna C
        pdf.cell(30, 10, f" {str(linha.iloc[3])}", 1, 0, 'C')   # Coluna D
        pdf.cell(50, 10, f" {str(linha.iloc[4])[:20]}", 1, 0, 'C') # Coluna E
        pdf.cell(40, 10, f" {str(linha.iloc[5])}", 1, 1, 'C')   # Coluna F

    # --- RODAPÉ COM ASSINATURA ---
    pdf.ln(25)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "__________________________", ln=True, align='R')
    pdf.cell(0, 5, "Laurindo Sabalo    ", ln=True, align='R')
    
    # LINHA DA DATA (O que tínhamos planeado)
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    data_luanda = datetime.now().strftime('%d/%m/%Y %H:%M')
    pdf.cell(0, 10, f"Documento gerado em Luanda: {data_luanda}", ln=True, align='R')
    
    nome_ficheiro = "Relatorio_Logistica_LCS.pdf"
    pdf.output(nome_ficheiro)
    return nome_ficheiro

def executar_sistema():
    try:
        # 1. Leitura do Excel
        ficheiro_excel = "meus_locais (1).xlsx"
        if not os.path.exists(ficheiro_excel):
            print("Erro: Excel nao encontrado!")
            return
            
        df = pd.read_excel(ficheiro_excel)
        pdf_final = gerar_pdf_oficial(df)
        
        # 2. Configuração de E-mail
        remetente = "laurindokutala.sabalo@gmail.com"
        senha = os.environ.get('MINHA_SENHA', '').replace(" ", "")
        destinatario = "laurinds10@gmail.com"
        
        msg = MIMEMultipart()
        msg['Subject'] = f"RELATORIO CONCLUIDO - {datetime.now().strftime('%d/%m/%Y')}"
        msg['From'] = remetente
        msg['To'] = destinatario
        
        corpo = "Bom dia Laurindo,\n\nSegue em anexo o relatorio de logistica com o mapeamento de colunas corrigido."
        msg.attach(MIMEText(corpo, 'plain'))
        
        # 3. Anexo do PDF
        with open(pdf_final, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=pdf_final)
            msg.attach(anexo)
            
        # 4. Envio Real
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(remetente, senha)
            s.sendmail(remetente, destinatario, msg.as_string())
            
        print("SUCESSO: O relatorio foi enviado para laurinds10@gmail.com")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    executar_sistema()
