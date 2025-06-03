import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime
import os

# Caminhos
PASTA_SAIDA = "notas_geradas"
LOGO_CAMINHO = "recursos/logo_borbas.png"
os.makedirs(PASTA_SAIDA, exist_ok=True)

def gerar_pdf_nota_com_imposto_e_frete(nota_info, produtos, nome_arquivo):
    c = canvas.Canvas(nome_arquivo, pagesize=A4)
    largura, altura = A4

    # Cabeçalho moderno
    c.setFillColorRGB(0.2, 0.2, 0.2)  # cinza escuro
    c.rect(0, altura - 80, largura, 80, stroke=0, fill=1)
    if os.path.exists(LOGO_CAMINHO):
        logo = ImageReader(LOGO_CAMINHO)
        c.drawImage(logo, 20, altura - 60, width=90, height=40, mask='auto')
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(120, altura - 40, "EMPRESA BORBAS LTDA")
    c.setFont("Helvetica", 10)
    c.drawString(120, altura - 55, "Rua Exemplo, 100 - Bairro Exemplo | CNPJ: 00.000.000/0001-00")

    # Cliente e dados da nota
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.drawString(20, altura - 100, f"Destinatário: {nota_info['Cliente']}")
    c.drawString(20, altura - 115, f"Endereço: {nota_info['Endereço']}")
    c.drawString(400, altura - 100, f"Nº: {nota_info['Nota']}")
    c.drawString(400, altura - 115, f"Série: {nota_info['Série']}")

    # Cabeçalho da tabela
    y = altura - 150
    c.setFont("Helvetica-Bold", 9)
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(20, y, largura - 40, 20, fill=1, stroke=0)
    c.setFillColor(colors.black)
    colunas = ["Produto", "Qtd", "Unitário", "Imposto (%)", "Total s/ Imp.", "Total c/ Imp."]
    pos_x = [25, 180, 230, 300, 380, 470]
    for i, col in enumerate(colunas):
        c.drawString(pos_x[i], y + 5, col)

    # Conteúdo da tabela
    y -= 20
    total_sem_imposto = 0
    total_com_imposto = 0
    c.setFont("Helvetica", 9)
    for _, prod in produtos.iterrows():
        valor_total = prod['Quantidade'] * prod['Valor Unitário']
        valor_imposto = valor_total * (prod['Imposto (%)'] / 100)
        total_final = valor_total + valor_imposto

        total_sem_imposto += valor_total
        total_com_imposto += total_final

        dados = [
            prod['Produto'][:25],
            str(prod['Quantidade']),
            f"{prod['Valor Unitário']:.2f}",
            f"{prod['Imposto (%)']:.1f}",
            f"{valor_total:.2f}",
            f"{total_final:.2f}"
        ]
        for i, dado in enumerate(dados):
            c.drawString(pos_x[i], y + 5, dado)
        c.setStrokeColor(colors.lightgrey)
        c.line(20, y, largura - 20, y)
        y -= 18

    # Frete
    frete = produtos['Frete'].iloc[0]
    total_com_imposto += frete

    # Totais
    y -= 10
    c.setFillColorRGB(0.96, 0.96, 0.96)
    c.rect(300, y - 60, 270, 60, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(310, y - 15, f"Subtotal s/ Impostos: R$ {total_sem_imposto:.2f}")
    c.drawString(310, y - 30, f"Total com Impostos: R$ {total_com_imposto - frete:.2f}")
    c.drawString(310, y - 45, f"Frete: R$ {frete:.2f}")
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(colors.darkblue)
    c.drawString(310, y - 60, f"Total Final: R$ {total_com_imposto:.2f}")

    # Rodapé
    c.setFillColorRGB(0.9, 0.9, 0.9)
    c.rect(0, 0, largura, 30, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Oblique", 8)
    c.drawCentredString(largura / 2, 10, f"Emitido em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - Borbas Software")

    c.save()

def gerar_notas_fiscais_completas(arquivo_excel):
    df = pd.read_excel(arquivo_excel)

    for (nota, serie), grupo in df.groupby(['Nota', 'Série']):
        dados_nota = grupo.iloc[0]
        nome_arquivo = os.path.join(PASTA_SAIDA, f"nota_{nota}_serie_{serie}.pdf")
        gerar_pdf_nota_com_imposto_e_frete(dados_nota, grupo, nome_arquivo)
        print(f"Nota {nota} (série {serie}) gerada: {nome_arquivo}")

if __name__ == "__main__":
    try:
        gerar_notas_fiscais_completas("modelo_notas_fiscais.xlsx")
    except Exception as e:
        print("Erro ao gerar nota:", e)
