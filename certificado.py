import openpyxl
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import portrait
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

# Definir fontes
fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

# Lista para armazenar os nomes dos arquivos de imagem gerados
nomes_arquivos_imagens = []

# Iterar sobre as linhas da planilha
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=2)):
    # Obter informações da linha
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    carga_horaria = linha[5].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    data_emissao = linha[6].value

    # Abrir a imagem do certificado
    image = Image.open('./certificado_padrao.jpg')
    draw = ImageDraw.Draw(image)

    # Adicionar informações do Excel à imagem do certificado
    draw.text((1020, 827), nome_participante, fill='black', font=fonte_nome)
    draw.text((1060, 950), nome_curso, fill='black', font=fonte_geral)
    draw.text((1435, 1065), tipo_participacao, fill='black', font=fonte_geral)
    draw.text((1480, 1182), str(carga_horaria), fill='black', font=fonte_geral)
    draw.text((750, 1770), str(data_inicio), fill='blue', font=fonte_data)
    draw.text((750, 1930), str(data_final), fill='blue', font=fonte_data)
    draw.text((2220, 1930), str(data_emissao), fill='blue', font=fonte_data)

    # Nome do arquivo de imagem a ser salvo
    nome_arquivo_imagem = f'{indice}_{nome_participante}_certificado.png'
    nomes_arquivos_imagens.append(nome_arquivo_imagem)

    # Salvar a imagem do certificado com as informações do Excel
    image.save(nome_arquivo_imagem)

# Função para criar o PDF com as imagens
def criar_pdf_certificados(imagens, pdf_filename, pdf_title):
    # Obter as dimensões da imagem do certificado
    imagem_certificado = Image.open(imagens[0])  # Use a primeira imagem como referência
    largura_imagem, altura_imagem = imagem_certificado.size

    # Criando PDF de acordo com tamanho de um certificado
    c = canvas.Canvas(pdf_filename, pagesize=(largura_imagem, altura_imagem))

    c.setTitle(pdf_title)

    for imagem_path in imagens:
        try:
            imagem = ImageReader(imagem_path)
            c.drawImage(imagem, 0, 0, width=largura_imagem, height=altura_imagem)
        except Exception as e:
            print(f"Erro ao adicionar a imagem ao PDF: {e}")

        c.showPage()

    c.save()

# Nome do arquivo PDF a ser gerado
pdf_filename = 'certificados.pdf'

#PDF Titulo
pdf_title = 'Certificado de Conclusão'

# Cria o PDF com as imagens
criar_pdf_certificados(nomes_arquivos_imagens, pdf_filename, pdf_title)

print("PDF de certificados gerado com sucesso!")
