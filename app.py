import openpyxl
from PIL import Image, ImageDraw, ImageFont


# ABRINDO A PLANILHA
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada celula que conte, a infor
    linha_curso = linha[0].value #nome do curso
    nome_participante = linha[1].value #nome do participante
    tipo_participacao = linha[2].value #tipo do participante
    data_inicio = linha[3].value #data de inicio
    data_final = linha[4].value #data de conclusao
    carga_horaria = linha[5].value #carga horaria
    data_emissao = linha[6].value #data de emissao

#TRANFESIR DADOS PARA O CERTIFICADO
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((1020,827), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1435,1065), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((1060,950), linha_curso, fill='black', font=fonte_geral)
    desenhar.text((1480,1182), str(carga_horaria), fill='black', font=fonte_geral)
    
    desenhar.text((775,1770), data_inicio, fill='blue', font=fonte_data)
    desenhar.text((750, 1930), data_final, fill='blue', font=fonte_data)
    desenhar.text((2220, 1930), data_emissao, fill='blue', font=fonte_data)


    imagem.save(f'./{indice}{nome_participante} certificado.png')