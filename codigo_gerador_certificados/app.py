import openpyxl
from PIL import Image, ImageDraw, ImageFont
import os

#criar pasta para os certicicados
pasta = "certificados_gerados/"
if not os.path.exists(pasta):
    os.mkdir(pasta)

#Abrir a planilha
workbook = openpyxl.load_workbook('planilha_alunos.xlsx')
student_sheet = workbook['Sheet1']

for indice, linha in enumerate(student_sheet.iter_rows(min_row=2)):
    # extrair dados da plan
    nome_curso = linha[0].value #Nome do curso
    nome_participante = linha[1].value #Nome do partícipante
    tipo_participante = linha[2].value # Tipo de partícipante
    data_inicio = linha[3].value # Data de Início
    data_termino = linha[4].value #Data do Término
    carga_horaria = linha[5].value #Carga Horária
    data_emissao = linha[6].value #Data de Emissão do certificado

    #transferir dados da planilha para a imagem
    #Definindo a fonte a ser usada
    fonte_nome = ImageFont.truetype('./tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 50)
    #abrindo imagem
    img = Image.open('./certificado_padrao.jpg')

    #escrevendo na imagem
    desenhar = ImageDraw.Draw(img)

    desenhar.text((1020,827),nome_participante,fill='black',font=fonte_nome)
    desenhar.text((1060,950),nome_curso,fill='black',font=fonte_geral)
    desenhar.text((1435,1065),tipo_participante,fill = 'black', font = fonte_geral)
    desenhar.text((1480,1182),str(carga_horaria), fill = 'black', font = fonte_geral)

    desenhar.text((700,1770), data_inicio, fill='black', font=fonte_geral)
    desenhar.text((700,1922), data_termino, fill='black', font=fonte_geral)

    desenhar.text((2180,1922), data_emissao, fill='black', font=fonte_geral)

    img.save(f'./certificados_gerados/{indice} {nome_participante} certificado.png')


