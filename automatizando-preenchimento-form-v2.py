import openpyxl
import pyperclip
import pyautogui
from time import sleep
import os

################################################

# Define a planilha com dados dos produtos
workbook = openpyxl.load_workbook('planilha.xlsx')
sheet_produtos = workbook['Produtos']

################################################

# Função para pressionar botão de ação
def pressionar_botao_acao(fileImg, aguardar = 2):

    # Verifica se o arquivo existe no diretório
    if os.path.isfile(fileImg):
        
        # Localiza o botão através do printscreen
        scrBtn = pyautogui.locateOnScreen(fileImg, confidence=0.8)
        
        # Clica no botão
        if scrBtn is not None:
            pyautogui.click(scrBtn, duration=0.5)
        else:
            print('Imagem não encontrada "btn ação".')
            quit()

    else: 
        print('Imagem não localizada no diretório #btn.')
        quit()

    sleep(aguardar)

################################################

# Copiar informações de um campo e colar em outro no site
for linha in sheet_produtos.iter_rows(min_row=2):

    for item in linha:

        if item.column == 1: # Nome do produto
            pressionar_botao_acao('imgs/nome_produto.png', 1)            
        elif item.column == 7: # Preço
            pressionar_botao_acao('imgs/preco_produto.png', 1)
        elif item.column == 11: # Tamanhos

            # Expande a lista de tamanhos
            pyautogui.hotkey('tab')
            pyautogui.hotkey('alt', 'down')

            # Seleciona o tamanho do produto na lista
            if item.value == 'Pequeno':
                pressionar_botao_acao('imgs/tam_pequeno.png', 0.5)
            elif item.value == 'Médio':
                pressionar_botao_acao('imgs/tam_medio.png', 0.5)
            elif item.value == 'Grande':
                pressionar_botao_acao('imgs/tam_grande.png', 0.5)

        elif item.column == 13: # Fabricante
            pressionar_botao_acao('imgs/fabricante.png', 1)
        else:
            pyautogui.hotkey('tab')

        ########################################

        # Avança pra o próximo campo e cola o valor
        pyperclip.copy(item.value)
        pyautogui.hotkey('ctrl', 'v')

        ########################################

        # Botões de ação
        if item.column == 6 or item.column == 12:
            # Avança abas
            pressionar_botao_acao('imgs/btn_proximo.png')
        elif item.column == 17:
            # Conclui cadastro
            pressionar_botao_acao('imgs/btn_concluir.png')
            pressionar_botao_acao('imgs/btn_ok.png')
            pressionar_botao_acao('imgs/btn_add.png')

############################################
############################################
# Mapa de ações

# linha[0] = Nome do produto (imgs/nome_produto.png)
# linha[1] = Descrição
# linha[2] = Categoria
# linha[3] = Código
# linha[4] = Peso
# linha[5] = Dimensões

# Nova aba

# linha[6]  = Preço (imgs/preco_produto.png)
# linha[7]  = Quantidade
# linha[8]  = Validade
# linha[9]  = Cor
# linha[10] = Tamanho (imgs/tam_pequeno.png, imgs/tam_medio.png, imgs/tam_grande.png)
# linha[11] = Material

# Nova aba

# linha[12] = Fabricante (imgs/fabricante.png)
# linha[13] = País de origem
# linha[14] = Observações
# linha[15] = Código de barras
# linha[16] = Localização

# Botão concluir (imgs/btn_concluir.png)
# Botão ok (imgs/btn_ok.png)
# Botão adicionar novo cadastro (imgs/btn_add.png)

############################################