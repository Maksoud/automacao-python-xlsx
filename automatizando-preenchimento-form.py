import openpyxl
import pyperclip
import pyautogui
import pyscreeze
from time import sleep

# Entra na planilha
workbook = openpyxl.load_workbook('planilha.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informações de um campo e colar em outro no site
for linha in sheet_produtos.iter_rows(min_row=2):

    # Copia dados do produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    # Localiza o campo através do printscreen
    scrNomeProduto = pyautogui.locateOnScreen('nome_produto.png')
    pyautogui.click(scrNomeProduto, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Copia dados da descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('ctrl', 'v')

    # Copia dados da categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1160, 562, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados do código
    codigo = linha[3].value
    pyperclip.copy(codigo)
    pyautogui.click(1160, 646, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados do peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1160, 734, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados das dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1160, 821, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    ############################################

    # Pressiona o botão próximo no formulário
    pyautogui.click(1179, 883, duration=0.5)
    sleep(2)
    
    ############################################
    
    # Copia dados do preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1160, 370, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados da quantidade
    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(1160, 450, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados da validade
    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(1160, 540, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados da cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1160, 625, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Abre a lista de tamanhos
    tamanho = linha[10].value
    pyautogui.click(1160, 705, duration=0.5)

    # Seleciona o tamanho na lista
    if tamanho == 'Pequeno':
        pyautogui.click(1175, 745, duration=0.5)
    elif tamanho == 'Médio':
        pyautogui.click(1175, 765, duration=0.5)
    else:
        pyautogui.click(1175, 785, duration=0.5)
    
    # Copia dados do material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1160, 800, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    ############################################

    # Pressiona o botão próximo no formulário
    pyautogui.click(1180, 855, duration=0.5)
    sleep(2)
    
    ############################################

    # Copia dados do fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1160, 400, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados do país de origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1160, 490, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados das observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1160, 575, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados do código de barras
    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.click(1160, 710, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados da localização
    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(1160, 795, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    ############################################

    # Pressiona o botão concluir no formulário
    pyautogui.click(1180, 855, duration=0.5)

    # Pressiona o botão confirmar inclusão
    pyautogui.click(1660, 170, duration=0.5)

    sleep(2)

    # Pressiona o botão de incluir novo cadastro
    pyautogui.click(1485, 620, duration=0.5)
    
    sleep(2)