import openpyxl
import pyperclip
import pyautogui
from time import sleep
import os

################################################

# Entra na planilha
workbook = openpyxl.load_workbook('planilha.xlsx')
sheet_produtos = workbook['Produtos']
aguardar = 2

################################################

# Copiar informações de um campo e colar em outro no site
for linha in sheet_produtos.iter_rows(min_row=2):

    # Copia dados do produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)

    # Localiza o campo através do printscreen
    file_nome_produto = 'nome_produto.png'
    if os.path.isfile(file_nome_produto):
        scrNomeProduto = pyautogui.locateOnScreen(file_nome_produto, confidence=0.8)
        if scrNomeProduto is not None:
            pyautogui.click(scrNomeProduto, duration=0.5)
            pyautogui.hotkey('ctrl', 'v')
        else:
            print('Imagem não encontrada na tela #nome.')
            quit()
    else: 
        print('Imagem não localizada no diretório #nome.')
        quit()

    # Copia dados da descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')

    # Copia dados da categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados do código
    codigo = linha[3].value
    pyperclip.copy(codigo)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados do peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados das dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')

    ############################################

    # Pressiona o botão próximo no formulário
    file_proximo = 'btn_proximo.png'
    if os.path.isfile(file_proximo):
        scrProximo = pyautogui.locateOnScreen(file_proximo, confidence=0.8)
        if scrProximo is not None:
            pyautogui.click(scrProximo, duration=0.5)
        else:
            print('Imagem não encontrada "próximo".')
            quit()
    else: 
        print('Imagem não localizada no diretório #prox.')
        quit()

    sleep(aguardar)
    
    ############################################
    
    # Copia dados do preço
    preco = linha[6].value
    pyperclip.copy(preco)
    # Localiza o campo através do printscreen
    file_preco = 'preco_produto.png'
    if os.path.isfile(file_nome_produto):
        scrPreco = pyautogui.locateOnScreen(file_preco, confidence=0.8)
        if scrPreco is not None:
            pyautogui.click(scrPreco, duration=0.5)
            pyautogui.hotkey('ctrl', 'v')
        else:
            print('Imagem não encontrada na tela #prec.')
            quit()
    else: 
        print('Imagem não localizada no diretório #prec.')
        quit()
    
    # Copia dados da quantidade
    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados da validade
    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados da cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Abre a lista de tamanhos
    tamanho = linha[10].value
    pyautogui.hotkey('tab')
    pyautogui.hotkey('alt', 'down')
    # Localiza o campo através do printscreen
    file_peq = 'tam_pequeno.png'
    file_med = 'tam_medio.png'
    file_grd = 'tam_grande.png'
    if os.path.isfile(file_peq) | os.path.isfile(file_med) | os.path.isfile(file_grd):
        scrPreco = pyautogui.locateOnScreen(file_preco, confidence=0.8)
        # Seleciona o tamanho na lista
        if tamanho == 'Pequeno':
            if file_peq is not None:
                pyautogui.click(file_peq, duration=0.5)
            else:
                print('Imagem não encontrada na tela #peq.')
                quit()
        elif tamanho == 'Médio':
            if file_med is not None:
                pyautogui.click(file_med, duration=0.5)
            else:
                print('Imagem não encontrada na tela #med.')
                quit()
        else:
            if file_grd is not None:
                pyautogui.click(file_grd, duration=0.5)
            else:
                print('Imagem não encontrada na tela #grd.')
                quit()
    else: 
        print('Imagem não localizada no diretório #tam.')
        quit()
        
    # Copia dados do material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')

    ############################################

    # Pressiona o botão próximo no formulário
    file_proximo = 'btn_proximo.png'
    if os.path.isfile(file_proximo):
        scrProximo = pyautogui.locateOnScreen(file_proximo, confidence=0.8)
        if scrProximo is not None:
            pyautogui.click(scrProximo, duration=0.5)
        else:
            print('Imagem não encontrada "próximo".')
            quit()
    else: 
        print('Imagem não localizada no diretório #prox.')
        quit()

    sleep(aguardar)
    
    ############################################

    # Copia dados do fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    # Localiza o campo através do printscreen
    file_fabricante = 'fabricante.png'
    if os.path.isfile(file_fabricante):
        scrFab = pyautogui.locateOnScreen(file_fabricante, confidence=0.8)
        if scrFab is not None:
            pyautogui.click(scrFab, duration=0.5)
            pyautogui.hotkey('ctrl', 'v')
        else:
            print('Imagem não encontrada na tela #fab.')
            quit()
    else: 
        print('Imagem não localizada no diretório #fab.')
        quit()
    
    # Copia dados do país de origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados das observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados do código de barras
    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')
    
    # Copia dados da localização
    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl', 'v')

    ############################################

    # Pressiona o botão concluir no formulário
    file_concluir = 'btn_concluir.png'
    if os.path.isfile(file_concluir):
        scrConcluir = pyautogui.locateOnScreen(file_concluir, confidence=0.8)
        if scrConcluir is not None:
            pyautogui.click(scrConcluir, duration=0.5)
        else:
            print('Imagem não encontrada "concluir".')
            quit()
    else: 
        print('Imagem não localizada no diretório #conc.')
        quit()

    # Pressiona o botão confirmar inclusão
    file_ok = 'btn_ok.png'
    if os.path.isfile(file_ok):
        scrOk = pyautogui.locateOnScreen(file_ok, confidence=0.8)
        if scrOk is not None:
            pyautogui.click(scrOk, duration=0.5)
            sleep(1)
            pyautogui.click(scrOk, duration=0.5)
        else:
            print('Imagem não encontrada "ok".')
            quit()
    else: 
        print('Imagem não localizada no diretório #ok.')
        quit()

    sleep(aguardar)

    # Pressiona o botão de incluir novo cadastro
    file_add = 'btn_add.png'
    if os.path.isfile(file_add):
        scrAdd = pyautogui.locateOnScreen(file_add, confidence=0.8)
        if scrAdd is not None:
            pyautogui.click(scrAdd, duration=0.5)
        else:
            print('Imagem não encontrada "add".')
            quit()
    else: 
        print('Imagem não localizada no diretório #add.')
        quit()
    
    sleep(aguardar)