### Vídeo aula
https://www.youtube.com/watch?v=UtkPIpov6h8

### Canal do Jhonatan de Souza
https://www.youtube.com/@DevAprender

### Site para Cadastro dos produtos
https://cadastro-produtos-devaprender.netlify.app/

### Instalando dependências
###### pip install openpyxl pyautogui pyperclip mouseinfo Pillow opencv-python

### Plugins Excel para VS Code
Excel Viewer

### Usando o MouseInfo para obter a localização do mouse (use o terminal fora do VS Code)
python 
from mouseinfo import mouseInfo
mouseInfo()

---

###### *Desmarque a caixa "3 Sec. Button Delay" e pressione F6 para obter o X e Y do mouse 
###### **O pyautogui.locateOnScreen não encontra a imagem screenshot no monitor secundário, apenas no primário

---
# Mapa de ações
```
linha[0] = Nome do produto (imgs/nome_produto.png)
linha[1] = Descrição
linha[2] = Categoria
linha[3] = Código
linha[4] = Peso
linha[5] = Dimensões

Nova aba

linha[6]  = Preço (imgs/preco_produto.png)
linha[7]  = Quantidade
linha[8]  = Validade
linha[9]  = Cor
linha[10] = Tamanho (imgs/tam_pequeno.png, imgs/tam_medio.png, imgs/tam_grande.png)
linha[11] = Material

Nova aba

linha[12] = Fabricante (imgs/fabricante.png)
linha[13] = País de origem
linha[14] = Observações
linha[15] = Código de barras
linha[16] = Localização

Botão concluir (imgs/btn_concluir.png)
Botão ok (imgs/btn_ok.png)
Botão adicionar novo cadastro (imgs/btn_add.png)

-- Fim da execução
```
