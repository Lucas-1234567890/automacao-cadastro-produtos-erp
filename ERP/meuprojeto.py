import pyautogui
import subprocess
from time import sleep, time
import os
import pandas as pd
import pyperclip

print(os.getcwd())

pyautogui.FAILSAFE = True

subprocess.Popen([r"C:\Program Files\Fakturama2\Fakturama.exe"])

def encontrar_imagem(imagem):
    timeout = 20
    inicio = time()

    encontrou = None

    while True:
        try:
            encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.8)
            if encontrou:
                break
        except Exception:
            pass

        if time() - inicio > timeout:
            print(f'Tempo limite atingido. Imagem não encontrada: {imagem}')
            break

        sleep(1)

    return encontrou

def direita(posicoes_imagem):
    return posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3]/2

def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')

pyautogui.alert('o programa vai começar, não mexa no computador')

# Ler tabela

tabela_produtos = pd.read_excel('Produtos.xlsx')
tabela_produtos["Status"] = ""

# Aguardar o programa abrir
encontrar_imagem('logo.png')

for linha in tabela_produtos.index:
    try:
        nome = tabela_produtos.loc[linha, "Nome"]
        id = tabela_produtos.loc[linha, "ID"]
        categoria = tabela_produtos.loc[linha, "Categoria"]
        gtin = tabela_produtos.loc[linha, "GTIN"]
        supplier = tabela_produtos.loc[linha, "Supplier"]
        descricao = tabela_produtos.loc[linha, "Descrição"]
        imagem = tabela_produtos.loc[linha, "Imagem"]
        preco = tabela_produtos.loc[linha, "Preço"]
        custo = tabela_produtos.loc[linha, "Custo"]
        estoque = tabela_produtos.loc[linha, "Estoque"]

        encontrar_imagem('new.png')
        pyautogui.click(pyautogui.center(encontrar_imagem('new.png')))

        encontrar_imagem('new_project.png')
        pyautogui.click(pyautogui.center(encontrar_imagem('new_project.png')))

        pyautogui.click(direita(encontrar_imagem('item_number.png')))
        escrever_texto(str(id))

        pyautogui.click(direita(encontrar_imagem('Name.png')))
        escrever_texto(str(nome))

        pyautogui.click(direita(encontrar_imagem('Category.png')))
        escrever_texto(str(categoria))

        pyautogui.click(direita(encontrar_imagem('GTIN.png')))
        escrever_texto(str(gtin))

        pyautogui.click(direita(encontrar_imagem('supplier code.png')))
        escrever_texto(str(supplier))

        pyautogui.click(direita(encontrar_imagem('descripition.png')))
        escrever_texto(str(descricao))

        pyautogui.click(direita(encontrar_imagem('price.png')))
        preco_texto = f'{preco:.2f}'.replace('.',',')
        escrever_texto(preco_texto)

        pyautogui.click(direita(encontrar_imagem('cost.png')))
        custo_texto = f'{custo:.2f}'.replace('.',',')
        escrever_texto(custo_texto)

        pyautogui.click(direita(encontrar_imagem('stock.png')))
        estoque_texto = f'{estoque:.2f}'.replace('.',',')
        escrever_texto(estoque_texto)

        pyautogui.click(pyautogui.center(encontrar_imagem('selecionar_imagem.png')))
        encontrar_imagem('nome_arquivo.png')
        escrever_texto(fr'C:\Users\Lucas\OneDrive\Arquivos diversos\Desktop\ERP\Imagens Produtos\{str(imagem)}')
        pyautogui.press('enter')

        pyautogui.click(direita(encontrar_imagem('save.png')))

        tabela_produtos.loc[linha, "Status"] = "OK"

    except Exception as erro:
        print(f'Erro na linha {linha}: {erro}')
        tabela_produtos.loc[linha, "Status"] = "Erro"

tabela_produtos.to_excel("Produtos_Automacao.xlsx", index=False)
print('Finalizou')

pyautogui.hotkey('alt', 'f4')

pyautogui.alert('pode mexer agora')