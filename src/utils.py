import pyautogui as pag
from pywinauto.application import Application
import pygetwindow as gw
from time import sleep
import os

# C:\Users\automacao\Documents\RPA_python\RPA_FaturamentoParaClientes\src\assets\botao.png

#base_diretorio = os.getcwd()
#caminho_imagem = os.path.join(base_diretorio, 'src\\assets\\botao.png')

def clicarEmImagem(imagem, i):
    ocorrencias = list(pag.locateAllOnScreen(imagem, confidence=0.9))
    posicao = ocorrencias[i]
    left, top, width, height = posicao
    centro_x = left + width / 2
    centro_y = top + height / 2
    pag.click(centro_x, centro_y)


def encontrarJanelaAviso():
    tela_aviso_outlook = Application(backend="win32").connect(title="Microsoft Outlook", timeout=10, found_index=0)
    tela_aviso_outlook.MicrosoftOutlook.set_focus()
    resposta = ''
    if tela_aviso_outlook.MicrosoftOutlook.set_focus():
        print("conexão com a janela ok")
        sleep(3)
        clicarEmImagem("C:\\Users\\automacao\\Documents\\RPA_python\\RPA_FaturamentoParaClientes\\src\\assets\\botao.png", 0)
        resposta = 'botão clicado com sucesso'
        return resposta
    else:
        resposta = 'falha ao clicar no botão'
        return resposta


