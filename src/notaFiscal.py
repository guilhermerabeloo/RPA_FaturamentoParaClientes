from pywinauto.application import Application
import pyautogui
import time

def downloadNotaFiscal(codNf, numeroNf, serieNf, tipo):
    app = Application(backend="win32").connect(class_name="FNWND3115")
    main_window = app.top_window()
    main_window.set_focus()
    time.sleep(1)

    # Entrando no módulo de emissao da nota
    pyautogui.hotkey('alt', 'u')
    time.sleep(.5)
    for i in range(3):
        pyautogui.press('up')
        time.sleep(.01)

    time.sleep(1)
    pyautogui.press('ENTER')
    time.sleep(3)
    pyautogui.write(codNf)
    time.sleep(.5)
    pyautogui.press('TAB')
    pyautogui.write(serieNf)
    time.sleep(.5)
    pyautogui.press('TAB')
    time.sleep(2)

    quantidadeTab = 9 if serieNf == 'S' else 10
    # Selecionando a opção de imprimir
    for i in range(quantidadeTab):
        pyautogui.press('TAB')
        time.sleep(.03)

    pyautogui.press('SPACE')
    time.sleep(.5)
    if tipo == 'boleto':
        pyautogui.hotkey('ALT', 'N')
        time.sleep(1)
    pyautogui.hotkey('ALT', 'S')
    time.sleep(4)

    if serieNf == 'S':
        for i in range(6):
            pyautogui.press('down')
            time.sleep(.01)

        pyautogui.press('ENTER')
        time.sleep(5)
        pyautogui.hotkey('ALT', 'N')
    else:
        for i in range(2):
            pyautogui.press('TAB')
            time.sleep(.01)
        pyautogui.press('down')
        time.sleep(.01)
        pyautogui.press('ENTER')

    # Salvando boleto
    time.sleep(10)
    explorerNota = Application(backend="win32").connect(title=f'Salvar Saída de Impressão como')
    main_window = explorerNota.top_window()
    main_window.set_focus()
    pyautogui.write(f'NF_{numeroNf}')
    time.sleep(1)

    pastaRaiz = False
    while pastaRaiz==False: # Produrando a pasta Desktop
        explorerNota.SalvarSaidaDeImpressaoComo.child_window(title="Barra de ferramentas da faixa superior", class_name="ToolbarWindow32").wrapper_object().click_input()
        pastaRaiz = explorerNota.SalvarSaidaDeImpressaoComo.child_window(title="Endereço: Área de Trabalho", class_name="ToolbarWindow32").exists()

    explorerNota.SalvarSaidaDeImpressaoComo.child_window(title="Endereço: Área de Trabalho", class_name="ToolbarWindow32").wrapper_object().click_input()
    pyautogui.press('ENTER') # ativar esta linha caso o click na barra de endereco do documento nao habilite para colar o caminho do  não tiver como
    pyautogui.write(f'C:\\Users\\automacao\\Documents\\RPA_docs\\Autocob\\NotasFiscais')
    pyautogui.press('ENTER')
    pyautogui.hotkey('alt', 'l')

    time.sleep(2)
    try: 
        substituirArquivo = Application(backend="win32").connect(title=f'Confirmar Salvar como')
        substituirArquivo.ConfirmarSalvarComo.wait('visible', timeout=2)
        pyautogui.hotkey('alt', 's')
        time.sleep(10)

    except:
        time.sleep(10)

    # Saindo da tela de nota fiscal de saída
    pyautogui.press('ESC')
    time.sleep(.5)
    app.ContasAReceberEmpresaMatrizUsuarioAutomacao.child_window(title="&S", class_name="Button").wrapper_object().click_input()

time.sleep(5)