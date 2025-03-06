from pywinauto.application import Application
import pyautogui
import time

def downloadBoleto(codEmpresa, lancamento, numeroNf,  serieNf):
    try:
        app = Application(backend="win32").connect(class_name="FNWND3115")
        main_window = app.top_window()
        main_window.set_focus()
        time.sleep(1)

        # Entrando no módulo de emissao do boleto
        pyautogui.hotkey('alt', 'u')
        time.sleep(.5)
        for i in range(15):
            pyautogui.press('down')
            time.sleep(.01)

        time.sleep(1)
        pyautogui.press('ENTER')
        time.sleep(7)

        # Emitindo boleto
        app.ContasAReceberEmpresaUsuarioAutomacao.child_window(title="Título", class_name="Button").wrapper_object().click_input()
        time.sleep(2)
        pyautogui.press('TAB')
        pyautogui.write(lancamento)
        pyautogui.press('TAB')
        time.sleep(2)
        app.ContasAReceberEmpresaUsuarioAutomacao.child_window(title="Confi&g.", class_name="Button").wrapper_object().click_input()
        pyautogui.press('ENTER')
        time.sleep(2)
        app.ContasAReceberEmpresaUsuarioAutomacao.child_window(title="&Imprimir", class_name="Button").wrapper_object().click_input()
        # time.sleep(6)
        atencao = Application(backend="win32").connect(title=f'Atenção', timeout=30)
        time.sleep(1)
        atencao.Atencao.child_window(title="&OK", class_name="Button").wrapper_object().click_input()

        # Salvando boleto
        # time.sleep(7)
        explorerBoleto = Application(backend="win32").connect(title=f'Salvar Saída de Impressão como', timeout=60)
        main_window = explorerBoleto.top_window()
        main_window.set_focus()
        time.sleep(10)
        pyautogui.write(f'C:\\Users\\automacao\\Documents\\RPA_docs\\EmailsClientes\\Boletos\\Boleto_{numeroNf}{serieNf}_{lancamento}_{codEmpresa}')
        time.sleep(2)

        # pastaRaiz = False
        # while pastaRaiz==False: # Produrando a pasta Desktop
        #     explorerBoleto.SalvarSaidaDeImpressaoComo.child_window(title="Barra de ferramentas da faixa superior", class_name="ToolbarWindow32").wrapper_object().click_input()
        #     pastaRaiz = explorerBoleto.SalvarSaidaDeImpressaoComo.child_window(title="Endereço: Área de Trabalho", class_name="ToolbarWindow32").exists()

        # explorerBoleto.SalvarSaidaDeImpressaoComo.child_window(title="Endereço: Área de Trabalho", class_name="ToolbarWindow32").wrapper_object().click_input()
        # pyautogui.press('ENTER') # ativar esta linha caso o click na barra de endereco do documento nao habilite para colar o caminho do  não tiver como
        # pyautogui.write(f'C:\\Users\\automacao\\Documents\\RPA_docs\\EmailsClientes\\Boletos\\')
        # pyautogui.press('ENTER')
        pyautogui.hotkey('alt', 'l')
        time.sleep(2)

        try: 
            substituirArquivo = Application(backend="win32").connect(title=f'Confirmar Salvar como')
            substituirArquivo.ConfirmarSalvarComo.wait('visible', timeout=2)
            pyautogui.hotkey('alt', 's')
            time.sleep(15)

        except:
            time.sleep(15)
        
        main_window = app.top_window()
        main_window.set_focus()
        time.sleep(3)

        app.ContasAReceberEmpresaUsuarioAutomacao.child_window(title="&Cancelar", class_name="Button").wrapper_object().click_input()
        time.sleep(5)
    except Exception as err:
        print(err)