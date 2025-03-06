from pywinauto.application import Application
import pyautogui
import time
import subprocess


def loginDealernet(modulo, senha):
    try:        
        comando_powershell = f"start {modulo}"
        processoAtivo = subprocess.run(["powershell", "-Command", f"Get-Process -Name scr"], shell=True, capture_output=True, text=True)

        if processoAtivo.returncode == 0:
            subprocess.run(["powershell", "-Command", "Stop-process -Name scr -f"], shell=True)
        time.sleep(5) 

        subprocess.run(["powershell", "-Command", comando_powershell], capture_output=False, text=True)
        try: 
            atencao = Application(backend="win32").connect(title=f'Atenção', timeout=10)
            atencao.Atencao.wait('visible', timeout=2)
            pyautogui.press('ENTER')
            time.sleep(5)
        except:
            time.sleep(5)

        telaLogin = Application(backend="win32").connect(title='Segurança', timeout=30)

        time.sleep(2)
        telaLogin.Seguranca.set_focus()

        telaLogin.Seguranca.child_window(class_name="Edit").wrapper_object().click_input()
        pyautogui.write(senha)
        time.sleep(2)
        telaLogin.Seguranca.child_window(title="&OK", class_name="Button").wrapper_object().click_input()
        time.sleep(5)
        
    except Exception as err:
        raise Exception(f'Erro ao realizar login: {err}')