import win32com.client as win32
from utils import encontrarJanelaAviso as eja
from time import sleep
import os
import json

html_file_path_carteira = 'C:\\Users\\automacao\\Documents\\RPA_python\\RPA_FaturamentoParaClientes\\src\\assets\\emailCarteira.html'
html_file_path_boleto = 'C:\\Users\\automacao\\Documents\\RPA_python\\RPA_FaturamentoParaClientes\\src\\assets\\emailBoleto.html'

def envioDoEmail(tipo, dados):
    try:
        with open("C:\\Users\\automacao\\Documents\\RPA_python\\RPA_FaturamentoParaClientes\\config\\config.json", "r", encoding="utf-8") as file:
            sensitive_data = json.load(file)
            emailsCopiaOculta = sensitive_data["enderecosEmailsCCo"]

        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        if tipo == 'carteira':
            with open(html_file_path_carteira, 'r', encoding='utf-8') as file:
                html_body = file.read()
                html_body = html_body.format(cliente=dados['cliente'], emissao=dados['emissao'], empresa=dados['empresa'], numeroNotaFiscal=dados['numeroNota'], vencimentoNotaFiscal=dados['vencimento'], valorBoleto=dados['boleto'], valorTitulo=dados['nota'])
            assunto = f"Olá {dados['cliente']}, a sua Nota Fiscal está disponível!"

        else:
            with open(html_file_path_boleto, 'r', encoding='utf-8') as file:
                html_body = file.read()
                html_body = html_body.format(cliente=dados['cliente'], emissao=dados['emissao'], empresa=dados['empresa'], numeroNotaFiscal=dados['numeroNota'], vencimentoNotaFiscal=dados['vencimento'], valorBoleto=dados['boleto'], valorTitulo=dados['nota'])
            assunto = f"Olá {dados['cliente']}, o boleto está disponível!"
            email.Attachments.Add(dados['caminhoBoleto'])

        email.To = dados['email']
        email.BCC = emailsCopiaOculta
        email.Subject = assunto
        email.HTMLBody = html_body
        email.Attachments.Add(dados['caminhoNota'])

        email.Send()

    except Exception as err:
        print(f'Erro email: {err}')
        raise Exception('Erro ao enviar email')
