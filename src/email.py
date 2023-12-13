import win32com.client as win32
import os

html_file_path_carteira = os.path.join(os.path.dirname(__file__), 'assets', 'emailCarteira.html')
html_file_path_boleto = os.path.join(os.path.dirname(__file__), 'assets', 'emailBoleto.html')

def envioDoEmail(tipo, dados):
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

    email.To = 'guilherme.rabelo@grupofornecedora.com.br'
    # email.To = dados['email']
    # email.BCC = "maiara.silveira@grupofornecedora.com.br;filipi.freitas@grupofornecedora.com.br;luis.castro@grupofornecedora.com.br;matheus.almeida@grupofornecedora.com.br;otavio.martins@grupofornecedora.com.br;guilherme.rabelo@grupofornecedora.com.br"
    email.Subject = assunto
    email.HTMLBody = html_body
    email.Attachments.Add(dados['caminhoNota'])

    email.Send()
