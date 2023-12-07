from login import loginDealernet
from boleto import downloadBoleto
from notaFiscal import downloadNotaFiscal
from sql import sqlPool
from email import envioDoEmail
from mudarEmpresa import selecionarEmpresa
import warnings
import json

warnings.filterwarnings("ignore", category=UserWarning)
with open("../config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    dealernetLogin = sensitive_data["acessoDealernet"]
    senha = dealernetLogin['senha']

    dealernetModulo = sensitive_data['modulosDealernet']['ContasAReceber']
    executavel = dealernetModulo['executavel']
    nomeModulo = dealernetModulo['nome']

def tratarErro(funcao, mensagem, param1, param2):
    try:
        funcao(param1, param2)
    except Exception as err:
        print(f'Erro ao "{mensagem}"')

tratarErro(loginDealernet, 'efetuar login', executavel, senha)
empresas = sqlPool("""
                    SELECT 
                        emp_cd,
                        emp_ds,
                        emp_banco
                    FROM [BD_MTZ_FOR]..ger_emp
                    WHERE 
                        emp_cd IN ('01')
                        --emp_cd NOT IN ('20', '10', '07', '06', '05', '08', '09')
                    ORDER BY emp_ds
                """)

for empresa in empresas:
    codEmpresa = empresa[0]
    empresa = empresa[1]

    selecionarEmpresa(empresa)

    titulosTotais = sqlPool(f"EXEC autocob.consulta_titulos '{codEmpresa}'")
    def emailsNaoEnviados(titulo):
        return titulo[14] == '1'

    titulosPendentes = list(filter(emailsNaoEnviados, titulosTotais))
    for i, titulo in enumerate(titulosPendentes):
        numeroNota = titulo[7]
        lancamento = titulo[5]
        serie = titulo[8]
        tipo = 'carteira' if titulo[4][0:2] == 'C.' else 'boleto'
        dados = {
            'cliente': titulo[3],
            'email': 'guilherme.rabelo@grupofornecedora.com.br',
            'emissao': titulo[11],
            'vencimento': titulo[12],
            'empresa': titulo[1],
            'nota': titulo[10],
            'boleto': titulo[9],
            'titulo': titulo[5],
            'caminhoNota': f'C:\\Users\\guilherme.rabelo\\Desktop\\TesteRPA\\NotasFiscais\\NF_{numeroNota}.pdf',
            'caminhoBoleto': f'C:\\Users\\guilherme.rabelo\\Desktop\\TesteRPA\\Boletos\\Boleto_{lancamento}.pdf'
        }

        if titulo[6] == '001':
            downloadNotaFiscal(numeroNota, serie, tipo)
        if tipo=='boleto':
            downloadBoleto(nomeModulo, lancamento)

        envioDoEmail(tipo, dados)

