from login import loginDealernet
from boleto import downloadBoleto
import json

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
tratarErro(downloadBoleto, 'inprimir boleto', nomeModulo, '0177772')
