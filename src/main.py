from login import loginDealernet
from boleto import downloadBoleto
from notaFiscal import downloadNotaFiscal
from sql import sqlPool
from email import envioDoEmail
from mudarEmpresa import selecionarEmpresa
import warnings
import json
import subprocess
import time
import os

warnings.filterwarnings("ignore", category=UserWarning)
with open("../config/config.json", "r", encoding="utf-8") as file:
    sensitive_data = json.load(file)
    dealernetLogin = sensitive_data["acessoDealernet"]
    senha = dealernetLogin['senha']
    enderecoBoleto = sensitive_data["enderecoBoleto"]
    enderecoNota = sensitive_data["enderecoNota"]

    dealernetModulo = sensitive_data['modulosDealernet']['ContasAReceber']
    executavel = dealernetModulo['executavel']

loginDealernet(executavel, senha)
empresas = sqlPool("SELECT", """
                    SELECT 
                        emp_cd,
                        emp_ds,
                        emp_banco
                    FROM [BD_MTZ_FOR]..ger_emp
                    WHERE 
                        emp_cd NOT IN ('20', '10', '07', '06', '05', '08', '09')
                    ORDER BY emp_ds
                """)

for empresa in empresas:
    codEmpresa = empresa[0]
    empresa = empresa[1]


    titulosTotais = sqlPool("SELECT", f"EXEC autocob.consulta_titulos '{codEmpresa}'")
    def emailsNaoEnviados(titulo):
        return titulo[15] != '1' and titulo[7] != '0024417'

    titulosPendentes = list(filter(emailsNaoEnviados, titulosTotais))
    if len(titulosPendentes):
        selecionarEmpresa(empresa)
        for i, titulo in enumerate(titulosPendentes):
            lancamento = titulo[5]
            numeroNota = titulo[8]

            dados = {
                'codCliente': titulo[2],
                'cliente': titulo[3],
                'codigoDaNota': titulo[7],
                'numeroNota': titulo[8],
                'lancamento': titulo[5],
                'serie': titulo[9],
                "parcela": titulo[6],
                'tipo': 'carteira' if titulo[4][0:2] == 'C.' else 'boleto',
                'email': titulo[14],
                'emissao': titulo[12],
                'vencimento': titulo[13],
                'codEmpresa': titulo[0],
                'empresa': titulo[1],
                'nota': titulo[11],
                'boleto': titulo[10],
                'titulo': titulo[5],
                'caminhoNota': f'{enderecoNota}NF_{numeroNota}.pdf',
                'caminhoBoleto': f'{enderecoBoleto}Boleto_{lancamento}.pdf'
            }

            try:
                if not os.path.exists(dados['caminhoNota']):
                    downloadNotaFiscal(dados['codigoDaNota'], dados['numeroNota'], dados['serie'], dados['tipo'])
                if dados['tipo']=='boleto':
                    downloadBoleto(dados['lancamento'])
                # envioDoEmail(dados['tipo'], dados)

                sqlPool("INSERT", f"""
                        DECLARE @codEmpresa VARCHAR(7) = '{dados['codEmpresa']}'
                        DECLARE @codCliente VARCHAR(7) = '{dados['codCliente']}'
                        DECLARE @titulo VARCHAR(7) = '{dados['lancamento']}'
                        DECLARE @dataOriginal DATE =  CONVERT(DATE, '{dados['emissao']}', 103)
                        DECLARE @possuiBoleto CHAR(1) = '{'0' if dados['tipo'] == 'carteira' else '1'}'
                        DECLARE @sucesso CHAR(1) = '1'
                    
                        DECLARE @dataFormatada VARCHAR(8) =  CONVERT(VARCHAR(8), @dataOriginal, 112)
                    
                        DECLARE @sqlText VARCHAR(MAX) = 
                        '
                            INSERT INTO autocob.log_execucoes
                            (empresa, cliente, titulo, dt_tituloCriacao, boleto, sucesso)
                            VALUES
                                ('''+@codEmpresa+''','''+@codCliente+''','''+@titulo+''', '''+@dataFormatada+''', '''+@possuiBoleto+''','''+@sucesso+''')
                        '
                        
                        EXEC(@sqlText)
                """)
                print(f"SUCESSO=Titulo: {dados['lancamento']}")
            except Exception as err:
                sqlPool("INSERT", f"""
                        DECLARE @codEmpresa VARCHAR(7) = '{dados['codEmpresa']}'
                        DECLARE @codCliente VARCHAR(7) = '{dados['codCliente']}'
                        DECLARE @titulo VARCHAR(7) = '{dados['lancamento']}'
                        DECLARE @dataOriginal DATE =  CONVERT(DATE, '{dados['emissao']}', 103)
                        DECLARE @possuiBoleto CHAR(1) = '{'0' if dados['tipo'] == 'carteira' else '1'}'
                        DECLARE @sucesso CHAR(1) = '0'
                    
                        DECLARE @dataFormatada VARCHAR(8) =  CONVERT(VARCHAR(8), @dataOriginal, 112)
                    
                        DECLARE @sqlText VARCHAR(MAX) = 
                        '
                            INSERT INTO autocob.log_execucoes
                            (empresa, cliente, titulo, dt_tituloCriacao, boleto, sucesso)
                            VALUES
                                ('''+@codEmpresa+''','''+@codCliente+''','''+@titulo+''', '''+@dataFormatada+''', '''+@possuiBoleto+''','''+@sucesso+''')
                        '
                    
                        EXEC(@sqlText)
                """)
                print(f"ERRO=Titulo: {dados['lancamento']}")
                subprocess.run(["powershell", "-Command", "Stop-process -Name scr"], shell=True)
                time.sleep(7)
                loginDealernet(executavel, senha)
                selecionarEmpresa(empresa)
                continue

subprocess.run(["powershell", "-Command", "Stop-process -Name scr"], shell=True)
