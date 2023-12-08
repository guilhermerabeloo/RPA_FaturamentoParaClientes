import pyodbc
import os
from configparser import ConfigParser

def sqlPool(operacao, script):
    def readConfig():
        script_dir = os.path.dirname(__file__)
        config_path = os.path.join(script_dir, '..', 'config', 'config.ini')

        config = ConfigParser()
        config.read(config_path)
        return config['DatabaseConfig']

    dadosConfig = readConfig()

    dados_conexao = (
        f"Driver={dadosConfig['Driver']};"
        f"Server={dadosConfig['Server']};"
        f"Database={dadosConfig['Database']};"
        f"UID={dadosConfig['UID']};"
        f"PWD={dadosConfig['PWD']};"
    )

    try: 
        conexao = pyodbc.connect(dados_conexao)
        # print("Conexão com banco de dados criada com sucesso!")
        try:
            cursor = conexao.cursor()
            cursor.execute(script)
            if operacao == 'SELECT':
                result = cursor.fetchall()
            
                return result
            else: 
                conexao.commit()

                return None
        except Exception as err:
            print(f"Erro ao executar a stored procedure: {err}")

    except Exception as err:
        print(f'Erro ao realizar conexão com o banco de dados: {err}')
    
    finally:
        conexao.close()


# sqlPool("INSERT", f"""
#         DECLARE @codEmpresa VARCHAR(7) = '{'01'}'
#         DECLARE @codCliente VARCHAR(7) = '{'01'}'
#         DECLARE @titulo VARCHAR(7) = '{'01'}'
#         DECLARE @dataOriginal DATE =  CONVERT(DATE, '{'12/12/2023'}', 103)
#         DECLARE @possuiBoleto CHAR(1) = '{'0'}'
#         DECLARE @sucesso CHAR(1) = '1'
        
#         DECLARE @dataFormatada VARCHAR(8) =  CONVERT(VARCHAR(8), @dataOriginal, 112)
        
#         DECLARE @sqlText VARCHAR(MAX) = 
#         '
#             INSERT INTO autocob.log_execucoes_teste
#             (empresa, cliente, titulo, dt_tituloCriacao, boleto, sucesso)
#             VALUES
#                 ('''+@codEmpresa+''','''+@codCliente+''','''+@titulo+''', '''+@dataFormatada+''', '''+@possuiBoleto+''','''+@sucesso+''')
#         '
        
#         EXEC(@sqlText)
# """)