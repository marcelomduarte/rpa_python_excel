import pandas as pd
import os
import datetime

data = datetime.datetime.now()

# criando um DataFrame vazio com a estrutura final do Consolidado
colunas = [
    'Segmento',
    'País',
    'Produto',
    'Qtde de Unidades Vendidas',
    'Preço Unitário',
    'Valor Total',
    'Desconto',
    'Valor Total c/ Desconto',
    'Custo Total',
    'Lucro',
    'Data',
    'Mês',
    'Ano'
]
consolidado = pd.DataFrame(columns=colunas)

# busca o nome dos arquivos a serem consolidados
arquivos = os.listdir("datasets")

# realiza a consolidação dos arquivos (apenas .xlsx)
for excel in arquivos:

    if excel.endswith('.xlsx'):
        dados_arquivo = excel.split('-')
        segmento = dados_arquivo[0]
        pais = dados_arquivo[1].replace('.xlsx', '')
        
        try:
            df = pd.read_excel(f'datasets\\{excel}')
            df.insert(0, 'Segmento', segmento)
            df.insert(1, 'País', pais)        
            consolidado = pd.concat([consolidado, df])  # Consolidar as planilhas usando o método concat    
        except:
            with open('log_erros.txt', 'a') as arquivo:
                arquivo.write(f'Erro ao tentar consolidar o arquivo {excel}.\n')
    else:
        with open('log_erros.txt', 'a') as arquivo:
            arquivo.write(f'O arquivo {excel} não é um arquivo Excel válido!\n')
            

# formata a coluna data
consolidado['Data'] = consolidado['Data'].dt.strftime("%d/%m/%Y")

# exporta o DataFrame consolidado para um arquivo Excel
consolidado.to_excel(f"Report-consolidado-{data.strftime('%d-%m-%Y')}.xlsx", 
                     index=False,
                     sheet_name='Report consolidado')