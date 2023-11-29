import requests

def consultar_cnpj(cnpj):
    url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj}'

    try:
        response = requests.get(url)
        data = response.json()

        # Criando um dicionário com todas as informações desejadas
        info = {
            'uf': data.get('uf', ''),
            'municipio': data.get('municipio', ''),
            # Adicione mais campos conforme necessário
        }

        return info  # Retorne o dicionário com as informações
    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição: {e}")
        return None



# Exemplo de uso
cnpj_desejado = '22564081000195'
informacoes = consultar_cnpj(cnpj_desejado)

if informacoes:
    municipio = informacoes.get('municipio', '')
    banana = informacoes.get('uf')
    print(f'Municipio: {municipio}')
    print(f'uf: {banana}')
else:
    print('Consulta de CNPJ falhou.')

