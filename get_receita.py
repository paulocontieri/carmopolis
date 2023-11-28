import requests

def consultar_informacoes(cnpj):
    def consultar_cnpj(cnpj):
        url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj}'
        
        try:
            response = requests.get(url)
            data = response.json()

            if response.status_code == 200:
                # As informações estão no dicionário 'data'
                return data
            else:
                print(f"Erro na requisição: {data['message']}")
                return None
        except Exception as e:
            print(f"Erro na requisição: {str(e)}")
            return None

    informacoes = consultar_cnpj(cnpj)

    if informacoes:
        print("Informacoes do CNPJ:")
        print(f"UF: {informacoes['uf']}")
        print(f"Municipio: {informacoes['municipio']}")
        print(f"Logradouro: {informacoes['logradouro']}")
        print(f"Bairro: {informacoes['bairro']}")
    else:
        print("Não foi possível obter informações do CNPJ.")

# Agora você pode chamar essa função em qualquer lugar do seu código, passando o CNPJ desejado
cnpj_desejado = '24237014000191'
consultar_informacoes(cnpj_desejado)
