                                                  ## IMPORTS

import tkinter as tk
import sqlite3
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import TimeoutException
from tkinter import filedialog
from PIL import Image, ImageTk
import customtkinter
from tkinter import ttk
import time
import re
import datetime
import threading
import sys
import tkinter.simpledialog
import requests


                                                ## CONFIGURAÇÕES

# VARIÁVEIS GLOBAIS -------------------------------------------------------------------
navegador = None
empresa = ""
log_conn = None
selected_period = ""


# LOGINS ------------------------------------------------------------------------------

#######################################################
# PRIMADO
primado_login = "30041020000171"
primado_senha = "Acero#2698"

# SERVIÇOS
servicos_login = "40044113000103"
servicos_senha = "Servicos#0103"

# ACERO ESTUFA
estufa_login = "05108821000593"
estufa_senha = "Acero#2698"

# ACERO MTZ
acero_login = "05108821000160"
acero_senha = "Acero#2697"


                                      ## CLASSE PARA ESPELHAR O PRINT NO LOG
# Class to redirect print statements to the textbox
class TextboxRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.insert(tk.END, text)
        self.textbox.see(tk.END)  # Scroll to the end of the textbox

    def flush(self):
        pass

                                    ## CÓDIGO DE AUTOMAÇÃO E BANCO DE DADOS

## CONFIGURAÇÃO DO BANCO DE DADOS PRINCIPAL -----------------------------------------------------------------------

#######################################################
# Função para lidar com o upload do arquivo Excel
def upload_arquivo():
    try:
    # Verifica se há dados no banco de dados antes de permitir o upload
        if verificar_dados_no_banco():
            tkinter.messagebox.showerror("Erro", f"Ainda existem dados no banco. Não é possível fazer o upload.")
            return

        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if file_path:
            carregar_dados(file_path)
            tkinter.messagebox.showinfo("Sucesso", f"Arquivos carregados!")
    except Exception as e:
        tkinter.messagebox.showerror("Erro", f"Arquivo Excel Inválido!")

#######################################################
# Função para verificar se há dados no banco de dados
def verificar_dados_no_banco():
    conn = sqlite3.connect('banco.db')
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) FROM banco')
    quantidade_dados = cursor.fetchone()[0]
    conn.close()

    return quantidade_dados > 0

#######################################################
# Carregar dados
def carregar_dados(arquivo):
    conn = sqlite3.connect('banco.db')
    cursor = conn.cursor()

    df = pd.read_excel(arquivo)

    for index, row in df.iterrows():
        # Validações
        if str(row['Descrição Utilização']) == "Servico Tomado (NF de Serviços-Prefeitura)" or str(row['Descrição Utilização']) == "Servico Prestado (NF de Serviços-Prefeitura)":
            cnpj = row['CNPJ']
            cnpj_entidade = row['CNPJ Entidade']
            cpf = row['CPF']

            if not pd.isna(cnpj):
                cnpj_str = '{:014d}'.format(int(cnpj))
            else:
                cnpj_str = 0

            if not pd.isna(cnpj_entidade):
                cnpj_entidade_str = '{:014d}'.format(int(cnpj_entidade))
            else:
                cnpj_entidade_str = None

            if not pd.isna(cpf):
                cpf_str = '{:011d}'.format(int(cpf))
            else:
                cpf_str = None

            # Utilize expressões regulares para extrair o número entre parênteses e remover o ponto
            descricao_item = row['Descrição do item']
            numero_entre_parenteses = re.search(r'\((\d+\.\d+)\)', descricao_item)
            
            if numero_entre_parenteses:
                item_formatado = numero_entre_parenteses.group(1).replace('.', '')
            else:
                item_formatado = None
            
            cursor.execute('''
                INSERT INTO banco (serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, "valor_contabil", cnpj_entidade, valor_iss)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                row['Serial'],
                row['Descrição Utilização'],
                cnpj_str,
                cpf_str,
                row['Data do doc.'],
                row['IE Entidade'],
                item_formatado,
                row['Valor contábil'],
                cnpj_entidade_str,
                row['Valor ISS']
            ))

    conn.commit()

    # Agrupar linhas com o mesmo valor em "Serial" e somar "Valor contábil"
    cursor.execute('''
        UPDATE banco
        SET valor_contabil = (
            SELECT SUM(valor_contabil)
            FROM banco AS m
            WHERE m.serial = banco.serial
        )
    ''')

    # Remover linhas duplicadas
    cursor.execute('''
        DELETE FROM banco
        WHERE rowid NOT IN (
            SELECT MIN(rowid)
            FROM banco
            GROUP BY serial
        )
    ''')

    conn.commit()
    conn.close()

#######################################################
# Configuração do banco de dados principal
def configurar_banco_dados():
    conn = sqlite3.connect('banco.db')
    cursor = conn.cursor()

    # Criação da tabela banco se não existir
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS banco (
            serial TEXT,
            descricao_utilizacao TEXT,
            cnpj TEXT,
            cpf TEXT,
            data_doc TEXT,
            ie_entidade TEXT,
            descricao_item TEXT,
            valor_contabil REAL,
            cnpj_entidade TEXT,
            valor_iss REAL
        )
    ''')

    conn.commit()
    conn.close()

configurar_banco_dados()


## CONFIGURAÇÃO DE BANCO DE DADOS (EXCLUIR, EDITAR ETC) -------------------------------------------------------------

#######################################################
# Função para excluir a primeira linha do banco de dados
def excluir_primeira_linha(status_lancamento):
    # Conectar ao banco de dados principal
    conn_principal = sqlite3.connect('banco.db')
    cursor_principal = conn_principal.cursor()

    # Conectar ou criar o banco de dados de logs
    conn_logs = sqlite3.connect('logs.db')
    cursor_logs = conn_logs.cursor()

    # Selecionar o ID da primeira linha
    cursor_principal.execute('SELECT MIN(rowid), * FROM banco')
    min_rowid, *linha_excluida = cursor_principal.fetchone()

    # Excluir a primeira linha do banco principal
    cursor_principal.execute('DELETE FROM banco WHERE rowid = ?', (min_rowid,))
    conn_principal.commit()

    # Inserir a linha excluída no banco de logs
    cursor_logs.execute('CREATE TABLE IF NOT EXISTS logs (serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, valor_contabil, cnpj_entidade, valor_iss, status)')
    cursor_logs.execute('INSERT INTO logs VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', (*linha_excluida, status_lancamento))
    conn_logs.commit()

    # Fechar as conexões
    conn_principal.close()
    conn_logs.close()

#######################################################
# Função para excluir todas as linhas do banco de dados
def excluir_todas_as_linhas():
    # Conectar ao banco de dados principal
    conn_principal = sqlite3.connect('banco.db')
    cursor_principal = conn_principal.cursor()

    # Excluir todas as linhas da tabela principal
    cursor_principal.execute('DELETE FROM banco')
    conn_principal.commit()

    tkinter.messagebox.showinfo("Sucesso", f"O Banco de Dados foi Excluído com Sucesso!")

    # Fechar a conexão do banco principal
    conn_principal.close()



#######################################################
# Função para excluir todas as linhas do banco de dados
def excluir_todas_as_linhas_logs():
    conn = sqlite3.connect('logs.db')
    cursor = conn.cursor()

    # Excluir todas as linhas da tabela
    cursor.execute('DELETE FROM logs')
    conn.commit()
    conn.close()
    tkinter.messagebox.showinfo("Sucesso", f"O histórico de Lançamentos foi excluído com sucesso!")

## CONFIGURAÇÃO DE BANCO DE DADOS (EXCLUIR, EDITAR ETC) -----------------------------------------------------------

#######################################################
# Função para exportar dados para Excel
def exportar_banco():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

    if file_path:
        conn = sqlite3.connect('banco.db')
        query = 'SELECT * FROM banco'
        df = pd.read_sql_query(query, conn)
        conn.close()

        df.to_excel(file_path, index=False)

        tkinter.messagebox.showinfo("Sucesso", f"Dados exportados com sucesso!")

#######################################################
# Função para exportar dados para Excel
def exportar_logs():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

    if file_path:
        conn = sqlite3.connect('logs.db')
        query = 'SELECT * FROM logs'
        df = pd.read_sql_query(query, conn)
        conn.close()

        df.to_excel(file_path, index=False)

        tkinter.messagebox.showinfo("Sucesso", f"Dados exportados com sucesso!")

## INFORMAÇÕES DA EMPRESA (GET RECEITA FEDERAL) ------------------------------------------------------------

#######################################################
# Função consultar cnpj
def consultar_cnpj(cnpj):
    url = f'https://www.receitaws.com.br/v1/cnpj/{cnpj}'

    try:
        response = requests.get(url)
        data = response.json()

        # Criando um dicionário com todas as informações desejadas
        info = {
            'tipo': data.get('tipo', ''),
            'nome': data.get('nome', ''),
            'uf': data.get('uf', ''),
            'municipio': data.get('municipio', ''),
            'logradouro': data.get('logradouro', ''),
            'numero': data.get('numero', ''),
            'cep': data.get('cep', ''),
            'email': data.get('email', ''),
            'bairro': data.get('bairro'),
            'cep': data.get('cep'),
            'email': data.get('email', ''),
            # Adicione mais campos conforme necessário
        }

        return info  # Retorne o dicionário com as informações
    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição: {e}")
        return None



## CONFIGURAÇÃO DE NAVEGADOR -------------------------------------------------------------------------------

#######################################################
# Função para abrir o navegador apenas uma vez
def abrir_navegador():
    print("Iniciando Sistema...")
    print("")
    print("♦♦*•.¸¸¸.•*¨¨*•.¸¸¸.•*•♦¸.•*¨¨*•.¸¸¸.•*•.¸.•*¨¨*•.¸¸¸.•*")
    print("   ♦♦   ░A░G░R░O░C░O░N░T░A░R░")
    print("*•♦*•.¸¸¸.•*¨¨*•.¸¸¸.•*•♦¸.•*¨¨*•.¸¸¸.•*•.¸.•*¨¨*•.¸¸¸.•*")
    print("             Soluções Contábeis para o Agronegócio")
    print("")
    print("")

    global navegador
    if not navegador:
        options = Options()
        # options.add_argument("--headless")  # Habilitar o modo headless
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico, options=options)
        navegador.maximize_window()
        navegador.get("https://nfse-carmopolisdeminas.atende.net/autoatendimento/servicos/nfse/") 

        # Store the ID of the original window
        original_window = navegador.current_window_handle
        assert len(navegador.window_handles) == 1

        # EFETUAR LOGIN
        # Aguardar até que a página esteja completamente carregada
        WebDriverWait(navegador, 10).until(
            lambda x: x.execute_script("return document.readyState == 'complete'")
        )


        if empresa == "SERVIÇOS":
            # Preencher campos de login
            campo_login = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[3]/input"))
            )
            campo_login.send_keys(servicos_login)

            # Preencher campos de senha
            campo_senha = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[5]/div/input"))
            )
            campo_senha.send_keys(servicos_senha)


        if empresa == "PRIMADO":
            # Preencher campos de login
            campo_login = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[3]/input"))
            )
            campo_login.send_keys(primado_login)

            # Preencher campos de senha
            campo_senha = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[5]/div/input"))
            )
            campo_senha.send_keys(primado_senha)


        if empresa == "ACERO ESTUFA":
            # Preencher campos de login
            campo_login = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[3]/input"))
            )
            campo_login.send_keys(estufa_login)

            # Preencher campos de senha
            campo_senha = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[5]/div/input"))
            )
            campo_senha.send_keys(estufa_senha)


        if empresa == "ACERO MTZ":
            # Preencher campos de login
            campo_login = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[3]/input"))
            )
            campo_login.send_keys(acero_login)

            # Preencher campos de senha
            campo_senha = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[5]/div/input"))
            )
            campo_senha.send_keys(acero_senha)


        # Aguardar até que a página esteja completamente carregada
        WebDriverWait(navegador, 10).until(
            lambda x: x.execute_script("return document.readyState == 'complete'")
        )

        # Clicar em entrar
        botao_entrar = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[6]/button"))
        )
        botao_entrar.click()

        janela_principal = navegador.current_window_handle
        #print("Janela principal: " + janela_principal)

        # Clicar em acessar
        elemento_acessar = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.LINK_TEXT, "Acessar"))
        )
        elemento_acessar.click()

        time.sleep(2) 

        all_handles = navegador.window_handles

        time.sleep(2)

        print(f"Login feito com sucesso na empresa {empresa}")
        print(f"Período selecionado: {selected_period}")

        # Itera sobre os handles das janelas e muda para a janela desejada
        for handle in all_handles:
            if handle != janela_principal:
                navegador.switch_to.window(handle)
                #print("Trabalhando na Nova Janela:", handle)
                break  # Sai do loop após encontrar a janela desejada

        all_nadles = navegador.window_handles

        time.sleep(2)

        # Clicar na Lateral (Fiscal)
        elemento_fiscal = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//*[@id='estrutura_menu_conjuntos']/ul/li[1]/div"))
        )
        elemento_fiscal.click()

        time.sleep(1)

        # Clicar em Escrita Fiscal
        escrita_fiscal = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//*[@id='estrutura_menu_conjuntos']/ul/li[1]/ul/li[1]/div"))
        )
        escrita_fiscal.click()

        time.sleep(2)




## AUTOMAÇÃO DENTRO DA PREFEITURA --------------------------------------------------------------------------
#######################################################
def processo():
    try:
        conn = sqlite3.connect('banco.db')
        cursor = conn.cursor()
        cursor.execute('SELECT serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, valor_contabil, cnpj_entidade, valor_iss FROM banco')
        linhas = cursor.fetchall()
        conn.close()

        quantidade_restante = len(linhas)
        quantidade_atual = 1
        quantidade_falha = 0
        quantidade_sucesso = 0
        
        print("Quantidade de lançamentos: " + str(quantidade_restante))
        print("")
        print("")
        time.sleep(5)
        print("")
        print("")
        print("--------------------------------------------------------------------------------------------- L A N Ç A M E N T O S -----------------------------------------------------------------------------")
        print("")


        for linha in linhas:
            serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, valor_contabil, cpj_entidade, valor_iss = linha

            # Formatar valor com 2 casas decimais
            valor_contabil_formatado = "{:.2f}".format(valor_contabil)

            # Verificar datas 
            data_atualizada = data_doc[-7:]

            # Verificar CNPJ
            informacoes = consultar_cnpj(cnpj)

            if informacoes is not None:
                uf = informacoes.get('uf', '')
                municipio = informacoes.get('municipio', '')
                tipo = informacoes.get('tipo', '')
                nome = informacoes.get('nome', '')
            else:
                # Lida com a situação em que não há informações para o CNPJ
                print(f"Nenhuma informação encontrada para o CNPJ: {cnpj}")


            # Verificar Situação Tributária
            situacao_tributaria_print = ""

            # Clicar no menu
            botao_menu_tomador = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//*[@id='estrutura_menu_sistema']/ul/li[5]"))
            )
            botao_menu_tomador.click()

            # Clicar tomador
            botao_tomador_tomador = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, ".estrutura_submenu_sistema:nth-child(6) > .estrutura_menu_item:nth-child(2) > .estrutura_submenu_item_titulo"))
            )
            botao_tomador_tomador.click()
            time.sleep(2)


            ## Se for SERVIÇO TOMADO
            if descricao_utilizacao == "Servico Tomado (NF de Serviços-Prefeitura)" and selected_period == data_atualizada:

                # Aguardar até que a página esteja completamente carregada
                WebDriverWait(navegador, 10).until(
                    lambda x: x.execute_script("return document.readyState == 'complete'")
                )

                # Clicar Consultar
                botao_consultar_tomador = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64022_101_1']/article/div[1]/aside[1]/div/div[3]/table/tbody/tr/td[3]"))
                )
                botao_consultar_tomador.click()
                time.sleep(1)

                # Clicar Ordernar Competência
                botao_ordenar_tomador = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64022_101_1']/article/div[1]/header/div[2]/table/tbody/tr[1]/td[3]"))
                )
                botao_ordenar_tomador.click()
                time.sleep(1)

                try:
                    # Clicar na Competência escolhida no app
                    linha_competencia_tomador = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, f"//tr[td[@aria-description='{selected_period}']]"))
                    )
                    linha_competencia_tomador.click() 
                    time.sleep(1)
                except:
                    print("Competência não encontrada!")

                # Clicar em declaração
                botao_declaracao_tomador = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, ".fa-calendar-week > .label_botao_acao"))
                )
                botao_declaracao_tomador.click() 
                time.sleep(1)

                # Clicar em declarar
                botao_declarar_tomador = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_101_1']/article/div[1]/aside[2]/div[1]/span"))
                )
                botao_declarar_tomador.click() 
                time.sleep(2)

                # Selecionar tipo 5
                selecionar_tipo = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[1]/table/tbody/tr[2]/td/fieldset/div/div/table/tbody/tr[1]/td[2]/span/select/option[4]"))
                )
                selecionar_tipo.click() 
                time.sleep(1)

                # Preencher serie
                campo_serie = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[1]/table/tbody/tr[2]/td/fieldset/div/div/table/tbody/tr[2]/td[2]/span/input"))
                )
                campo_serie.send_keys("E")
                time.sleep(1)

                # Selecionar situacao
                selecionar_situacao = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[1]/table/tbody/tr[2]/td/fieldset/div/div/table/tbody/tr[2]/td[4]/span/select/option[1]"))
                )
                selecionar_situacao.click() 
                time.sleep(1)

                # Preencher numero/serial
                campo_serial = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[1]/table/tbody/tr[2]/td/fieldset/div/div/table/tbody/tr[3]/td[2]/span/input"))
                )
                campo_serial.send_keys(serial)
                time.sleep(1)

                # Preencher dia
                campo_dia = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[1]/table/tbody/tr[2]/td/fieldset/div/div/table/tbody/tr[4]/td[2]/span/input"))
                )
                campo_dia.send_keys(data_doc[:2])
                time.sleep(1)


                if cnpj != '0':
                
                    cnpj_option = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[1]/table/tbody/tr[3]/td/fieldset/div/div/table/tbody/tr[1]/td[2]/span/select/option[2]"))
                    )
                    cnpj_option.click() 
                    time.sleep(1)

                    campo_cnpj = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[1]/table/tbody/tr[3]/td/fieldset/div/div/table/tbody/tr[2]/td[2]/span/input[2]"))
                    )
                    campo_cnpj.send_keys(cnpj)
                    time.sleep(2)
                    campo_cnpj.send_keys(Keys.ENTER)
                    time.sleep(2)


                try:
                    cadastro_empresa = WebDriverWait(navegador, 1).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='estrutura_container_sistema']/div[4]/section/footer/button[1]"))
                    )
                    tkinter.messagebox.showerror("Erro", f"O CNPJ do Prestador({cnpj}) não está Cadastrado no Sistema! Feche o Aplicativo, Faça o Cadastro e volte a Executar.")
                    cadastro_empresa.click()

                except TimeoutException:
                    pass

                #Clica em proximo
                proximo_button = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='janela_64023_139_1']/div[2]/footer/button[1]"))
                )
                proximo_button.click() 
                time.sleep(2)

                #Local da prestação
                if valor_iss == 0:
                    local_prestacao = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[2]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[1]/table/tbody/tr[2]/td[2]/span/input[4]"))
                    )
                    local_prestacao.send_keys("CARMÓPOLIS DE MINAS")
                    time.sleep(2)
                    local_prestacao.send_keys(Keys.ENTER)
                    time.sleep(1)
                elif valor_iss != 0:
                    local_prestacao = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[2]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[1]/table/tbody/tr[2]/td[2]/span/input[4]"))
                    )
                    local_prestacao.send_keys(municipio)
                    time.sleep(2)
                    local_prestacao.send_keys(Keys.ENTER)
                    time.sleep(1)
                
                #Lista Serviço
                lista_servico = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[2]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[1]/table/tbody/tr[3]/td[2]/span/input[1]"))
                )
                lista_servico.send_keys(descricao_item)
                time.sleep(1)
                
                # Mesmo municipio que Prestador
                if municipio == "CARMOPOLIS DE MINAS":
                    situacao_tributaria = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[2]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[1]/table/tbody/tr[5]/td[2]/span/select/option[13]"))
                    )
                    situacao_tributaria.click() 
                    situacao_tributaria_print = "NTPEM"
                    time.sleep(1)

                # Municipio diferente mas sem valor no campo iss
                elif municipio != "CARMOPOLIS DE MINAS" and valor_iss == 0:
                    situacao_tributaria = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[2]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[1]/table/tbody/tr[5]/td[2]/span/select/option[14]"))
                    )
                    situacao_tributaria.click() 
                    situacao_tributaria_print = "NTREP"
                    time.sleep(1)

                # Municipio diferente e valor no campo iss
                elif municipio != "CARMOPOLIS DE MINAS" and valor_iss != 0:
                    situacao_tributaria = WebDriverWait(navegador, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[2]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[1]/table/tbody/tr[5]/td[2]/span/select/option[2]"))
                    )
                    situacao_tributaria.click() 
                    situacao_tributaria_print = "TIRF"
                    time.sleep(1)

                # Valor do Serviço:
                valor_servico = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64023_139_1']/div/section/div[2]/article[2]/table/tbody/tr/td/div/div/table/tbody/tr[2]/td[1]/table/tbody/tr[7]/td[2]/span/input"))
                )
                valor_servico.click()
                time.sleep(1)
                valor_servico.send_keys(valor_contabil_formatado)
                time.sleep(1)


                fechar1 = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='janela_64023_139_1']/div[2]/footer/button[4]"))
                )
                fechar1.click()
                time.sleep(2)
                
                fechar2 = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='estrutura_janelas_abertas']/div/div[3]/ul/ul/li[2]/span/span[2]"))
                )
                fechar2.click()
                time.sleep(2)
                
                
                print("")
                print(f"NOTA Nº {quantidade_atual} LANÇADA!")
                print("")
                print("Dados do Prestador: ")
                if cnpj != '0':
                    print(f"CNPJ: {cnpj}")
                    print(f"Nome: {nome}")
                    print(f"Tipo: {tipo}")
                    print(f"UF: {uf}")
                    print(f"Município: {municipio}")
                else:
                    print("O Prestador é um CPF")
                print("")
                print("Dados da Nota:")
                print(f"Serial: {serial}")
                print(f"Data: {data_doc}")
                print(f"CNPJ: {cnpj}")
                print(f"Descrição Utilização: {descricao_utilizacao}")
                if valor_iss != 0:
                    print(f"Local da Prestação: {municipio} - {uf}")
                else:
                    print("Local da Prestação: Carmópolis de Minas - MG")
                print(f"Descrição Item: {descricao_item}")
                print(f"Situação Tributária: {situacao_tributaria_print}")
                print(f"Valor ISS: R$ {valor_iss} reais")
                print(f"Valor Contábil: R$ {valor_contabil_formatado} reais")
                print("")
                print("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")

                quantidade_atual = quantidade_atual + 1
                quantidade_sucesso = quantidade_sucesso + 1
                status_lancamento = "LANÇADO"
                excluir_primeira_linha(status_lancamento)

                time.sleep(5)


            else:
                primeira_condicao = ""
                segunda_condicao = ""
                terceira_condicao = ""

                if descricao_utilizacao == "Servico Prestado (NF de Serviços-Prefeitura)":
                    primeira_condicao = "* O sistema está configurado para efetuar exclusivamente lançamentos de Notas de Serviço Tomado."

                if data_atualizada != selected_period:
                    segunda_condicao = "* O sistema não efetua lançamentos de Notas com datas que não coincidem com o período escolhido, em comparação com a data presente no arquivo importado."

                if cnpj == '0':
                    terceira_condicao = "* O Prestador é uma pessoa física (CPF)."
      
                print("")
                print(f"NOTA Nº {quantidade_atual} NÃO LANÇADA!")
                print("")
                print("Dados do Prestador: ")
                if cnpj != '0':
                    print(f"CNPJ: {cnpj}")
                    print(f"Nome: {nome}")
                    print(f"Tipo: {tipo}")
                    print(f"UF: {uf}")
                    print(f"Município: {municipio}")
                else:
                    print("O Prestador é um CPF")
                print("")
                print("Dados da Nota:")
                print(f"Serial: {serial}")
                print(f"Data: {data_doc}")
                print(f"CNPJ: {cnpj}")
                print(f"Descrição Utilização: {descricao_utilizacao}")
                print(f"Descrição Item: {descricao_item}")
                print(f"Situação Tributária: {situacao_tributaria_print}")
                print(f"Valor ISS: R$ {valor_iss} reais")
                print(f"Valor Contábil: R$ {valor_contabil_formatado} reais")
                print("")
                print(f"Motivo da falha: \n\n{primeira_condicao}\n{segunda_condicao}\n{terceira_condicao}")
                print("")
                print("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
                quantidade_falha = quantidade_falha + 1
                quantidade_atual = quantidade_atual + 1
                status_lancamento = "NÃO LANÇADO"
                excluir_primeira_linha(status_lancamento)

                time.sleep(45)

        print("")
        print("")
        print("╔═════════╗")
        print("┃ ▁▂▃▅▆▇ 100% |")
        print("╚═════════╝")
        print(f"LANÇADAS: {quantidade_sucesso}")
        print(f"NÃO LANÇADAS: {quantidade_falha}")



    except Exception as e:
        tkinter.messagebox.showerror("Erro", f"Ocorreu um erro com o sistema! Reinicie a aplicação.")
        print("Feche a Aplicação, mas não se esqueça de Salvar o Log caso seja necessário :)")
        print(e)



## INICIAR PROCESSO --------------------------------------------------------------------------

#######################################################
# Função para iniciar processo
def iniciar_processo():

    conn = sqlite3.connect('banco.db')
    cursor = conn.cursor()
    cursor.execute('SELECT serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, valor_contabil, cnpj_entidade, valor_iss FROM banco')
    cnpj_entidade = cursor.fetchone()
    conn.close()


    cnpj_escolhido = ""

    if empresa == "ACERO MTZ":
        cnpj_escolhido = "05108821000160"
    elif empresa == "ACERO ESTUFA":
        cnpj_escolhido = "05108821000593"
    elif empresa == "PRIMADO":
        cnpj_escolhido = "30041020000171"
    elif empresa == "SERVIÇOS":
        cnpj_escolhido = "40044113000103"

    if empresa == "":
        tkinter.messagebox.showerror("Erro", f"Escolha uma Empresa para Iniciar o Processo.")
    elif selected_period == "":
        tkinter.messagebox.showerror("Erro", f"Escolha um Período para Iniciar o Processo.") 
    elif not verificar_dados_no_banco():
        tkinter.messagebox.showerror("Erro", f"Não há dados no banco de dados. Carregue um arquivo primeiro.")
    else:
        # Verifica se há dados no banco de dados antes de iniciar o processo
        if cnpj_escolhido == cnpj_entidade[8]:
                # Inicie a função abrir_navegador_thread em uma thread separada
                # Inicie o navegador e o processo na mesma thread
                processo_navegador_thread = threading.Thread(target=lambda: [abrir_navegador(), processo()])
                processo_navegador_thread.start()
        else:
            tkinter.messagebox.showerror("Erro", f"A empresa escolhida não confere com a empresa no banco de dados. \n Empresa Escolhida: {cnpj_escolhido} \n Empresa no Banco de Dados: {cnpj_entidade[8]}")

            
            




                                        ## INTERFACE DO APP

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        # Set the window icon
        self.iconbitmap("logo.ico")

        # configure window
        self.title("Automação Prefeitura Carmópolis de Minas")
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create top menu
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        # create system menu
        self.system_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.system_menu.add_command(label="Excluír Banco de Dados", command=excluir_todas_as_linhas)
        self.system_menu.add_command(label="Excluir Histórico de Lançamentos", command=excluir_todas_as_linhas_logs)
        self.menu_bar.add_cascade(label="Configurações", menu=self.system_menu)

        # create export menu
        self.export_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.export_menu.add_command(label="Exportar Banco de Dados", command=exportar_banco)
        self.export_menu.add_command(label="Exportar Lançamentos Feitos", command=exportar_logs)
        self.export_menu.add_command(label="Exportar Logs", command=self.exportar_logs)
        self.menu_bar.add_cascade(label="Exportar", menu=self.export_menu)  # Fix here, use self.export_menu instead of self.system_menu

        # create help menu
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label="Período", command=self.clear_textbox_periodo)
        self.help_menu.add_command(label="Upload de Arquivo", command=self.clear_textbox_upload)
        self.help_menu.add_command(label="Iniciar Processo", command=self.clear_textbox_iniciar_processo)
        self.help_menu.add_command(label="Limpar Logs", command=self.clear_textbox_limpar_logs)
        self.help_menu.add_command(label="Exportar Logs", command=self.clear_textbox_exportar_logs)
        self.help_menu.add_command(label="Exportar Banco de Dados", command=self.clear_textbox_exportar_banco)
        self.help_menu.add_command(label="Exportar Histórico de Lançamentos", command=self.clear_textbox_exportar_historico)
        self.help_menu.add_command(label="Excluir Banco de Dados", command=self.clear_textbox_excluir_banco)
        self.help_menu.add_command(label="Excluir Histórico de Lançamentos", command=self.clear_textbox_excluir_lancamentos)

        self.menu_bar.add_cascade(label="Ajuda", menu=self.help_menu)


        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Ações", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))


        # Período ##########################################################

        self.empresa_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Empresa", "ACERO MTZ", "ACERO ESTUFA", "PRIMADO", "SERVIÇOS" ], command=self.update_selected_empresa)
        self.empresa_option_menu.set("Empresa")
        self.empresa_option_menu.grid(row=1, column=0, padx=20, pady=10)

        # Calculate current, previous, and next months
        today = datetime.datetime.now()
        current_month_year = today.strftime("%m/%Y")
        previous_month_year = (today - datetime.timedelta(days=30)).strftime("%m/%Y")
        next_month_year = (today + datetime.timedelta(days=30)).strftime("%m/%Y")

        # Create the CTkOptionMenu with the command attribute
        self.period_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Período", previous_month_year, current_month_year, next_month_year], command=self.update_selected_period)
        self.period_option_menu.set("Período")
        self.period_option_menu.grid(row=2, column=0, padx=20, pady=10)

        #####################################################################

        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Upload do Arquivo", command=upload_arquivo)
        self.sidebar_button_1.grid(row=3, column=0, padx=20, pady=10)


        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Tema:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))

        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="Escala da Interface:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.set("100%")
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create textbox in column 1
        self.textbox = customtkinter.CTkTextbox(self, wrap="word", height=330)
        self.textbox.grid(row=0, column=1, padx=(20, 20), pady=(20, 0), sticky="nsew", columnspan=2, rowspan="2")

        self.clear_button = customtkinter.CTkButton(self, text="Limpar Logs", command=self.clear_textbox, fg_color=("#D8544E"), hover_color=("#B23E39"), text_color="white")
        self.clear_button.grid(row=2, column=1, padx=(20, 5), pady=(0, 30), sticky="e")  # Reduce the horizontal padding

        self.start_button = customtkinter.CTkButton(self, text="Iniciar Processo", command=lambda: (iniciar_processo(), self.clear_textbox()), fg_color=("#338633"), hover_color=("#307430"))
        self.start_button.grid(row=2, column=2, padx=(5, 20), pady=(0, 30), sticky="w")

        # Add this line to get the current year
        current_year = datetime.datetime.now().year

        self.footer_label = customtkinter.CTkLabel(self, text=f"Agrocontar © {current_year}", anchor="w")
        self.footer_label.grid(row=3, column=1, padx=20, pady=10)

        # Redirect print statements to the textbox
        sys.stdout = TextboxRedirector(self.textbox)

    def export_to_txt(self, filename):
        try:
            with open(filename, 'w', encoding='utf-8') as file:
                content = self.textbox.get("1.0", "end-1c")
                file.write(content)
            print("")
            print("Sucesso:", f"O log foi salvo em: {filename}")
        except Exception as e:
            print("")
            print("Erro ao exportar arquivo", f"Um erro ocorreu: {e}")

    def exportar_logs(self):
        # Prompt the user to choose a file for exporting
        filename = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if filename:
            self.export_to_txt(filename)

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(text="Type in a number:", title="CTkInputDialog")
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)
        
    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def config_option(self):
        print("Config option selected")

    def test_option(self):
        print("Test option selected")

    def clear_textbox(self):
        self.textbox.delete(1.0, tk.END)

    def update_selected_period(self, selected_value):
        global selected_period
        selected_period = selected_value

    def update_selected_empresa(self, selected_value):
        global empresa
        empresa = selected_value


    ########################## MANUAIS ############################
    def clear_textbox_upload(self):
        self.textbox.delete(1.0, tk.END)
        print("UPLOAD DE ARQUIVO (Tela Principal > Upload de Arquivo):")
        print("")
        print("")
        print("O botão de upload permite selecionar um arquivo Excel (.xlsx) para carregar dados no banco de dados do sistema.")
        print("")
        print("Antes do upload, o sistema verifica se já existem dados no banco. Se existirem, você receberá uma mensagem informando que não é possível fazer o upload enquanto houver dados presentes.")
        print("")
        print("Após selecionar o arquivo, os dados são carregados no banco de dados. Uma mensagem será exibida para informar que o processo de carregamento foi concluído com sucesso.")


    def clear_textbox_periodo(self):
        self.textbox.delete(1.0, tk.END)
        print(" PERÍODO (Tela Principal > Período):")
        print("")
        print("")
        print("A opção 'Período' permite selecionar o período de lançamento. Sem escolher um período, o Bot não iniciará.")
        print("\nVocê pode escolher entre o período atual, o mês anterior e o próximo mês.")
        print("\nCertifique-se de escolher um período relacionado com o excel importado.")
        print("\nLembre-se de que é necessário carregar um arquivo antes de iniciar o processo. Se não houver dados no banco de dados, a automação não será iniciada.")

    def clear_textbox_iniciar_processo(self):
        self.textbox.delete(1.0, tk.END)
        print(" INICIAR PROCESSO (Tela Principal > Iniciar Processo): ")
        print("")
        print("")
        print("O botão 'Iniciar Processo' é fundamental para iniciar o processo de automação.")
        print("\nAntes de pressionar esse botão, é necessário escolher um período válido utilizando a opção 'Período'. Essa escolha influenciará o comportamento do processo ao interagir com os dados.")
        print("\nAlém disso, é importante verificar se há dados no banco de dados antes de iniciar o processo. Se não houver dados, a automação não será iniciada.")
        print("\nAo clicar em 'Iniciar Processo', o sistema começará a fazer os lançamentos de acordo com as informações do excel que foi feito o upload.")

    def clear_textbox_limpar_logs(self):
        self.textbox.delete(1.0, tk.END)
        print(" LIMPAR LOGS (Tela Principal > Limpar Logs): ")
        print("")
        print("")
        print("O botão 'Limpar Logs' tem a função de apagar o conteúdo da caixa de texto onde os logs são exibidos.")
        print("\nEsta ação é útil quando você deseja limpar a área de logs para ter uma visão mais clara das informações mais recentes ou para começar um novo ciclo de atividades.")
        print("\nAo clicar em 'Limpar Logs', a caixa de texto será esvaziada, removendo qualquer log anteriormente exibido.")

    def clear_textbox_exportar_logs(self):
        self.textbox.delete(1.0, tk.END)
        print(" EXPORTAR LOGS (Exportar > Exportar Logs): ")
        print("")
        print("")
        print("O botão 'Exportar Logs' permite salvar os registros exibidos na caixa de texto em um arquivo Excel (.xlsx).")
        print("\nAo clicar em 'Exportar Logs', o sistema solicitará um local para salvar o arquivo. Depois de escolher o local e fornecer um nome para o arquivo, os registros serão exportados para um arquivo Excel.")
        print("\nEssa funcionalidade é útil para armazenar e compartilhar registros de logs, possibilitando análises mais detalhadas ou a manutenção de registros para referência futura.")

    def clear_textbox_excluir_banco(self):
        self.textbox.delete(1.0, tk.END)
        print(" EXCLUÍR BANCO DE DADOS (Configurações > Excluír Banco de Dados): ")
        print("")
        print("")
        print("A opção 'Excluir Banco de Dados' permite apagar todas as entradas do banco de dados principal.")
        print("\nAo clicar em 'Excluir Banco de Dados', todas as informações contidas no banco de dados principal serão removidas. Isso inclui dados relacionados a lançamentos, configurações e outras informações armazenadas.")
        print("\nAo excluír o Banco de Dados, o Histórico de Lançamentos também será excluído, para melhor funcionamento da regra de negócio.")

    def clear_textbox_excluir_lancamentos(self):
        self.textbox.delete(1.0, tk.END)
        print(" EXCLUÍR HISTÓRICO DE LANÇAMENTOS (Configurações > Excluír Histórico de Lançamentos):")    
        print("")
        print("")
        print("A opção 'Excluir Histórico de Lançamentos' permite apagar todos os Histórico de Lançamentos de notas que foram lançadas.")
        print("\nAo clicar em 'Excluir Histórico de Lançamentos', todas as informações contidas no banco de dados de Histórico de Lançamentos serão removidas. Isso inclui registros antigos, históricos de atividades e outros dados relacionados aos Lançamentos Feitos de um arquivo que foi feito o upload.")
        print("\nAo excluír o Histórico de Lançamentos, o Banco de Dados principal não será excluído.")

    
    def clear_textbox_exportar_banco(self):
        self.textbox.delete(1.0, tk.END)
        print(" EXPORTAR BANCO DE DADOS (Exportar > Exportar Banco de Dados):")   
        print("")
        print("")
        print("O botão 'Exportar Banco de Dados' possibilita a exportação dos dados do banco (Informações do arquivo excel tratado que foi feito o upload) para um arquivo Excel (.xlsx).")
        print("\nAo clicar em 'Exportar Banco de Dados', o sistema solicitará um local para salvar o arquivo Excel. Após escolher o local e fornecer um nome para o arquivo, os dados serão exportados, permitindo análises mais detalhadas ou o armazenamento dos registros para referência futura.")

    def clear_textbox_exportar_historico(self):
        self.textbox.delete(1.0, tk.END)
        print(" EXPORTAR HISTÓRICO DE LANÇAMENTOS (Exportar > Histórico de Lançamentos): ")
        print("")
        print("")
        print("O botão 'Exportar Histórico de Lançamentos' permite salvar os registros de notas lançadas em um arquivo Excel (.xlsx).")
        print("\nAo clicar em 'Exportar Histórico de Lançamentos', o sistema solicitará um local para salvar o arquivo. Depois de escolher o local e fornecer um nome para o arquivo, os registros serão exportados para um arquivo Excel.")
        print("\nEssa funcionalidade é útil para armazenar e compartilhar registros de Históricos de Lançamentos, possibilitando análises mais detalhadas ou a manutenção de registros para referência futura.")



if __name__ == "__main__":
    app = App()
    app.mainloop()

