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


                                                ## CONFIGURAÇÕES

# VARIÁVEIS GLOBAIS -------------------------------------------------------------------
navegador = None
empresa = ""
log_conn = None
selected_period = ""
empresa = ""



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
estufa_login = "05108821000160"
estufa_senha = "Acero#2697"


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
                cnpj_str = None

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
                INSERT INTO banco (serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, "valor_contabil", cnpj_entidade)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                row['Serial'],
                row['Descrição Utilização'],
                cnpj_str,
                cpf_str,
                row['Data do doc.'],
                row['IE Entidade'],
                item_formatado,
                row['Valor contábil'],
                cnpj_entidade_str
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
            cnpj_entidade TEXT
        )
    ''')

    conn.commit()
    conn.close()

configurar_banco_dados()


## CONFIGURAÇÃO DE BANCO DE DADOS (EXCLUIR, EDITAR ETC) -------------------------------------------------------------

#######################################################
# Função para excluir a primeira linha do banco de dados
def excluir_primeira_linha():
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
    cursor_logs.execute('CREATE TABLE IF NOT EXISTS logs (serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, valor_contabil, cnpj_entidade)')
    cursor_logs.execute('INSERT INTO logs VALUES (?, ?, ?, ?, ?, ?, ?, ?)', linha_excluida)
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

    # Fechar a conexão do banco principal
    conn_principal.close()

    # Conectar ao banco de dados de logs
    conn_logs = sqlite3.connect('logs.db')
    cursor_logs = conn_logs.cursor()

    # Excluir todas as linhas da tabela de logs
    cursor_logs.execute('DELETE FROM logs')
    conn_logs.commit()
    tkinter.messagebox.showinfo("Sucesso", f"O Banco de Dados foi Excluído com Sucesso! Isso inclui o Histórico de Lançamentos.")
    

    # Fechar a conexão do banco de logs
    conn_logs.close()

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


## CONFIGURAÇÃO DE NAVEGADOR -------------------------------------------------------------------------------

#######################################################
# Função para abrir o navegador apenas uma vez
def abrir_navegador():
    global navegador
    if not navegador:
        options = Options()
        # options.add_argument("--headless")  # Habilitar o modo headless
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico, options=options)
        navegador.maximize_window()
        navegador.get("https://nfse-carmopolisdeminas.atende.net/autoatendimento/servicos/nfse/") 

        print("Iniciando Sistema...")

        # Store the ID of the original window
        original_window = navegador.current_window_handle
        assert len(navegador.window_handles) == 1

        # EFETUAR LOGIN
        # Aguardar até que a página esteja completamente carregada
        WebDriverWait(navegador, 10).until(
            lambda x: x.execute_script("return document.readyState == 'complete'")
        )
        # Preencher campos de login
        campo_login = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[3]/input"))
        )
        campo_login.send_keys("05108821000160")

        # Preencher campos de senha
        campo_senha = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/span[5]/div/input"))
        )
        campo_senha.send_keys("Acero#2697")

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

        print("Login feito com sucesso...")
        print(" ")

        # Itera sobre os handles das janelas e muda para a janela desejada
        for handle in all_handles:
            if handle != janela_principal:
                navegador.switch_to.window(handle)
                #print("Trabalhando na Nova Janela:", handle)
                break  # Sai do loop após encontrar a janela desejada

        all_nadles = navegador.window_handles

        time.sleep(2)


## AUTOMAÇÃO DENTRO DA PREFEITURA --------------------------------------------------------------------------
#######################################################
def processo():
    try:
        conn = sqlite3.connect('banco.db')
        cursor = conn.cursor()
        cursor.execute('SELECT serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, valor_contabil, cnpj_entidade FROM banco')
        linhas = cursor.fetchall()
        conn.close()

        quantidade_restante = len(linhas)
        print("Quantidade de lançamentos: " + str(quantidade_restante))
        print(" ")

        for linha in linhas:
            serial, descricao_utilizacao, cnpj, cpf, data_doc, ie_entidade, descricao_item, valor_contabil, cpj_entidade = linha

            if descricao_utilizacao == "Servico Tomado (NF de Serviços-Prefeitura)":
                # Aguardar até que a página esteja completamente carregada
                WebDriverWait(navegador, 10).until(
                    lambda x: x.execute_script("return document.readyState == 'complete'")
                )

                # Clicar no menu
                botao_menu = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='estrutura_menu_sistema']/ul/li[5]"))
                )
                botao_menu.click()

                # Clicar tomador
                botao_tomador = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, ".estrutura_submenu_sistema:nth-child(6) > .estrutura_menu_item:nth-child(1) > .estrutura_submenu_item_titulo"))
                )
                botao_tomador.click()

                # Aguardar até que a página esteja completamente carregada
                WebDriverWait(navegador, 10).until(
                    lambda x: x.execute_script("return document.readyState == 'complete'")
                )

                # Clicar Consultar
                botao_consultar = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64014_101_1']/article/div[1]/aside[1]/div/div[3]/table/tbody/tr/td[3]"))
                )
                botao_consultar.click()

                time.sleep(3)

                # Clicar Ordernar Competência
                botao_ordenar = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, "//*[@id='conteudo_64014_101_1']/article/div[1]/header/div[2]/table/tbody/tr[1]/td[3]/div[2]/div"))
                )
                botao_ordenar.click()

                time.sleep(2)

                linha_competencia_12_2023 = WebDriverWait(navegador, 10).until(
                    EC.visibility_of_element_located((By.XPATH, f"//tr[td[@aria-description='{selected_period}']]"))
                )
                linha_competencia_12_2023.click()

                time.sleep(2)

                print(f"NOTA LANÇADA: ({descricao_utilizacao}): Serial: {serial} | "
                    f"Data de Lançamento: {data_doc} | "
                    f"Descrição Item: {descricao_item} | "
                    f"Valor Contábil: R$ {valor_contabil}")
                print("")

                excluir_primeira_linha()

            else:
                time.sleep(2)
                print("Nao achou")

            time.sleep(2)
    except Exception as e:
        tkinter.messagebox.showerror("Erro", f"Ocorreu um erro com o sistema! Reinicie a aplicação.")






## INICIAR PROCESSO --------------------------------------------------------------------------

#######################################################
# Função para iniciar processo
def iniciar_processo():
    if selected_period == "":
      tkinter.messagebox.showerror("Erro", f"Escolha um período para Iniciar o Processo.")
    else:
        # Verifica se há dados no banco de dados antes de iniciar o processo
        if verificar_dados_no_banco():
            # Desabilita o botão Iniciar Processo durante a execução do processo
            print("Período de Lançamento: ", selected_period)

            # Inicie a função abrir_navegador_thread em uma thread separada
            # Inicie o navegador e o processo na mesma thread
            processo_navegador_thread = threading.Thread(target=lambda: [abrir_navegador(), processo()])
            processo_navegador_thread.start()
        else:
          tkinter.messagebox.showerror("Erro", f"Não há dados no banco de dados. Carregue um arquivo primeiro.")



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
        self.export_menu.add_command(label="Exportar Logs", command=self.exportar_logs_textbox)
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

        # Calculate current, previous, and next months
        today = datetime.datetime.now()
        current_month_year = today.strftime("%m/%Y")
        previous_month_year = (today - datetime.timedelta(days=30)).strftime("%m/%Y")
        next_month_year = (today + datetime.timedelta(days=30)).strftime("%m/%Y")

        # Create the CTkOptionMenu with the command attribute
        self.period_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Período", previous_month_year, current_month_year, next_month_year], command=self.update_selected_period)
        self.period_option_menu.set("Período")
        self.period_option_menu.grid(row=1, column=0, padx=20, pady=10)

        #####################################################################
        self.empresa_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Empresa", "Acero Agronegocios" ], command=self.update_selected_period)
        self.empresa_option_menu.set("Empresa")
        self.empresa_option_menu.grid(row=2, column=0, padx=20, pady=10)

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

    def exportar_logs_textbox(self):  # Add 'self' as the first parameter
        # Open a file dialog to choose the file location and name
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])

        # Check if the user selected a file
        if file_path:
            try:
                # Open the selected file in write mode
                with open(file_path, 'w') as file:
                    # Get the text from the textbox and write it to the file
                    logs_text = self.textbox.get(1.0, tk.END)  # Use 'self.textbox' instead of 'app.textbox'
                    file.write(logs_text)

                # Display a success message in a modal popup
                tkinter.messagebox.showinfo("Export Successful", f"Logs exported successfully to:\n{file_path}")
            except Exception as e:
                # Display an error message in a modal popup
                tkinter.messagebox.showerror("Export Error", f"Error exporting logs:\n{str(e)}")



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

