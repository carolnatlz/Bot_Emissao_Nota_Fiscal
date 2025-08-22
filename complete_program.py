from tkinter import *
from PIL import ImageTk, Image
import tkinter as tk
from tkinter.filedialog import askopenfilename
import openpyxl
from pathlib import Path
import threading
from item import dic_itens_completo
from subitem import dic_subitem_completo
from servico import dic_servico_completo
import pandas as pd
import time
from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

window = tk.Tk()
window.title("Robô Emissão de Nota Fiscal")
window.rowconfigure(0, weight=1)
window.columnconfigure([0,1,2], weight=1)
window.geometry("700x700")

frame = Frame(window, width=700, height=150)
frame.grid(row=0,column=0,columnspan=7,sticky="NSEW")

current_path = Path().absolute()
nc_path = current_path/"nc.png"
header = ImageTk.PhotoImage(Image.open(nc_path))
header_pic = Label(frame, image = header)
header_pic.grid(row=0,column=0,columnspan=6, sticky="NSEW")

msg1 = tk.Label(text="1º Passo: Informe o código contendo item, sub-item e serviço \n Ex: 01.02.15 -> digite os pontos", bg='#BFDDF3', fg='black', borderwidth=1, relief='sunken', width=50, height=2)
msg1.grid(row=1, column=0, columnspan=6, pady=10, ipady= 7, sticky="NSEW") #NSEW -> norte/sul/leste/oeste

lista_itens = list(dic_itens_completo.keys())
lista_subitens = list(dic_subitem_completo.keys())
lista_servico = list(dic_servico_completo.keys())

servico = tk.Entry()
servico.grid(row=2,column=2, pady=10)

msg2 = tk.Label(text="2º Passo: Clique no botão abaixo e faça login com seu CPF/CNPJ e senha, \n em seguida, digite o captcha e clique em ENTRAR", bg='#BFDDF3', fg='black', borderwidth=1, relief='sunken', width=50, height=2)
msg2.grid(row=6,column=0, columnspan=6, pady=10, ipady= 7, sticky="NSEW") 

msg3 = tk.Label(text="3º Passo: Selecione o arquivo em excel(.xlsx) contendo os dados para emissão das notas", bg='#BFDDF3', fg='black', borderwidth=1, relief='sunken', width=50, height=2)
msg3.grid(row=9,column=0, columnspan=6, pady=10, ipady= 7, sticky="NSEW")

msg3 = tk.Label(text="4º Passo: Certifique-se de que o navegador está logado na Nota Carioca \n e clique para iniciar a automação", bg='#BFDDF3', fg='black', borderwidth=1, relief='sunken', width=50, height=2)
msg3.grid(row=12,column=0, columnspan=6, pady=10, ipady= 7, sticky="NSEW")

# --- modules ---

class SearchItens:
    def itens_nf(self):
        cod_servico = servico.get()

        id_item = cod_servico[:2]
        id_subitem = cod_servico[3:5]
        id_servico = cod_servico[-2:]
        
        subitem_all = id_item + "." + id_subitem

        especf_item = dic_itens_completo.get(id_item)
        especf_subitem = dic_subitem_completo.get(subitem_all)
        especf_servico = dic_servico_completo.get(cod_servico)

        self.msg_padrao_item = tk.Label(text= "valor não encontrado", wraplengt=700)
        self.msg_padrao_item.grid(row=3,column=0, columnspan=6)

        self.msg_padrao_subitem = tk.Label(text= "valor não encontrado", wraplengt=700)
        self.msg_padrao_subitem.grid(row=4,column=0, columnspan=6)

        self.msg_padrao_servico = tk.Label(text= "valor não encontrado", wraplengt=700)
        self.msg_padrao_servico.grid(row=5,column=0, columnspan=6)

        if especf_item:
            self.msg_padrao_item["text"] = id_item + " - " + especf_item
        if especf_subitem:
            self.msg_padrao_subitem["text"] = id_subitem + " - " + especf_subitem
        if especf_servico:
            self.msg_padrao_servico["text"] = id_servico + " - " + especf_servico

        print(id_item, id_subitem, id_servico)
        print(cod_servico)

    def clear_output(self):
        self.msg_padrao_item.grid_forget()
        self.msg_padrao_subitem.grid_forget()
        self.msg_padrao_servico.grid_forget()

class Actions:
    def __init__(self):
        self.driver = None
        
    def on_open(self):
        thread = threading.Thread(target=self.executar_selenium)
        thread.start()
        window.mainloop()
    
    def executar_selenium(self):
        if self.driver is None:
            chrome_options = webdriver.ChromeOptions()
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        self.driver.get("https://notacarioca.rio.gov.br/senhaweb/login.aspx")

    def on_close(self):
        if self.driver:
            self.driver.close()

    def init_ok_method(self):
        if init_ok.get() == 0:
            button_issue_invoice["state"] = "disable"
        else:
            button_issue_invoice["state"] = "active"

class Data:
    def __init__(self,file_path,action_driver):
        self.file_path = file_path
        self.action_driver = action_driver
        
    def select_file(self):
        self.file_path = askopenfilename(title="Selecione o Arquivo para Emissão da NF")
        file_path_var.set(self.file_path)
        if self.file_path:
            path_choice['text'] = f"Arquivo Selecionado: {self.file_path}"
        
    def issue_invoice(self):
        #self.action_driver.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_btEntrar"]').click()
        cod_servico = servico.get()
        
        id_item = int(cod_servico[:2]) + 1
        id_subitem = int(cod_servico[3:5]) + 1
        id_servico = int(cod_servico[-2:]) + 1

        table = pd.read_excel(self.file_path, converters={'CPF': str, 'Telefone': str})
        df = pd.DataFrame(table)
        df.columns = df.columns.str.replace(' ', '_')
        print(table)
        wait = WebDriverWait(self.driver, 5)
        
        for row in table.index:
            if table.loc[row,'Status'] == 'Não Emitido' or pd.isna(table.loc[row,'Status']):
                try:
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_pnCaminho"]/a[2]').click()
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_pnEmite"]/p/a[1]').click() #emitir nota
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbCPFCNPJTomador"]').send_keys(str(table.loc[row,'CPF'])) #CPF
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_btAvancar"]').click()
                    
                    razao_social = self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbRazaoSocial"]').get_attribute('value')
                    nome = table.loc[row,'Nome'].upper()
                    if len(razao_social) == 0:
                        self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbRazaoSocial"]').send_keys(nome) #nome
                    else:
                        pass
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbEmail"]').send_keys(table.loc[row,'Email']) #email
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbTelefone"]').send_keys(table.loc[row,'Telefone']) #telefone
                    
                    self.driver.find_element(By.XPATH, f'//*[@id="ctl00_cphCabMenu_ctrlServicos_ddlGrupos"]/option[{id_item}]').click() #cod_grupo_servico_06
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphCabMenu_ctrlServicos_ddlSubGrupos"]')))
                    self.driver.find_element(By.XPATH, f'//*[@id="ctl00_cphCabMenu_ctrlServicos_ddlSubGrupos"]/option[{id_subitem}]').click() #cod_subgrupo_servico_06.02
                    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphCabMenu_ctrlServicos_ddlServicos"]')))
                    self.driver.find_element(By.XPATH, f'//*[@id="ctl00_cphCabMenu_ctrlServicos_ddlServicos"]/option[{id_servico}]').click() #cod_servico_06.02.01

                    descricao = str(table.loc[row,'Nome_do_Produto']) + str(' ') + str(table.loc[row,'Data_de_Venda']) + str(' R$') + str(table.loc[row,'Preço_do_Produto']) + str(' ') + str(table.loc[row,'Tipo_de_Pagamento'])
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbDiscriminacao"]').send_keys(descricao) #descricao: nome do produto + data da compra + preço + parcelado ou à vista
                    
                    valor_da_nota = table.loc[row,'Preço_do_Produto']
                    valor_nota_corrigida = format(valor_da_nota, '.2f')
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbValor"]').send_keys(valor_nota_corrigida) #valor da nota
                    
                    self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_btEmitir"]').click() #emitir

                    time.sleep(1)
                    alerta = Alert(self.driver)
                    alerta.dismiss()

                    wb = openpyxl.load_workbook(self.file_path)
                    sheet = wb.active #or: wb.worksheets[0]
                    column_index = df.columns.get_loc("Status") + 1
                    row_index = row + 2
                    cell_status = sheet.cell(row=row_index, column=column_index)
                    cell_status.value = 'Emitido'
                    wb.save(self.file_path)
                    
                    #outro campo de email após aceitar
                    #self.driver.find_element(By.XPATH, '//*[@id="ctl00_cphCabMenu_tbEmail"]').send_keys(table.loc[row,'Email']) #email
                    #botão de enviar
                    
                    time.sleep(4)
                except:
                    print(f"Nota nº{row} não emitida por algum erro no cadastro")

# --- main ---

search_itens = SearchItens()
nf = search_itens.itens_nf
clear = search_itens.clear_output

actions = Actions()
open_button = actions.on_open
close_button = actions.on_close
init_ok_method = actions.init_ok_method

data = Data(file_path=askopenfilename,action_driver=actions)
select_file_inst = data.select_file
issue_invoice_inst = data.issue_invoice

button_choose = tk.Button(text="Ok", command=nf)
button_choose.grid(row=2,column=4, pady=10, sticky="NSEW")

open_chrome = tk.Button(text='Fazer Login', command=open_button)
open_chrome.grid(row=8,column=5, padx=10)

file_path_var = tk.StringVar()
path_choice = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
path_choice.grid(row=10, column=0, columnspan=6, padx=10, sticky='nsew')

button_select_file = tk.Button(text='Selecionar Arquivo', command=select_file_inst)
button_select_file.grid(row=11, column=5, padx=10, pady=5, sticky='nsew')

init_ok = tk.IntVar()
checkbox = tk.Checkbutton(text=' Estou com o navegador aberto no site da Nota Carioca', anchor='e', variable=init_ok, command=init_ok_method)
checkbox.grid(row=13, column=0, columnspan=4, padx=10, pady=10)

button_issue_invoice = tk.Button(text='Rodar Automação', command=issue_invoice_inst, state = DISABLED)
button_issue_invoice.grid(row=13, column=5, padx=10, pady=10)

close_chrome = tk.Button(text='Fechar Navegador', command=close_button)
close_chrome.grid(row=14, column=0, columnspan=6, pady=25)

window.mainloop()

'''
Como criar um ambiente virtual:
Criar o ambiente: python3.10 -m venv myenv

Ativar o ambiente:
- Windows: .|myenv|Scripts|activate (substitua as barras | pela contrabarra)
- macOS/Linux: source myenv/bin/activate

Desativar o ambiente: deactivate
Listar os pacotes: pip list

pip freeze > requirements.txt
Instalar os pacotes: pip install -r requirements.txt

Configurar o interpreter do python para abrir automaticamente nesse ambiente virtual:
Apertar Ctrl+Shift+P (ou Cmd+Shift+P no Mac) e digitar "Python: Select Interpreter"
___________________

Considerações sobre webdriver:
Para verificar a versão do chromedriver use o comando: 
chromedriver -v

Atualizar o webdriver-manager: 
pip install --upgrade webdriver-manager
'''
