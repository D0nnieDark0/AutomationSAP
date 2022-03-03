from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from datetime import date, timedelta
import win32com.client as win32
import time
import os

LISTA = ("D", "F", "J", "K", "O")
PATH = "C:\\Users\\guilherme.sousa\\Videos\\SAP\\"

yesterday = date.today() - timedelta(days=1)
yesterday_day = yesterday.strftime('%d.%m.%y')

class Sharepoint():

    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.get("https://planinternational.sharepoint.com/sites/brazilit/Lists/FromSap/AllItems.aspx?origin=createList")
        self.driver.maximize_window()
        self.xl = win32.gencache.EnsureDispatch("Excel.Application")
        self.xl.Visible = False
        time.sleep(10)
        #self.xl.WindowState = win32.constants.xlMaximized

    def check_download(self):
        os.chdir(PATH)

    def bt_sim(self):
        self.clicar = self.driver.find_element_by_id("idSIButton9")
        self.clicar.click()
        time.sleep(10)

    def login(self, username, password):
        self.user = self.driver.find_element_by_id("okta-signin-username")
        self.user.send_keys(username)
        self.senha = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "okta-signin-password")))
        self.senha.send_keys(password)
        self.senha.send_keys(Keys.RETURN)
        time.sleep(10)

        self.bt_sim()

    def login_microsoft(self, email):
        self.micro = self.driver.find_element_by_name("loginfmt")
        self.micro.send_keys(email)
        self.micro.send_keys(Keys.RETURN)
        time.sleep(10)

    def novo_form(self):
        self.novo = self.driver.find_element_by_name("Novo")
        self.novo.click()
        time.sleep(5)

    def salvar_form(self):
        self.salvar = self.driver.find_element_by_name("Salvar")
        self.salvar.click()

    def input_purchase(self, compra):
        self.compra = self.driver.find_elements_by_tag_name("input")
        self.compra[3].send_keys(compra)

    def input_criado(self, criado):
        self.criado = self.driver.find_elements_by_tag_name("input")
        self.criado[4].send_keys(criado)

    def input_conteudo(self, conteudo):
        self.conteudo = self.driver.find_elements_by_tag_name("input")
        self.conteudo[5].send_keys(conteudo)

    def input_quantidade(self, quantidade):
        self.quantidade = self.driver.find_elements_by_tag_name("input")
        self.quantidade[6].send_keys(quantidade)

    def input_valor(self, valor):
        self.valor = self.driver.find_elements_by_tag_name("input")
        self.valor[7].send_keys(valor)

    def open_excel(self):
        os.chdir(PATH)
        check = os.listdir()
        self.xl.Workbooks.Open(PATH + check[0])
        time.sleep(10)

    def fill_form(self, i=2):

        dado = str(self.xl.Worksheets("Sheet1").Range("D" + str(i)))

        while dado != None:
            self.novo_form()
            self.input_purchase(str(self.xl.Worksheets("Sheet1").Range("D" + str(i))))
            self.input_criado(str(self.xl.Worksheets("Sheet1").Range("F" + str(i))))
            self.input_conteudo(str(self.xl.Worksheets("Sheet1").Range("J" + str(i))))
            self.input_quantidade(int(self.xl.Worksheets("Sheet1").Range("K" + str(i))))
            self.input_valor(int(self.xl.Worksheets("Sheet1").Range("O" + str(i))))
            self.salvar_form()
            time.sleep(5)
            i += 1

    def close_excel(self):
        self.xl.Workbooks.Close()

    def quit_application(self):
        self.xl.Workbooks.Quit()