from argparse import Action
from multiprocessing.dummy import active_children
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import win32com.client as win32
import time

LISTA = ("D", "F", "J", "K", "O")

class Navigator():

    def __init__(self):
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : "C:\\Users\\guilherme.sousa\\Videos\\SAP\\"}
        chrome_options.add_experimental_option('prefs', prefs)

        self.driver = webdriver.Chrome(chrome_options=chrome_options)
        
        # init excel
        self.xl = win32.gencache.EnsureDispatch("Excel.Application")
        self.xl.Visible = False
        self.xl.WiwdowState = win32.constants.xlMaximized

    def get_site(self, site):
        self.driver.get(site)
        self.driver.maximize_window()

    def close(self):
        self.driver.close()
    
    def quit(self):
        self.driver.quit()

    def find_element(self, field):
        return self.drive.find_element(By.ID, field)
        
    def find_fill_id(self, field, text):
        self.driver.find_element(By.ID, field).send_keys(text)
        return self.driver.find_element(By.ID, field)

    def find_bt_id(self, field, tempo):
        self.driver.find_element(By.ID, field).click()
        time.sleep(tempo)

    def command(self, tecla, tempo):
        self.action = ActionChains(self.driver)
        self.action.send_keys(Keys.SHIFT, tecla).perform()
        time.sleep(tempo)

    def handle_window(self):
        self.driver.switch_to.window(self.driver.window_handles[1])
        self.close()
        self.driver.switch_to.window(self.driver.window_handles[0])
        time.sleep(5)

    def iframe_handle(self, iframe):
        self.main_window = self.driver.current_window_handle
        #Seleciona iFrame aonde fica o campo para digitar a transação
        try:
            self.iframeSAP = self.driver.find_element(By.ID, iframe)
        except NoSuchElementException:
            print("Elemento não encontrado")
        self.iframeSAP = self.driver.find_element(By.ID, iframe)
        self.driver.switch_to.frame(self.iframeSAP)
        time.sleep(15)

    def iframe_switch(self, iframe):
        self.transition_window = self.driver.current_window_handle
        self.driver.switch_to.window(self.transition_window)
        #Seleciona iFrame aonde fica o campo para digitar a transação
        try:
            self.iframeSAP = self.driver.find_element(By.ID, iframe)
        except NoSuchElementException:
            print("Elemento não encontrado")
        self.iframeSAP = self.driver.find_element(By.ID, iframe)
        self.driver.switch_to.frame(self.iframeSAP)
        time.sleep(15)

# ! Funções para a parte do Excel:

    def open_excel(self, path, name):
        self.xl.Workbooks.Open(path + name)

    def read_columns(self, i=2, b=0):
        dado = str(self.xl.Worksheets("Sheet1").Range(LISTA[b] + str(i)))

        while dado != None:
            self.xl.Worksheets("Sheet1").Range(LISTA[b] + str(i))
            i, b += 1

    def close_excel(self):
        self.xl.Workbooks.Close()

    def quit_appclication(self):
        self.xl.Workbooks.Quit()
