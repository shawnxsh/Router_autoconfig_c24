from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import requests
import time
import re
import pandas as pd
import openpyxl
import openpyxl.cell._writer


class FileHandler:
    def __init__(self, fileName):
        self.fileName = fileName
    
    def getNewConfigureInfo(self):
        df = pd.read_excel(self.fileName, engine="openpyxl", sheet_name='Sheet1', dtype=str)
        if len(df) == 0:
            print("\n'configurationInfo.xlsx' has no information to use, please update it...\nThis program will be shut down...")
            time.sleep(5)
            quit()
        firstRow = df.iloc[0]           
        return firstRow

    # def appendConfiguredInfo(self, configuredInfo):
    #     wb = openpyxl.load_workbook(self.fileName)
    #     sheet = wb['Sheet2']
    #     sheet.append(configuredInfo)
    #     wb.save(self.fileName)   
    #     return
    

    # def deleteConfiguredInfo(self):
    #     wb = openpyxl.load_workbook(self.fileName)
    #     sheet = wb['Sheet1']
    #     sheet.delete_rows(2, 1)
    #     wb.save(self.fileName)
    #     return
    
    def appendConfiguredInfo(self, configuredInfo):
        df = pd.read_excel(self.fileName, engine="openpyxl", sheet_name='Sheet2', dtype=str)
        df.loc[len(df.index)] = configuredInfo
        
        with pd.ExcelWriter(self.fileName, engine="openpyxl", mode='a', if_sheet_exists='replace' ) as writer:
            df.to_excel(writer, sheet_name='Sheet2', index=False)
            
        return
    

    def deleteConfiguredInfo(self):
        df = pd.read_excel(self.fileName, engine="openpyxl", sheet_name='Sheet1', dtype=str)
        df = df.iloc[1: , :]
        
        with pd.ExcelWriter(self.fileName, engine="openpyxl", mode='a', if_sheet_exists='replace' ) as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            
        return


class Router:
    def __init__(self, routerID, routerModel, pppoeUsername, pppoePassword, loginPassword):
        self.routerID = routerID
        self.routerModel = routerModel
        self.pppoeUsername = pppoeUsername
        self.pppoePassword = pppoePassword
        self.loginPassword = loginPassword
        
    def handleInputItem(self, xpath, content):
        inputItem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))
        inputItem.clear()
        inputItem.send_keys(content)
        return

    def handleButtonItem(self, xpath):
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath))).click()
        time.sleep(0.5)   
        return
    
    def switchToFrame(self, id):
        driver.switch_to.default_content()
        WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID,id)))
        return
    
    def createAccount(self):
        self.handleInputItem('/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div[3]/div[1]/div/div/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/span[2]/input[1]', self.loginPassword)
        self.handleInputItem('//*[@id="confirm-pwd-tb"]/div[2]/div[1]/span[2]/input[1]', self.loginPassword)
        self.handleButtonItem('//*[@id="local-login-button"]/div[2]/div[1]/a')
        time.sleep(1)
        return
    
    def loginAccount(self):
        self.handleInputItem('//*[@id="local-pwd-tb"]/div[2]/div[1]/span[2]/input[1]', LoginPassword)
        self.handleButtonItem('//*[@id="local-login-button"]/div[2]/div[1]/a')
        time.sleep(1)
        return

    def replaceInputItem(self, xpath, oldPattern, newPattern): 
        inputItem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))
        oldValue = inputItem.get_attribute('value')
        newValue = re.sub(oldPattern, newPattern, oldValue)
        inputItem.clear()
        inputItem.send_keys(newValue)
        return newValue
    
    def getRouterInfoByXpath(self, xpath, elementType):
        selectedItem = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))
        if elementType == 'INPUT':
            value = selectedItem.get_attribute('value')
        elif elementType == 'SPAN':
            value = selectedItem.get_attribute('innerHTML')
        return value

    def changePPPoEInfo(self, reconfigure):
        self.handleButtonItem('//*[@id="main-menu"]/div/div[1]/ul/li[2]/a')
        
        self.handleButtonItem('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[2]/div/div/div[2]/div[2]/div[2]/div/div[2]/div[1]/div[2]/div[1]/span[2]/input')
        
        self.handleButtonItem('/html/body/div[3]/div/div[3]/div/div/ul/li[3]')
        
        self.handleInputItem('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[2]/div/div/div[2]/div[2]/div[2]/div/div[2]/div[4]/div/div/div[1]/div[2]/div[1]/span[2]/input', self.pppoeUsername)
        
        self.handleInputItem('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[2]/div/div/div[2]/div[2]/div[2]/div/div[2]/div[4]/div/div/div[2]/div[2]/div[1]/span[2]/input[1]', self.pppoePassword)
        
        if reconfigure == False or (reconfigure and driver.find_element_by_xpath('//*[@id="save-data"]/div[2]/div[1]/a').is_displayed()):
            self.handleButtonItem('//*[@id="save-data"]/div[2]/div[1]/a')

        return

    def changeWirelessNetworkName(self, reconfigure):
        self.handleButtonItem('//*[@id="main-menu"]/div/div[1]/ul/li[3]/a')
        
        self.replaceInputItem('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[1]/div[2]/div[2]/div[2]/div[5]/div[2]/div[1]/div[2]/div[1]/span[2]/input', '^TP-Link', 'Occom')
        
        self.replaceInputItem('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[1]/div[2]/div[2]/div[2]/div[6]/div/div[3]/div[2]/div[1]/div[2]/div[1]/span[2]/input', '^TP-Link', 'Occom')
        
        WiFiSSID2G = self.getRouterInfoByXpath('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[1]/div[2]/div[2]/div[2]/div[5]/div[2]/div[1]/div[2]/div[1]/span[2]/input', 'INPUT')
        
        WiFiSSID5G = self.getRouterInfoByXpath('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[1]/div[2]/div[2]/div[2]/div[6]/div/div[3]/div[2]/div[1]/div[2]/div[1]/span[2]/input', 'INPUT')
        
        WiFiPassword = self.getRouterInfoByXpath('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div[1]/div[2]/div[2]/div[2]/div[5]/div[2]/div[4]/div[2]/div/div[2]/div[1]/span[2]/input', 'INPUT')
        
        if reconfigure == False or (reconfigure and driver.find_element_by_xpath('//*[@id="save-data"]/div[2]/div[1]/a').is_displayed()):
            self.handleButtonItem('//*[@id="save-data"]/div[2]/div[1]/a')
        
        return WiFiSSID2G, WiFiSSID5G, WiFiPassword 

    def getMACAddress(self):
        self.handleButtonItem('//*[@id="main-menu"]/div/div[1]/ul/li[4]/a')
        
        self.handleButtonItem('//*[@id="navigator"]/div/div[1]/ul/li[3]/a')
        
        self.handleButtonItem('//*[@id="navigator"]/div/div[1]/ul/li[3]/ul/li[3]/a')
        
        MACAddress = self.getRouterInfoByXpath('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div[2]/div[1]/div[3]/div/div/div/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/span[2]', 'SPAN')
        return MACAddress
    
    
    def presetting(self):
        
        self.handleButtonItem('/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div/div[1]/div[2]/div/div/div[2]/div[2]/div[2]/div[1]/div/div[2]/div[1]/span[2]/input')
        
        ActionChains(driver).move_to_element(driver.find_element_by_xpath('//*[@id="global-combobox-options"]/div/div[3]/div/div/ul/li[125]')).perform()
        time.sleep(1)
        
        self.handleButtonItem('//*[@id="global-combobox-options"]/div/div[3]/div/div/ul/li[125]')
        
        self.handleButtonItem('//*[@id="set-tz-confirm-btn"]/div[2]/div[1]/a')
        
        self.handleButtonItem('//*[@id="qs-disconnect-skip-btn"]/div[2]/div[1]/a/span[2]')
        
        ActionChains(driver).move_to_element(driver.find_element_by_xpath('//*[@id="qs-wireless-next-step-btn"]/div[2]/div[1]/a')).perform()
        time.sleep(1)
        
        self.handleButtonItem('//*[@id="qs-wireless-next-step-btn"]/div[2]/div[1]/a')
        
        self.handleButtonItem('//*[@id="summary-next-btn"]/div[2]/div[1]/a')
        
        WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div/div[2]/div[2]/div/div/div[1]/div[7]/div/div[1]/div[1]/div')))
        time.sleep(1)
        
        self.handleButtonItem('//*[@id="internet-create-skip"]/div[2]/div[1]/a/span[2]')
        return   

    def configuration(self, reconfigure):
        
        if reconfigure == False:
            self.presetting()
        
        self.changePPPoEInfo(reconfigure)
            
        #change WiFiSSID2G, WiFiSSID5G & get WiFiSSID2G, WiFiSSID5G, WiFiPassword
        WiFiSSID2G, WiFiSSID5G, WiFiPassword = self.changeWirelessNetworkName(reconfigure)
        
        # get MACAddress
        MACAddress = self.getMACAddress()

        # get Timestamp
        TimeConfigured = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
        
        # update excel file
        configInfoFileHandler.appendConfiguredInfo([self.routerID, self.routerModel, self.pppoeUsername, self.pppoePassword, self.loginPassword, MACAddress, WiFiSSID2G, WiFiSSID5G, WiFiPassword, TimeConfigured])
        
        configInfoFileHandler.deleteConfiguredInfo()
                    
        driver.close()
        
        return


PATH = 'chromedriver.exe' 
IPAddress = 'http://192.168.0.1/'

configInfoFileHandler = FileHandler('configurationInfo.xlsx')

while True: 
    
    OccomRouterID, RouterModel, PPPoEUsername, PPPoEPassword, LoginPassword = configInfoFileHandler.getNewConfigureInfo()
    
    routerConfig = Router(OccomRouterID, RouterModel, PPPoEUsername, PPPoEPassword, LoginPassword)
    
    try:
        response = requests.get(IPAddress)
        
        try: 
            options = webdriver.ChromeOptions()
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_argument('--headless')
            driver = webdriver.Chrome(PATH, options=options)
            driver.get(IPAddress)
            
            routerConfig.handleButtonItem('//*[@id="confirm-pwd-tb"]/div[2]/div[1]/span[2]/input[1]')

            try:
                print('\nThe router starts configuring...')
                options.headless = False
                driver = webdriver.Chrome(PATH, options=options)
                driver.get(IPAddress)
                
                routerConfig.createAccount()
                routerConfig.configuration(False)
                
                print('This router is configured successfully !!! Unplug it and connect a new one !!!')
                time.sleep(5)
            except:
                print('\nWoops...something went wrong, shut down the program and try again...')  

        except:
            print('\nThe current plugged router has been configured. Please disconnect it and process the next one.')
            command = input('Enter "Y" if you want to reconfigure the current router with new parameters: ')
            if command == 'Y':
                try:
                    response = requests.get(IPAddress)
                    
                    options = webdriver.ChromeOptions()
                    options.add_experimental_option('excludeSwitches', ['enable-logging'])
                    options.add_argument('--headless')
                    driver = webdriver.Chrome(PATH, options=options)
                    driver.get(IPAddress)
                    
                    routerConfig.handleInputItem('//*[@id="local-pwd-tb"]/div[2]/div[1]/span[2]/input[1]', LoginPassword)
                    
                    try:
                        print('\nThe router starts reconfiguring...')
                        options.headless = False
                        driver = webdriver.Chrome(PATH, options=options)
                        driver.get(IPAddress)
                        
                        routerConfig.loginAccount()
                        routerConfig.configuration(True)
                        
                        print('This router is configured successfully !!! Unplug it and connect a new one !!!')
                        time.sleep(5)
                    except:
                        print('\nWoops...something went wrong, shut down the program and try again...')
                except: 
                    continue  
            else:
                print('Please connect a new router...')
                time.sleep(2)    
    except :
        print('No router is connected yet, please connect a router...')
    