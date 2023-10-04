import pandas as pd
import os.path
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

class RpaChallengeBot:

    def __init__(self):
        self.default_directory = os.path.dirname(os.path.abspath(__file__))
        self.challenge_path = os.path.join(self.default_directory, "challenge.xlsx")
        self.screen_path = os.path.join(self.default_directory, "result.png")
        self.driver = self.setup_driver()

    def setup_driver(self):

        chrome_options = Options()
        
        chrome_options.add_experimental_option("detach", True)

        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": self.default_directory,
            "download.prompt_for_download": False,  
        })

        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        return driver
    

    def read_challenge_file(self):

        if os.path.exists(self.challenge_path):
            df = pd.read_excel(self.challenge_path)
            return df
        else:
            return None
    

    def start(self):

        df = self.read_challenge_file()

        if (df is not None):
            self.start_challenge()
            self.fill_form(df)
            self.make_screenshot()
        else:
            print(f"\nФайл 'challenge.xlsx' не найден в директории со скриптом.\nОжидалось: {self.challenge_path}\n")
            print("Выполняю скачивание файла")
            self.download_challenge()
            self.start()
    
        self.quit()

    def start_challenge(self):
        self.driver.get("https://www.rpachallenge.com/")
        start_locator = "//button[text() ='Start']"
        self.driver.find_element(By.XPATH, start_locator).click()

    def fill_form(self, df):

        for _, row in df.iterrows():
            for field, value in row.items():
                field = field.replace(" ", "")[:4]
                
                input_locator = f"//input[contains(@ng-reflect-name, 'label{field}')]"
                input_element = self.driver.find_element(By.XPATH, input_locator)
                input_element.send_keys(value)
            
            submit_locator = "//input[@value='Submit']"
            self.driver.find_element(By.XPATH, submit_locator).click()
    
    def download_challenge(self):
        self.driver.get("https://www.rpachallenge.com/")
        download_locator = "//a[text()=' Download Excel ']"
        self.driver.find_element(By.XPATH, download_locator).click()
        time.sleep(1)

    def make_screenshot(self): 
        self.driver.save_screenshot(self.screen_path)
    
    def quit(self):
        self.driver.quit()
    



