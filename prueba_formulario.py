import unittest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from time import sleep

class usando_unittest(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Chrome(
            executable_path=r"C:\Driver_Chrome\chromedriver.exe")
        driver = self.driver
        driver.maximize_window()
        # driver.set_window_size(2880, 1312) #minimiza o maximiza una ventana
        driver.get('https://demoqa.com/automation-practice-form')

    def test_form(self):
        driver = self.driver

        #Manejo archivo excel
        filessheet = "./a.xlsx"
        wb = load_workbook(filessheet)
        hojas = wb.get_sheet_names()
        print(hojas)
        datos = wb.get_sheet_by_name("Hoja1")
        wb.close()
  
        #Recorrer el archivo de excel
        for i in range(1, 3): #Filas excel

            #Columnas excel
            name, last, email, gender, number, date, subject, check, picture, address, state, city = datos[F'A{i}:L{i}'][0]

            #Llena un formulario con datos que se toman de un archivo de excel
            fecha = date.value.split('/')
            driver.find_element(By.ID, "firstName").send_keys(name.value)
            driver.find_element(By.ID, "lastName").send_keys(last.value)
            driver.find_element(By.ID, "userEmail").send_keys(email.value)

            # --------------------------------------------------------------------

            if gender.value == "male":
                ActionChains(driver).double_click(driver.find_element(By.ID, "gender-radio-1")).perform()
            elif gender.value == "female":
                ActionChains(driver).double_click(driver.find_element(By.ID, "gender-radio-2")).perform()
            else:
                ActionChains(driver).double_click(driver.find_element(By.ID, "gender-radio-3")).perform()
                
            # --------------------------------------------------------------------

            driver.find_element(By.ID, "userNumber").send_keys(number.value)

            # --------------------------------------------------------------------

            if  fecha[1] == '1':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Jan ' + fecha[0])).perform()
            elif fecha[1] == '2':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Feb ' + fecha[0])).perform()
            elif fecha[1] == '3':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Mar ' + fecha[0])).perform()
            elif fecha[1] == '4':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Apr ' + fecha[0])).perform()
            elif fecha[1] == '5':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' May ' + fecha[0])).perform()
            elif fecha[1] == '6':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Jun ' + fecha[0])).perform()
            elif fecha[1] == '7':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Jul ' + fecha[0])).perform()
            elif fecha[1] == '8':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Aug ' + fecha[0])).perform()
            elif fecha[1] == '9':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Sep ' + fecha[0])).perform()
            elif fecha[1] == '10':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Oct ' + fecha[0])).perform()
            elif fecha[1] == '11':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Nov ' + fecha[0])).perform()
            elif fecha[1] == '12':
                ActionChains(driver).click(driver.find_element(By.ID, "dateOfBirthInput").send_keys(fecha[2] + ' Dec ' + fecha[0])).perform()

            # -----------------------------------------------------------------------------------------

            driver.find_element(By.ID, "subjectsInput").send_keys(subject.value)
            driver.find_element(By.ID, "subjectsInput").send_keys(Keys.TAB)
        
            # -----------------------------------------------------------------------------------------

            if check.value == "music":
                ActionChains(driver).click(driver.find_element(By.ID, "hobbies-checkbox-3")).perform()
            elif check.value == "reading":
                ActionChains(driver).click(driver.find_element(By.ID, "hobbies-checkbox-2")).perform()
            else:
                ActionChains(driver).click(driver.find_element(By.ID, "hobbies-checkbox-1")).perform()

            # ------------------------------------------------------------------------------------------

            driver.find_element(By.ID, "uploadPicture").send_keys(picture.value)
            driver.find_element(By.ID, "currentAddress").send_keys(address.value)
            driver.execute_script(f"arguments[0].innerHTML = '{state.value}'", driver.find_element(By.XPATH, "//*[@id='state']/div/div[1]/div[1]"))
            driver.execute_script(f"arguments[0].innerHTML = '{city.value}'", driver.find_element(By.XPATH, "//*[@id='city']/div/div[1]"))
            driver.execute_script("window.location.reload()")
            # driver.find_element(By.ID, "submit").click() --> click en el submit

    def tearDown(self):
        self.driver.close()

if __name__ == '__main__':
    unittest.main()
