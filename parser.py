import openpyxl
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time

driver = webdriver.Chrome(ChromeDriverManager().install())  # pip install webdriver_manager


def poisk_skolko_ehat_i_rastoyaniya(otkuda="55,439349 37,74569", kuda="55,42 38,2679129"):
    driver.get("https://yandex.ru/maps/10735/krasnogorsk/?ll=37.330192%2C55.831099&mode=routes&rtext=&rtt=auto&z=13")
    assert "Яндекс" in driver.title
    # time.sleep(2)
    try:
        elem = driver.find_element_by_xpath(
            '/html/body/div/div/div/div/div/div/div/div/div/div/div/div/div/form/div/div/div/div[1]/div/div/div/div/span/span/input')
    except:
        elem = driver.find_element_by_xpath(
            "/html/body/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/form/div/div/div/div[1]/div/div/div/div/span/span/input")

    #
    elem.clear()
    elem.send_keys(otkuda)
    time.sleep(1)
    elem.send_keys(Keys.DOWN)

    elem.send_keys(Keys.RETURN)
    try:
        elem = driver.find_element_by_xpath(
            "/html/body/div/div/div/div/div/div/div/div/div/div/div/div/div/form/div/div/div/div[2]/div/div/div/div/span/span/input")
    except:
        elem = driver.find_element_by_xpath(
            "/html/body/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/form/div/div/div/div[2]/div/div/div/div/span/span/input")
    elem.clear()
    elem.send_keys(kuda)
    time.sleep(1)
    elem.send_keys(Keys.DOWN)

    elem.send_keys(Keys.RETURN)
    time.sleep(2)
    try:
        skok_ehat = driver.find_element_by_xpath(
            "/html/body/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div[1]")
    except:
        skok_ehat = ""
    try:
        rastoyanie = driver.find_element_by_xpath(
            '/html/body/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div[2]')
    except:
        rastoyanie = ""

    return skok_ehat.text + " " + rastoyanie.text


puth = "shablon.xlsx"
wb = openpyxl.load_workbook(puth)

sheet = wb.active

otkuda_mass = []
kuda_mass = []

for cell in list(sheet.columns)[0]:
    if str(cell.value) != "" and str(cell.value) != "None":
        print(cell.value)
        otkuda_mass.append(str(cell.value))

for cell in list(sheet.columns)[1]:
    if str(cell.value) != "" and str(cell.value) != "None":
        print(cell.value)
        kuda_mass.append(str(cell.value))

print(len(otkuda_mass))

massiv_rast_i_killometrov = [" ", " "]
for i in range(2, len(otkuda_mass) - 1):
    print(otkuda_mass[i], kuda_mass[i], "В обработке")
    massiv_rast_i_killometrov.append(poisk_skolko_ehat_i_rastoyaniya(otkuda_mass[i], kuda_mass[i]))
    time.sleep(1)

for row in range(2, len(massiv_rast_i_killometrov)):
    cell = sheet.cell(row=row + 1, column=3)
    cell.value = massiv_rast_i_killometrov[row]

driver.close()
wb.save(puth)
print("Записано в шаблон")
