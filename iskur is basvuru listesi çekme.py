import requests
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from openpyxl import Workbook

def check_internet_connection():
    try:
        requests.get("https://www.google.com")
        return True
    except:
        return False

def login_to_website(username, password):
    # Driver yolunu belirtin
    browser = webdriver.Edge('msedgedriver.exe')

    # Siteye gidin
    browser.get("https://esube.iskur.gov.tr/")
    time.sleep(2)

    # Giriş butonuna tıklayın
    login_button = browser.find_element(By.CSS_SELECTOR, ".btn.btn-large")
    login_button.click()

    # Kullanıcı bilgilerini girin
    username_input = browser.find_element(By.XPATH, '/html/body/form/div[5]/div/div/div/div/div/div/div[2]/div[3]/div[2]/div/div/div/div/div[1]/div/input')
    password_input = browser.find_element(By.XPATH, '/html/body/form/div[5]/div/div/div/div/div/div/div[2]/div[3]/div[2]/div/div/div/div/div[2]/div/input')
    username_input.send_keys(username)
    time.sleep(1)
    password_input.send_keys(password)

    # Giriş yapın
    submit_button = browser.find_element(By.XPATH, '/html/body/form/div[5]/div/div/div/div/div/div/div[2]/div[3]/div[2]/div/div/div/div/div[4]/div/div[1]/input[1]')
    submit_button.click()
    time.sleep(2)

    return browser

def get_applications(browser):
    # Başvurularınıza gidin
    browser.get("https://esube.iskur.gov.tr/Istihdam/Basvurularim.aspx")
    time.sleep(3)

    # Excel dosyası oluştur
    workbook = Workbook()
    sheet = workbook.active

    # Başlık satırını doldur
    table_headers = ["Davet Mektubu", "No", "İşveren", "Pozisyon (Meslek)", "Çalışma Yeri", "Başvuru/Seçme Tarihi", "Durumu", "Statüsü", "Başvuru/İptal"]
    sheet.append(table_headers)

    # Verileri tabloya ekle
    max_retries = 5
    retries = 0

    while retries < max_retries:
        try:
            Davet_Mektubu = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[4]')
            No_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[2]')
            isveren_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[6]')
            Pozisyon_meslek_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[7]')
            calisma_yeri_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[8]')
            basvuru_tarihi_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[9]')
            durumu_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[10]')
            statu_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[11]')
            basvuru_iptal_hucreler = browser.find_elements(By.XPATH, '//*[@id="ctl02_ctlGridBasvurular"]/tbody/tr/td[12]')

            for i in range(len(No_hucreler)):
                Davet_text = Davet_Mektubu[i].text
                Davet_text = "Başvurunuz onaylandı"
                No_text = No_hucreler[i].text
                isveren_text = isveren_hucreler[i].text
                Pozisyon_meslek_text = Pozisyon_meslek_hucreler[i].text
                calisma_yeri_text = calisma_yeri_hucreler[i].text
                basvuru_tarihi_text = basvuru_tarihi_hucreler[i].text
                durumu_text = durumu_hucreler[i].text
                statu_text = statu_hucreler[i].text
                basvuru_iptal_text = basvuru_iptal_hucreler[i].text

                sheet.append([Davet_text, No_text, isveren_text, Pozisyon_meslek_text, calisma_yeri_text, basvuru_tarihi_text, durumu_text, statu_text, basvuru_iptal_text])

            # Veriler başarıyla çekildi, döngüyü sonlandır
            break
        except:
            # Elementler bulunamadı, tekrar deneme
            retries += 1
            time.sleep(1)

    # Excel dosyasını kaydet
    workbook.save("basvurular.xlsx")

def main():
    username = input("Lütfen T.C Kimlik No giriniz: ")
    password = input("Lütfen Şifrenizi giriniz: ")

    if not check_internet_connection():
        print("İnternet bağlantısı yok.")
        return

    browser = login_to_website(username, password)
    get_applications(browser)
    browser.quit()
    print("İşlem tamamlandı.")

main()
