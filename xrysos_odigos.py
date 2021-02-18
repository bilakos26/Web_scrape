import os
import sys 
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from time import sleep
import requests
import time
import xlsxwriter

def main():
    start_time = time.time()
    #HERE IT FINDS THE PATH
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        try:
            app_full_path = os.path.realpath(__file__)
            application_path = os.path.dirname(app_full_path)
        except NameError:
            application_path = os.getcwd()

    #Here we create the variable that is going to be used to all the functions for the path
    path = os.path.join(application_path)

    #Here it takes the Chrome Path
    chrome_path = (path + "\\chromedriver.exe")

    with open(path + "\\link.txt", 'r') as link_file:
        link = link_file.read()
    

    options = webdriver.ChromeOptions() 
    options.headless = False
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    browser = webdriver.Chrome(options=options, executable_path=chrome_path)
    browser.get(link)
    sleep(2)

    text = name(browser, link)
    question(browser, link, path, text)
    sleep(3)
    browser.quit()
    time_waiting = time.time() - start_time
    hours, minutes, seconds = waitingTime(time_waiting)
    print("Η διαδικασία ολοκληρώθηκε")
    if hours == 1:
        ores = 'Ώρα'
    else:
        ores = 'Ώρες'
    print('Ο χρόνος αναμονής ήτανε: ', int(hours), ' ', ores, ' ', int(minutes), ' λεπτά ',
          'και ',int(seconds), ' δευτερόλεπτα.', sep='')
    
#Here it extracts the ulrs
def extract_urls(browser, link, path, text):
    header = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36 OPR/73.0.3856.344"}
    requests.get(link, headers = header)
    sleep(5)
    print("Αυτή η διαδικασία θα διαρκέσει μερικά λεπτά.\n", "Μην κλείσετε το πρόγραμμα......")
    print('' * 2)
    last_height = 0
    proceed = ''
    links = []
    while True:
        browser.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        sleep(1)

        #GET THE URLS
        elements = browser.find_elements_by_xpath('//a[@content]')
        sleep(1)
        for elem in elements:
            urls = elem.get_attribute('content')
            elements = browser.find_elements_by_xpath('//a[@content]')
            if urls not in links:
                links.append(urls)
        print(links)
        print('')
        sleep(2)
        new_height = browser.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            try:
                #epomeno button
                browser.find_element_by_class_name("page_next").click()
                sleep(4)
            except Exception:
                break
        last_height = new_height
        if False:
            proceed = False
        else:
            proceed = True
    sleep(10)

    #Create a folder with the name of the label
    if proceed == True:
        print("Περιμένετε μέχρι να δημιουργήθει ο φάκελος για να περαστούν τα εξαγμένα urls.\nΜην κλείσετε το πρόγραμμα...")
        print('' * 2)
        sleep(5)
        sleep(2)
        try:
            os.mkdir(path + '\\' + text)
            link_extraction = open(path + "\\" + text + '\\extracted_links.txt', 'a')
            sleep(2)
            print("Το αρχείο extracted_links.txt δημιουργήθηκε.")
            print('' * 2)
            for i in links:
                link_extraction.write(i + '\n')
            link_extraction.close()
            sleep(2)
            print('Τα urls περαστήκανε με επιτυχία.')
            print('')
        except FileExistsError:
            print('Ο φάκελος υπάρχει ήδη.')
            print('')
            link_extraction = open(path + "\\" + text + '\\extracted_links.txt', 'w')
            sleep(2)
            print("Το αρχείο extracted_links.txt δημιουργήθηκε.")
            print('' * 2)
            for i in links:
                link_extraction.write(i + '\n')
            link_extraction.close()
            sleep(2)
            print('Τα urls περαστήκανε με επιτυχία.')
            print('')

#Here it extracts the information of the urls like: names, addresses, etc
def extract_informations(browser, path, text):
    print("Ποιά είναι η πόλη που θέλετε να κάνετε εξαγωγή τα δεδομένα; (π.χ. Λάρισα)")
    city = input("Δώσε την ονομασία της πόλης: ")
    print(""*2)
    ex_f = open(path + "\\" + text + "\\extracted_links.txt", 'r')
    url = ex_f.readline()
    count = 0
    #EXCEL
    workbook = xlsxwriter.Workbook(path + "\\" + text + f"\\{text}.xlsx")
    worksheet = workbook.add_worksheet("ΚΑΤΑΛΟΓΟΣ")
    bold = workbook.add_format({"bold" : True})
    worksheet.write("A1", "ONOMA", bold) 
    worksheet.write("B1", "ΤΟΠΟΘΕΣΙΑ", bold)
    worksheet.write('C1', 'ΤΗΛΕΦΩΝΟ', bold)
    extraction = []
    while url != '':
        browser.get(url)
        sleep(3)
        try:
            #Address
            address = browser.find_element_by_class_name("streetAddressProf")
            address = address.text
            if city in address:
                #Name
                name = browser.find_element_by_id("ProfileLabel")
                name = name.text
                #Telephone number
                links2 = []
                elements = browser.find_elements_by_xpath('//a[@data-event="phone1.profile"]')
                for elem in elements:
                    urls = elem.text
                    elements = browser.find_elements_by_xpath('//a[@data-event="phone1.profile"]')
                    if urls not in links2:
                        links2.append(urls)
                #More tel numbers (βρίσκει αν υπάρχει και δευτερο νουμερο τηλεφωνου)
                try:
                    browser.find_element_by_class_name("btnphone.et-v2").click()
                    sleep(2)
                    elements2 = browser.find_elements_by_xpath('//a[@data-event="phone2.profile"]')
                    for elem in elements2:
                        urls = elem.text
                        elements2 = browser.find_elements_by_xpath('//a[@data-event="phone2.profile"]')
                        elements2 = browser.find_elements_by_xpath('//a[@data-event="mobile.profile"]')
                        if urls not in links2:
                            links2.append(urls)
                    elements3 = browser.find_elements_by_xpath('//a[@data-event="mobile.profile"]')
                    for elem in elements3:
                        urls = elem.text
                        elements3 = browser.find_elements_by_xpath('//a[@data-event="mobile.profile"]')
                        if urls not in links2:
                            links2.append(urls)
                    count += 1
                    print('%.f) Όνομα: '%(count), name, '\nΤοποθεσία: ', address, f'\nΤηλέφωνο: {links2}')
                    print("")
                    extraction.append(name)
                    extraction.append(address)
                    extraction.append(links2)
                
                except Exception:
                    count += 1
                    print('%.f) Όνομα: '%(count), name, '\nΤοποθεσία: ', address, f'\nΤηλέφωνο: {links2}')
                    print("")
                    extraction.append(name)
                    extraction.append(address)
                    extraction.append(links2)
                url = ex_f.readline()
            else:
                url = ex_f.readline()
        except Exception:
            sleep(3)
            url = ex_f.readline()
    ex_f.close()
    count = 0
    True_1 = False
    True_2 = False
    #Here it pass the content of the list to Excel and .txt files
    with open(path + "\\" + text + f"\\{text}.txt", "w") as text_f:
        col_count = 0
        for i in extraction:
            count += 1
            print("col_count = ", col_count)
            if count == 3:
                try:
                    if True_1 == False:
                        i_1 = i[0][0:10]
                        i_2 = i[1][0:10]
                        text_f.write(str(i_1) + ", ")
                        text_f.write(str(i_2) + "\n")
                        if col_count == 0:
                            worksheet.write(f"C{col_count+2}", str(i_1))
                            worksheet.write(f"C{col_count + 3}", str(i_2))
                            print("7")
                            True_1 = True
                            True_2 = False
                            col_count += 2
                        else:
                            worksheet.write(f"C{col_count+1}", str(i_1))
                            worksheet.write(f"C{col_count + 2}", str(i_2))
                            print("7")
                            True_1 = True
                            True_2 = False
                            col_count += 1
                    else:
                        i_1 = i[0][0:10]
                        i_2 = i[1][0:10]
                        text_f.write(str(i_1) + ", ")
                        text_f.write(str(i_2) + "\n")
                        worksheet.write(f"C{col_count + 2}", str(i_1))
                        worksheet.write(f"C{col_count + 3}", str(i_2))
                        print("6")
                        True_1 = True
                        True_2 = False
                        col_count += 2
                except Exception:
                    if i == []:
                        i_1 = "-"
                    else:
                        i_1 = i[0][0:10]
                    text_f.write(str(i_1) + "\n")
                    if True_2 == False:
                        worksheet.write(f"C{col_count+2}", str(i_1))
                        col_count += 2
                    else:
                        worksheet.write(f"C{col_count+1}", str(i_1))
                        col_count += 1
                    True_1 = False
                    True_2 = True
                    print("5")
                count = 0
            elif count == 2:
                if True_1 == True:
                    i = str(i)
                    i = i.strip("['")
                    i = i.strip("']")
                    text_f.write(str(i) + "\n")
                    worksheet.write(f"B{col_count + 2}", str(i))
                    print("4")
                else:
                    i = str(i)
                    i = i.strip("['")
                    i = i.strip("']")
                    text_f.write(str(i) + "\n")
                    if True_2 != True:
                        worksheet.write(f"B{col_count+2}", str(i))
                    else:
                        worksheet.write(f"B{col_count+1}", str(i))
                    print("3")
            else:
                if True_1 == True:
                    i = str(i)
                    i = i.strip("['")
                    i = i.strip("']")
                    text_f.write(str(i) + "\n")
                    worksheet.write(f"A{col_count + 2}", str(i))
                    print("2")
                else:
                    i = str(i)
                    i = i.strip("['")
                    i = i.strip("']")
                    text_f.write(str(i) + "\n")
                    if True_2 != True:
                        worksheet.write(f"A{col_count+2}", str(i))
                    else:
                        worksheet.write(f"A{col_count+1}", str(i))
                    print("1")
    workbook.close()


def question(browser, link, path, text):
    print("Ποιά διαδικασία θες να εκτελέσεις;\nΓια εξαγωγή των urls δώσε --> 1\nΓια εξαγωγή πληροφοριών από τα "
          "urls δώσε --> 2")
    print('')
    answer = int(input("Δώσε την επιλογή σου: "))
    print('')
    while answer < 1 or answer > 2:
        print("Λάθος επιλογή. Ξανά προσπάθησε.")
        answer = int(input("Δώσε την επιλογή σου: "))
        print('')
    if answer == 1:
        extract_urls(browser, link, path, text)
    else:
        extract_informations(browser, path, text)

def name(browser, link):
    name_ = browser.find_element_by_xpath("""//*[@id="MainSearchContainer"]/div/div[1]/div[1]/h1""")
    text = name_.text
    return text


def waitingTime(time_waiting):
    ores = time_waiting // 3600  
    lepta = (time_waiting // 60) % 60
    defterolepta = time_waiting % 60
    return ores, lepta, defterolepta


main()