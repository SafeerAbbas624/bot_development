# Importing all required modules of python data.
import random
import time

import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager

# print statement for reference and contact.
print("This bot is coded and developed by Mr.SAFEER ABBAS. \n https://safeerabbas624.github.io/safeerabbas/"
      "\n https://github.com/SafeerAbbas624")


# Sleep function.
def sleep(sleep_min, sleep_max):
    time.sleep(random.uniform(sleep_min, sleep_max))


# Chrome driver requirements.
chrome_options = Options()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.implicitly_wait(10)

# URL opening in Chrome browser for Propstream.
url = "https://login.propstream.com/"
driver.get(url)
driver.maximize_window()
sleep(2, 5)

# Propstream ID and password.
id = ""
password = ""

# Login element selection.
login_input_id = driver.find_element(By.NAME, "username")
login_input_pass = driver.find_element(By.NAME, "password")
submit = driver.find_element(By.CLASS_NAME, "gradient-btn")

# login Input details writing.
for character in id:
    login_input_id.send_keys(character)
    sleep(0.1, 0.2)

# login Input details writing.
for character in password:
    login_input_pass.send_keys(character)
    sleep(0.1, 0.2)

# Submitting login details.
submit.click()
driver.implicitly_wait(20)
sleep(3, 10)

# Opening search text file to get search item for searching location,
# search text file can be edited to get required data of location.
with open("Search text.txt", "r") as item:
    search_item = [line.strip() for line in item]
del item
print(f"Your search item in search text file is: {search_item}")

# getting search element.
search_tab = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div['
                                           '1]/div[1]/div[1]/div/div/div/div/input')
# sending search location appended from above search text file.
for character in search_item:
    search_tab.send_keys(character)
    sleep(0.2, 0.3)

driver.implicitly_wait(10)

sleep(3, 10)
search_tab.send_keys(Keys.ENTER)

# Getting element for different tabs of search results, all tabs data will be extracted.
pre_foreclosures = driver.find_element(By.XPATH,
                                       '//*[@id="root"]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div['
                                       '1]/div/div[1]/div/button/div/div[1]/div/div/div')
mls = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div['
                                    '1]/div/div[2]/div/button/div/div[1]/div/div/div')
auction = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div['
                                        '1]/div/div[3]/div/button/div/div[1]/div/div/div')
bank_owned = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div['
                                           '1]/div/div[4]/div/button/div/div[1]/div/div/div')
cash_byers = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div['
                                           '1]/div/div[5]/div/button/div/div[1]/div/div/div')
liens = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div['
                                      '1]/div/div[6]/div/button/div/div[1]/div/div/div')
vacant = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div['
                                       '1]/div/div[7]/div/button/div/div[1]/div/div/div')
high_equity = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/header/div['
                                            '2]/div[1]/div/div[8]/div/button/div/div[1]/div/div/div')
# Getting all elements into a list.
all_pages = [mls, auction, bank_owned, cash_byers, liens, vacant, high_equity]

# First tab of search result.
pre_foreclosures.click()
print(pre_foreclosures.text)
sleep(3, 5)

# Getting list view button element.
list_view = driver.find_element(By.XPATH,
                                '//*[@id="root"]/div/div[2]/div/div/div[3]/div[1]/div/header/div[2]/div[2]/div/div[2]')
# Getting list view click to get list view in browser.
list_view.click()
# Getting Extend element.
extend = driver.find_element(By.CLASS_NAME, '_2rE15__rightExpand')
# Clicking extend button to view all listing in browser.
extend.click()

# Heading list to append it as Headers in Excel file for output of extracted data.
headings = ['S.No', 'Listing Type', 'Address', 'City', 'State', 'Zip', 'APN', 'Beds', 'Baths', 'Property Type',
            'Estimated Value', 'List Date', 'List Price', 'Status', 'Document Type', 'Default Amount', 'Last Updated',
            'Sale Amount', 'Sale Date', 'Estimated Equity']

# While loop to loop every page in the selected tab.
while True:
    # Try Except exception not to get (NoSuchElementException) this error.
    try:
        # Getting element of tables data in html.
        data = driver.find_element(By.XPATH,
                                   "/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/section/div["
                                   "2]/div/div/div/div/div[2]/div/table").get_attribute('outerHTML')

        # Appending tables from html to dataframe using pandas lib.
        list1 = pd.read_html(data)

        # Try Except condition not to get FileNotFoundError.
        try:
            # Opening propstream_data Excel file if present, and appending dataframe to it.
            with pd.ExcelWriter('Propstream_data.xlsx', mode='a', if_sheet_exists="overlay") as writer:
                list1[0].to_excel(writer, index=True, header=False, startrow=writer.sheets['Sheet1'].max_row)
        except FileNotFoundError:
            # Writing new propstream_data file if not present in above code. Appending headers list also.
            with pd.ExcelWriter('Propstream_data.xlsx', mode='w') as writer:
                list1[0].to_excel(writer, header=headings)
        print("appended data into Propstream_data file")

        # Getting Next button element.
        next_button = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div/div/div[3]/div[1]/div/section/div['
                                                    '2]/div/div/div/div/div[3]/button[3]/span')
        # Executing click on next button forcefully.
        driver.execute_script("arguments[0].click();", next_button)

        # If next button is disabled it will move to select next tab. using if statement.
        if 'disabled' in driver.find_element(By.XPATH,
                                             '/html/body/div[1]/div/div[2]/div/div/div[3]/div[1]/div/section/div['
                                             '2]/div/div/div/div/div[3]/button[3]').get_attribute('class'):
            print("No more pages")
            # using if else statement to get all pages in all tabs.
            if not all_pages:
                break
            if int(all_pages[0].text) != 0:
                all_pages[0].click()
                driver.implicitly_wait(10)
                sleep(3, 6)
                list_view.click()
                sleep(3, 5)
                try:
                    extend.click()
                    sleep(5, 10)
                except StaleElementReferenceException:
                    extend = driver.find_element(By.CLASS_NAME, '_2rE15__rightExpand')
                    extend.click()
                    sleep(3, 5)
                del all_pages[0]

            else:
                del all_pages[0]
                pass

    except NoSuchElementException:
        # using if else statement to get all pages in all tabs.
        print("No more pages")
        if not all_pages:
            break
        if int(all_pages[0].text) != 0:
            all_pages[0].click()
            driver.implicitly_wait(10)
            sleep(3, 6)
            list_view.click()
            sleep(3, 5)
            try:
                extend.click()
                sleep(5, 10)
            except StaleElementReferenceException:
                extend = driver.find_element(By.CLASS_NAME, '_2rE15__rightExpand')
                extend.click()
                sleep(3, 5)
            del all_pages[0]

        else:
            del all_pages[0]
            pass
print("All data appended to Propstream_data excel file. Please find it in same directory.")
time.sleep(15)

# Starting Coles' bot to extract data.
print("Starting Cloes bot.")

# URL for Cloes site.
url2 = "https://realtyresource.coleinformation.com/login/"

# ID and Password for coles site.
coles_ID = ""
coles_pass = ""

# Opening coles site using above URL.
driver.get(url2)
driver.maximize_window()
sleep(1.3, 2.5)

# Opening Propstream_data file to get required data for Cole's site input.
data = pd.read_excel("Propstream_data.xlsx", sheet_name='Sheet1')
address = data["Address"].tolist()
zip_code = data["Zip"].tolist()
state = data["State"].tolist()
print("Getting Data from Propstream_data file")

# Getting element of login and sending ID and password to login.
login_id_input = driver.find_element(By.XPATH, '/html/body/form/div[3]/div/div/input[1]')
for word in coles_ID:
    login_id_input.send_keys(word)
    sleep(0.1, 0.2)

# Getting element of login and sending ID and password to login.
login_pass_input = driver.find_element(By.XPATH, '/html/body/form/div[3]/div/div/input[2]')
for char in coles_pass:
    login_pass_input.send_keys(char)
    sleep(0.1, 0.2)

# agreement checkbox Element.
driver.find_element(By.XPATH, '/html/body/form/div[3]/div/div/p[1]/input').click()

# login details submit button element.
driver.find_element(By.XPATH, '/html/body/form/div[3]/div/div/input[3]').click()
driver.implicitly_wait(10)
sleep(1.3, 2.5)

# prospect iq tab clicking.
driver.find_element(By.XPATH,
                    '/html/body/form/div[4]/div/div/div[6]/div[1]/div/div/ul/li[2]/a/span/span').click()
driver.implicitly_wait(5)
sleep(2, 4)

#  Clicking By neighborhood radio button.
driver.find_element(By.XPATH, '//*[@id="primary-entry"]/div[3]/div[1]/label').click()
driver.implicitly_wait(5.1)
sleep(2, 4)

code = 0
state_name = 0

# Heading list to append it as Headers in Excel file for output of extracted data.
headers = [
    'Last Name', 'First Name', 'House #', 'Dir', 'Street Name', 'Apt/Box', 'City', 'ST', 'Zip', 'Area Code',
    'Phone', 'Cell Area Code', 'Cell Phone', 'Email Address', 'Map Link', 'SEL']

# Starting for loop to pass address and zip codes into cole's site to get result against passing data
for add in address:

    # Splitting address and separating street no and house no into different lists.
    add_split = list(add.split())
    print(add_split)
    house_num = add_split[0]
    print(house_num)
    add_split.remove(add_split[0])

    # Try Except condition not to get IndexError.
    try:
        add_split.remove(add_split[-1])
    except IndexError:
        pass
    print(add_split)
    house_name_input_keys = " ".join(add_split)
    print(house_name_input_keys)
    print(state[state_name])

    # Getting House no input element and sending required data into it.
    house_num_input = driver.find_element(By.XPATH, '/html/body/form/div[4]/div/div/div[6]/div[2]/div[1]/div[2]/div['
                                                    '1]/div[2]/div[3]/div[3]/div[2]/div[1]/div[1]/input')
    house_num_input.send_keys(Keys.CONTROL, 'a')
    house_num_input.send_keys(Keys.BACKSPACE)
    house_num_input.send_keys(house_num)
    sleep(0.3, 0.8)

    # Getting Street no input element and sending required data into it.
    street_name_input = driver.find_element(By.XPATH, '/html/body/form/div[4]/div/div/div[6]/div[2]/div[1]/div[2]/div['
                                                      '1]/div[2]/div[3]/div[3]/div[2]/div[2]/div/input')
    street_name_input.send_keys(Keys.CONTROL, 'a')
    street_name_input.send_keys(Keys.BACKSPACE)
    street_name_input.send_keys(house_name_input_keys)
    sleep(0.3, 0.8)

    # Getting Zip Code input element and sending required data into it.
    zip_code_input = driver.find_element(By.XPATH, '/html/body/form/div[4]/div/div/div[6]/div[2]/div[1]/div[2]/div['
                                                   '1]/div[2]/div[3]/div[3]/div[2]/div[4]/div/input')
    zip_code_input.send_keys(Keys.CONTROL, 'a')
    zip_code_input.send_keys(Keys.BACKSPACE)
    zip_code_input.send_keys(zip_code[code])
    code += 1
    if zip_code[code] == 'Zip':
        code += 1
    sleep(0.3, 0.8)

    # Selecting dropdown state element and checking with if statement whether which one to select.
    select = Select(driver.find_element(By.XPATH, '/html/body/form/div[4]/div/div/div[6]/div[2]/div[1]/div[2]/div['
                                                  '1]/div[2]/div[3]/div[3]/div[2]/div[3]/div[2]/select'))
    if state[state_name] == "NY":
        select.select_by_visible_text("New York")
        sleep(0.3, 0.8)
    if state[state_name] == "AL":
        select.select_by_visible_text("Alabama")
        sleep(0.3, 0.8)
    if state[state_name] == "GA":
        select.select_by_visible_text("Georgia")
        sleep(0.3, 0.8)
    if state[state_name] == "SC":
        select.select_by_visible_text("South Carolina")
        sleep(0.3, 0.8)
    if state[state_name] == "TN":
        select.select_by_visible_text("Tennessee")
        sleep(0.3, 0.8)
    if state[state_name] == "FL":
        select.select_by_visible_text("Florida")
        sleep(0.3, 0.8)

    # Getting html tag to scroll through webpage.
    html = driver.find_element(By.TAG_NAME, 'html')
    # Scrolling down.
    html.send_keys(Keys.PAGE_DOWN)
    html.send_keys(Keys.PAGE_DOWN)
    driver.implicitly_wait(3)
    sleep(0.3, 0.8)

    # Getting search button element.
    search_button = driver.find_element(By.XPATH, '/html/body/form/div[4]/div/div/div[6]/div[2]/div[1]/div[2]/div['
                                                  '1]/div[ '
                                                  '7]/input')
    # Clicking Search button to get data against input data.
    search_button.click()
    driver.implicitly_wait(10)
    sleep(0.3, 0.8)

    # Try Except not to get NoSuchElementException.
    try:

        # Getting element of tables data in html.
        comm_data = driver.find_element(By.XPATH,
                                        '/html/body/form/div[7]/div[1]/div[2]/div[2]/div[1]/div/div[2]/div[2]/table') \
            .get_attribute('outerHTML')

        # Appending tables from html to dataframe using pandas lib.
        df = pd.read_html(comm_data)

        # Try Except condition not to get FileNotFoundError.
        try:
            # Opening propstream_data Excel file if present, and appending dataframe to it.
            with pd.ExcelWriter('coles_data.xlsx', mode='a', if_sheet_exists="overlay") as writer:
                df[0].to_excel(writer, index=True, header=False, startrow=writer.sheets['Sheet1'].max_row)
        except FileNotFoundError:

            # Writing new Coles_data file if not present in above code. Appending headers list also.
            with pd.ExcelWriter('coles_data.xlsx', mode='w') as writer:
                df[0].to_excel(writer, header=headers)
        print("appended the data into coles_data file")
        sleep(3, 5)

        # Closing overlay popup showing result.
        driver.find_element(By.ID, 'cboxClose').click()
        sleep(3, 5)

        # Scrolling up to top on webpage.
        html.send_keys(Keys.PAGE_UP)
        html.send_keys(Keys.PAGE_UP)
        sleep(3, 5)

    except NoSuchElementException:

        # Closing overlay popup showing result.
        sleep(3, 5)
        driver.find_element(By.ID, 'cboxClose').click()

        # Scrolling up to top on webpage.
        sleep(3, 5)
        html.send_keys(Keys.PAGE_UP)
        html.send_keys(Keys.PAGE_UP)
        sleep(3, 5)
        # Printing no data statement.
        print(f"No Data Found against this: {add}")
        pass

    state_name += 1
    if state[state_name] == 'State':
        state_name += 1

sleep(5, 10)

print("All data of Propstream and Coles extracted successfully, please find Propstream_data and Coles_data files.")

# print statement for reference and contact.
print("\n \n \n \n \n \n \nThis bot is coded and developed by Mr.SAFEER ABBAS.\n "
      "https://safeerabbas624.github.io/safeerabbas/ "
      "\n https://github.com/SafeerAbbas624")
driver.close()
