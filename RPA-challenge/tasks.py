from robocorp.tasks import task
from robocorp import workitems
from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver import ChromeOptions
from selenium.webdriver import FirefoxOptions
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import requests
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from RPA.Excel.Files import Files

@task
def extract_news_article_data():
    """Extracting news article data from specified webpage
       using search parameters entered by the user."""
    #search_phrase = get_search_phrase()
    opts = ChromeOptions()
    opts.add_argument('--no-sandbox')
    opts.add_argument('--headless')
    opts.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=opts)
    driver.implicitly_wait(10)
    open_news_website(driver)
    search_phrase = enter_search_phrase(driver)
    load_all_results(driver)
    get_search_results_list(driver, search_phrase)

def open_news_website(driver):
    """Open news website"""
    driver.implicitly_wait(10)
    driver.get("https://gothamist.com/")

def get_search_phrase():
    """Get search phrase input from user"""
    #search_phrase_input = workitems.inputs.current
    #search_phrase = input("Enter search phrase: ")
    #return search_phrase

def enter_search_phrase(driver):
    """Enter the search phrase in search input field"""
    search_phrase = "Trump"
    
    search_icon_found = False
    while (search_icon_found != True):
        try:
            search_icon = driver.find_element(By.XPATH, "//button[.//span[contains(@class, 'pi-search')]]")
            search_icon_found = True
        except (NoSuchElementException, ElementClickInterceptedException):
            print("No search button found")
    search_icon.click()
    search_field = driver.find_element(By.CLASS_NAME, "search-page-input")
    search_field.send_keys(search_phrase)
    search_button = driver.find_element(By.CLASS_NAME, "search-page-button")
    search_button.click()
    content = driver.find_element(By.ID, "resultList")
    wait = WebDriverWait(driver, timeout=30)
    wait.until(lambda d : content.is_displayed())
    return search_phrase

def load_all_results(driver):
    """Load the entire list of search results"""
    load_more = True
    """
    while (load_more):
        try:
            load_more_button = driver.find_element(By.XPATH, "//button[.//span[text()[contains(., 'Load More')]]]")
            load_more_button.click()
        except (NoSuchElementException, ElementClickInterceptedException):
            load_more = False
    """
    for i in range(7):
        try:
            load_more_button = driver.find_element(By.XPATH, "//button[.//span[text()[contains(., 'Load More')]]]")
            load_more_button.click()
        except (NoSuchElementException, ElementClickInterceptedException):
            load_more = False
        
        
def get_search_results_list(driver, search_phrase):
    """Retrieve the list of search results"""
    #content_list = driver.find_element(By.ID, "resultList")
    content = driver.find_elements(By.XPATH, "//div[contains(@class, 'v-card')][.//div[contains(@class, 'card-image-link')]]")
    i = 1
    excel = create_excel_sheet()
    #content_list_items = content[1].find_elements(By.CLASS_NAME, "col")
    
    for list_item in content:
        title_element = list_item.find_element(By.CLASS_NAME, "h2")
        title = str(title_element.text)
        description_element = list_item.find_element(By.CLASS_NAME, "desc")
        description = str(description_element.text)
        image_element = list_item.find_element(By.CLASS_NAME, "image")
        image_url = image_element.get_attribute("src")
        image_file_name = "output/downloaded_image_"+str(i)+".png"
        download_article_image(image_url, image_file_name, i)
        search_phrase_count = get_search_phrase_count(title,description, search_phrase)
        has_money = has_amount_of_money_in_title_or_desc(title, description)
        populate_excel_sheet(excel, title, description, image_file_name, search_phrase_count, has_money)
        i += 1
    save_excel_file(excel)
        

def create_excel_sheet():
    """Create excel file to store search result details."""
    excel = Files()
    excel.create_workbook("output/results.xlsx")
    excel.create_worksheet("Search Results")
    excel.append_rows_to_worksheet([["Title", "Description", "Image File Name", "Search Phrase Count", "Has Money in Title or Description?"]], "Search Results")
    return excel

def populate_excel_sheet(excel, title, desc, img_file_name, search_phrase_count, has_money):
    """Populate the excel sheet with search results details"""
    excel.append_rows_to_worksheet([[title, desc, img_file_name, search_phrase_count, has_money]], "Search Results")
    
def save_excel_file(excel):
    """Save the excel worksheet as .xlsx file"""
    excel.save_workbook()
    excel.close_workbook()


def download_article_image(image_url,image_file_name, i):
    """Download the article image file"""
    response = requests.get(image_url)
    with open(image_file_name, "wb") as file:
        file.write(response.content)

def get_search_phrase_count(title, desc, search_phrase):
    """Calculate and return the number of times search phrase 
       appears in title or description"""
    count_in_title = title.count(search_phrase)
    count_in_description =  desc.count(search_phrase)
    return count_in_title + count_in_description

def has_amount_of_money_in_title_or_desc(title, desc):
    """Check whether title or description contains any amount of money"""
    currency = ["$","dollars","USD"]
    has_money = False
    for curr in currency:
        if (curr in title or curr in desc):
            has_money = True
    return str(has_money) 

    
    

