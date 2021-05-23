import bs4
import sys
import time
import random
import smtplib
import datetime
import threading
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException

'''
Multithreaded Best Buy Bot that will run multiple chrome windows for different products at same time
This program will send text message to user when program found product is avaiable

There are some things necessary to be done to make program run:
1. Download chromedriver that matches your current version of chrome and copy the path to line #31
2. Create Best Buy Account and update your Shipping Address and Credit Card Info
3. <Optional> Change email settings so that text message could be sent to phone. Current program set only for gmail accounts
4. Add, delete, or comment any product urls into all_products list below in line #54. Works best with 3 urls. Links has to be similar format at below:
   https://www.bestbuy.com/site/amd-ryzen-7-5800x-4th-gen-8-core-16-threads-unlocked-desktop-processor-without-cooler/6439000.p?skuId=6439000

If want to test program comment out line #106 so it doesn't click on final checkout button
'''

# Will need to download chromedriver matching your chrome version
# https://chromedriver.chromium.org/
# Copy path of the the driver in line below
chromedriver_path = 'C:/Users/<USER>/Downloads/chromedriver.exe'

first_name = 'FIRST_NAME'
last_name = 'LAST_NAME'
# DROPDOWN ONLY WORKS FOR CA ONLY
# WILL NEED TO CHANGE OPTION NUMBER FOR OTHER STATES
street_address = 'STREET_ADDRESS'
city = 'CITY'
zipcode = 'ZIPCODE'
cvv = 'CVV'

# Will need to create Best Buy account
bestbuy_username = 'BESTBUY_USERNAME'
bestbuy_password = 'BESTBUY_PASSWORD'

# Can send text to phone when item is available. Email format in link below
# https://smith.ai/blog/how-send-email-text-message-and-text-via-email
my_phone_email = '123456789@vtext.com'
# Need to enable security settings in email to send text to phone
email_user = 'EMAIL_USER'
email_password = 'EMAIL_PASSWORD'

# Works best with 3 products, but may vary depending how good your CPU is
all_products = [
    # 'https://www.bestbuy.com/site/amd-ryzen-7-5800x-4th-gen-8-core-16-threads-unlocked-desktop-processor-without-cooler/6439000.p?skuId=6439000', # Testing product with queue
    # 'https://www.bestbuy.com/site/nvidia-geforce-rtx-nvlink-bridge-for-3090-cards-space-gray/6441554.p?skuId=6441554',                            # Testing with normal product
    'https://www.bestbuy.com/site/nvidia-geforce-rtx-3080-10gb-gddr6x-pci-express-4-0-graphics-card-titanium-and-black/6429440.p?skuId=6429440',
    'https://www.bestbuy.com/site/pny-geforce-rtx-3080-10gb-xlr8-gaming-epic-x-rgb-triple-fan-graphics-card/6432658.p?skuId=6432658',
    'https://www.bestbuy.com/site/evga-geforce-rtx-3080-xc3-black-gaming-10gb-gddr6-pci-express-4-0-graphics-card/6432399.p?skuId=6432399',
    # 'https://www.bestbuy.com/site/pny-geforce-rtx-3080-10gb-xlr8-gaming-epic-x-rgb-triple-fan-graphics-card/6432655.p?skuId=6432655',
    # 'https://www.bestbuy.com/site/evga-geforce-rtx-3080-xc3-gaming-10gb-gddr6-pci-express-4-0-graphics-card/6436194.p?skuId=6436194',
    # 'https://www.bestbuy.com/site/evga-geforce-rtx-3080-xc3-ultra-gaming-10gb-gddr6-pci-express-4-0-graphics-card/6432400.p?skuId=6432400',
    # 'https://www.bestbuy.com/site/evga-geforce-rtx-3080-ftw3-gaming-10gb-gddr6x-pci-express-4-0-graphics-card/6436191.p?skuId=6436191',
    # 'https://www.bestbuy.com/site/evga-geforce-rtx-3080-ftw3-ultra-gaming-10gb-gddr6-pci-express-4-0-graphics-card/6436196.p?skuId=6436196',
]

def try_except_decorator(function):
    def wrapper(*args):
        try:
            return function(*args)
        except Exception as e:
            print(f'Failed {function.__name__}')
            print(e)
    return wrapper

class BestBuyBot():

    def __init__(self, url):
        """
        Initializes chrome driver and opens up url
        Looks for availability of item and purchases it
        """
        options = webdriver.ChromeOptions()
        options.add_argument("window-size=1280,720")
        # Need to copy chromedriver to same folder has program
        self.driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
        self.driver.implicitly_wait(10)
        self.total_seconds = 0
        self.start_time = time.time()
        self.url = url
        self.login_to_best_buy_account()
        while True:
            try:
                self.check_and_add_to_queue_bestbuy_product()
                self.send_text_message()
                self.click_add_to_cart_button_second_time()
                self.verify_item_in_cart()
                # self.login_again()
                # self.click_shipping_options()
                self.fill_shipping_info()
                # Click on "Continue to Payment Information"
                self.driver.maximize_window()
                self.driver.find_element_by_xpath('//*[@id="checkoutApp"]/div[2]/div[1]/div[1]/main/div[2]/div[2]/form/section/div/div[2]/div/div/button').click()
                self.fill_cvv_number()
                self.fill_billing_info()
                self.click_on_final_checkout() # Comment this line if you want to test code
                print('Item has been purchased')
                time.sleep(1800)
                break
            except Exception as e:
                print(e)
                print('Failed to order item. Trying Again')

    @try_except_decorator
    def login_to_best_buy_account(self):
        print('Logging in to BestBuy Account')
        login_link = 'https://www.bestbuy.com/identity/global/signin'
        self.driver.get(login_link)
        self.driver.find_element_by_xpath('//*[@id="fld-e"]').send_keys(bestbuy_username)
        self.driver.find_element_by_xpath('//*[@id="fld-p1"]').send_keys(bestbuy_password)
        self.driver.find_element_by_xpath('/html/body/div[1]/div/section/main/div[2]/div[1]/div/div/div/div/form/div[4]/button').click()
        self.wait_on_element('NextBestActionContainer', wait_type='visible', by='id')
        # self.driver.minimize_window()

    def check_and_add_to_queue_bestbuy_product(self):
        # Go to BestBuy product page
        self.driver.get(self.url)

        while True:
            # Look for "add to cart" button from html of page
            html = self.driver.page_source
            soup = bs4.BeautifulSoup(html, 'html.parser')
            add_to_cart_button = soup.find('button', {'class': 'btn btn-primary btn-lg btn-block btn-leading-ficon add-to-cart-button'})

            if add_to_cart_button:
                # Bring window to the front
                self.driver.maximize_window()
                print('Add To Cart Button Found!')
                self.wait_on_element('.add-to-cart-button', wait_type='clickable', by='css')
                self.driver.find_element_by_css_selector('.add-to-cart-button').click()
                time.sleep(5)
                self.driver.refresh()
                time.sleep(5)
                break
            else:
                # Print program runtime and refresh page
                runtime = round(time.time() - self.start_time)
                sys.stdout.write('\r')
                sys.stdout.write('Program Run Time {}'.format(str(datetime.timedelta(seconds=runtime))))
                sys.stdout.flush()
                self.driver.execute_script('window.localStorage.clear();')
                self.driver.refresh()

    @try_except_decorator
    def send_text_message(self):
        # creates SMTP session
        s = smtplib.SMTP('smtp.gmail.com', 587)
        # start TLS for security
        s.starttls()
        # Authentication
        s.login(email_user, email_password)
        # message to be sent
        message = f'You are in queue system on Bestbuy! {self.url}'
        # sending the mail
        s.sendmail(email_user, my_phone_email, message)
        print('Sent text message to phone')

    def click_add_to_cart_button_second_time(self):
        while True:
            try:
                add_to_cart = self.driver.find_element_by_css_selector('.add-to-cart-button')
                please_wait_enabled = add_to_cart.get_attribute('aria-describedby')

                if please_wait_enabled:
                    self.driver.refresh()
                    # time.sleep(random.randint(1, 3))
                else:
                    self.wait_on_element('.add-to-cart-button', wait_type='clickable', by='css')
                    self.driver.find_element_by_css_selector('.add-to-cart-button').click()
                    print('Clicked "Add To Cart" button second time')
                    break
            except(NoSuchElementException, TimeoutException) as error:
                print(f'Failed to click "Add To Cart" button second time: \n{error}')

    def verify_item_in_cart(self):
        self.driver.get('https://www.bestbuy.com/cart')
        self.wait_on_element('//*[@id="cartApp"]/div[2]/div[1]/div/div[1]/div[1]/section[1]/div[1]/div/h1', wait_type='visible', by='xpath')
        self.wait_on_element('//*[@id="cartApp"]/div[2]/div[1]/div/div[1]/div[1]/section[1]/div[1]/div/h1', wait_type='clickable', by='xpath')
        your_cart_text = self.driver.find_element_by_xpath('//*[@id="cartApp"]/div[2]/div[1]/div/div[1]/div[1]/section[1]/div[1]/div/h1').text
        # Checking the Header to see that is it says "Your Cart" instead of "Your cart is empty"
        if your_cart_text == 'Your Cart':
            self.driver.find_element_by_xpath('//*[@id="cartApp"]/div[2]/div[1]/div/div[1]/div[1]/section[2]/div/div/div[3]/div/div[1]/button').click()
            print('Verified item is in cart and clicked on "Checkout"')
        else:
            raise Exception('Your cart is empty. Trying again')

    @try_except_decorator
    def login_again(self):
        self.driver.find_element_by_xpath('//*[@id="fld-e"]').send_keys(bestbuy_username)
        self.driver.find_element_by_xpath('//*[@id="fld-p1"]').send_keys(bestbuy_password)
        self.driver.find_element_by_xpath('/html/body/div[1]/div/section/main/div[2]/div[1]/div/div/div/div/form/div[3]/button').click()

    @try_except_decorator
    def click_shipping_options(self):
        self.wait_on_element("//*[@class='btn btn-lg btn-block btn-primary button__fast-track']", wait_type='visible', by='xpath')
        self.wait_on_element("//*[@class='btn btn-lg btn-block btn-primary button__fast-track']", wait_type='clickable', by='xpath')
        self.driver.find_element_by_xpath("//*[@class='ispu-card__switch']").click()

    @try_except_decorator
    def fill_shipping_info(self):
        self.driver.find_element_by_xpath('//*[@id="consolidatedAddresses.ui_address_2.firstName"]').send_keys(first_name)
        self.driver.find_element_by_xpath('//*[@id="consolidatedAddresses.ui_address_2.lastName"]').send_keys(last_name)
        self.driver.find_element_by_xpath('//*[@id="consolidatedAddresses.ui_address_2.street"]').send_keys(street_address)
        self.driver.find_element_by_xpath('//*[@id="consolidatedAddresses.ui_address_2.city"]').send_keys(city)
        self.driver.find_element_by_xpath('//*[@id="consolidatedAddresses.ui_address_2.state"]/option[9]').click() # CA ONLY
        self.driver.find_element_by_xpath('//*[@id="consolidatedAddresses.ui_address_2.zipcode"]').send_keys(zipcode)
        print('Filled shipping info')

    @try_except_decorator
    def fill_cvv_number(self):
        self.wait_on_element('credit-card-cvv', wait_type='visible', by='id')
        self.wait_on_element('credit-card-cvv', wait_type='clickable', by='id')
        self.driver.find_element_by_id('credit-card-cvv').send_keys(cvv)
        print('Filled CVV info')

    @try_except_decorator
    def fill_billing_info(self):
        self.wait_on_element('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[2]/label/div/input', wait_type='visible', by='xpath')
        self.wait_on_element('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[2]/label/div/input', wait_type='clickable', by='xpath')
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[2]/label/div/input').send_keys(first_name)
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[3]/label/div/input').send_keys(last_name)
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[6]/div/div[1]/label/div/input').send_keys(city)
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[6]/div/div[2]/label/div/div/select/option[10]').click() # CA ONLY
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[7]/div/div[1]/label/div/input').send_keys(zipcode)
        self.driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/form/div/section/div[4]/label/div[2]/div/div/input').send_keys(street_address)
        self.driver.find_element_by_xpath('//*[@id="remember-this-information-for-next-time-generic"]').click()
        print('Filled billing info')

    def click_on_final_checkout(self):
        self.wait_on_element('//*[@id="checkoutApp"]/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/div[4]/button', wait_type='visible', by='xpath')
        self.driver.find_element_by_xpath('//*[@id="checkoutApp"]/div[2]/div[1]/div[1]/main/div[2]/div[3]/div/section/div[4]/button').click()
        print('Clicked final checkout button')

    def wait_on_element(self, desired_element, wait_type='visible', by='xpath', wait_time=30):
        """
        Waits for element of desired_xpath

        :return:
        """
        if wait_type == 'visible':
            if by == 'xpath':
                WebDriverWait(self.driver, wait_time).until(ec.visibility_of_element_located((By.XPATH, desired_element)))
            if by == 'name':
                WebDriverWait(self.driver, wait_time).until(ec.visibility_of_element_located((By.NAME, desired_element)))
            if by == 'class':
                WebDriverWait(self.driver, wait_time).until(ec.visibility_of_element_located((By.CLASS_NAME, desired_element)))
            if by == 'css':
                WebDriverWait(self.driver, wait_time).until(ec.visibility_of_element_located((By.CSS_SELECTOR, desired_element)))
            if by == 'id':
                WebDriverWait(self.driver, wait_time).until(ec.visibility_of_element_located((By.ID, desired_element)))
        if wait_type == 'clickable':
            if by == 'xpath':
                WebDriverWait(self.driver, wait_time).until(ec.element_to_be_clickable((By.XPATH, desired_element)))
            if by == 'name':
                WebDriverWait(self.driver, wait_time).until(ec.element_to_be_clickable((By.NAME, desired_element)))
            if by == 'class':
                WebDriverWait(self.driver, wait_time).until(ec.element_to_be_clickable((By.CLASS_NAME, desired_element)))
            if by == 'css':
                WebDriverWait(self.driver, wait_time).until(ec.element_to_be_clickable((By.CSS_SELECTOR, desired_element)))
            if by == 'id':
                WebDriverWait(self.driver, wait_time).until(ec.element_to_be_clickable((By.ID, desired_element)))
        if wait_type == 'invisible':
            if by == 'xpath':
                WebDriverWait(self.driver, wait_time).until(ec.invisibility_of_element_located((By.XPATH, desired_element)))
            if by == 'name':
                WebDriverWait(self.driver, wait_time).until(ec.invisibility_of_element_located((By.NAME, desired_element)))
            if by == 'class':
                WebDriverWait(self.driver, wait_time).until(ec.invisibility_of_element_located((By.CLASS_NAME, desired_element)))
            if by == 'css':
                WebDriverWait(self.driver, wait_time).until(ec.invisibility_of_element_located((By.CSS_SELECTOR, desired_element)))
            if by == 'id':
                WebDriverWait(self.driver, wait_time).until(ec.invisibility_of_element_located((By.ID, desired_element)))

if __name__ == '__main__':
    all_threads = []
    for item in all_products:
        t = threading.Thread(target=BestBuyBot, args=(item, ))
        all_threads.append(t)
    
    for thread in all_threads:
        thread.start()
