from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import StaleElementReferenceException
import time
import yaml

class ChromeDriver:

    def __init__(self):
        """
        Initializes chrome driver and opens up url

        Creates self for driver
        :return:
        """
        options = webdriver.ChromeOptions()

        # uncomment to keep window open
        # options.add_experimental_option("detach", True)
        # options.add_experimental_option('prefs', {'download.default_directory': os.path.join(os.getcwd(), 'data'), })
        options.add_argument("window-size=1280,720")

        # Need to copy chromedriver to this location
        self.driver = webdriver.Chrome(options=options)
        self.driver.implicitly_wait(10)

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

    def perform_action_list(self, action_list):
        for action in action_list:
            # if value is a 'str' -> click action
            if type(action) is str:
                self.driver.find_element_by_xpath(action).click()
            # if value is a 'dict' -> send_keys action
            if type(action) is dict:
                self.driver.find_element_by_xpath(list(action.keys())[0]).send_keys(list(action.values())[0])
            # if value is a 'list' -> wait on element OR dropdown action
            if type(action) is list:
                if action[0] == 'WAIT':
                    self.wait_on_element(action[2], wait_type=action[1])
                elif action[0] == 'SLEEP':
                    time.sleep(action[1])
                else:
                    tab_element = self.driver.find_element_by_xpath(action[0])
                    dropdown_element = self.driver.find_element_by_xpath(action[1])
                    ActionChains(self.driver).move_to_element(tab_element).move_to_element(
                        dropdown_element).click().perform()

    def get_working_element_by_name(self, element_name):
        elements = self.driver.find_elements_by_name(element_name)
        for index, name in enumerate(elements):
            try:
                if name.is_enabled() and name.is_displayed():
                    return name
            except ElementNotInteractableException as e:
                print('Failed', index)
                print(e)

    def wait_til_available_and_click(self, xpath, wait_type='visible', by='xpath', wait_time=30):
        # Waiting until element is visible
        self.wait_on_element(xpath, wait_type=wait_type, by=by)
        start_time = time.time()
        # Gets web element of the desired element
        if by == 'xpath':
            element = self.driver.find_element_by_xpath(xpath)
        if by == 'name':
            element = self.driver.find_element_by_name(xpath)
        # Waits til element is enabled and displayed and continously try to click it
        while wait_time > int(time.time() - start_time):
            if element.is_enabled() and element.is_displayed():
                while True:
                    try:
                        element.click()
                        return
                    except(ElementNotInteractableException,
                           ElementClickInterceptedException,
                           StaleElementReferenceException) as e:
                        # print(e)
                        continue

    def check_for_other_windows(self):
         if len(self.driver.window_handles) > 1:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    def take_screenshot(self, pic_path):
        """
        Takes screenshot of the current selenium Chrome Driver and stores into temp_folder

        :param pic_path:
        :return:
        """
        self.driver.save_screenshot('{}.png'.format(pic_path))

    def article_path_word_count(self, article_path, by='xpath'):
        word_count = 0
        tags=['p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'span']
        try:
            if by == 'xpath':
                article = self.driver.find_element_by_xpath(article_path)    
            elif by == 'name':
                article = self.driver.find_element_by_name(article_path)  
            elif by == 'class':
                article = self.driver.find_element_by_class_name(article_path)  
            elif by == 'id':
                article = self.driver.find_element_by_id(article_path)
        except:
            print(article_path, 'does not exist')
            return word_count
        for t in tags:
            # Finds text if tag exists
            if '<{}>'.format(t) in article.get_attribute('innerHTML'):
                elements = article.find_elements_by_tag_name(t)
                for el in elements:
                    word_count += len(el.text.split(' '))
        return word_count

    def get_nested_elements(self, nested_elements, elements=[], idx=0):
        # Return Attributes if last index is attribute
        if idx == len(nested_elements)-1 and nested_elements[idx].get('attribute'):
            return [el.get_attribute(nested_elements[idx]['attribute']) for el in elements]
        # Return when nested through all elements
        if idx == len(nested_elements):
            return elements

        # Initialize variables
        all_elements = None
        by = nested_elements[idx].get('by')
        element_path = nested_elements[idx].get('prop')
        s = nested_elements[idx].get('elements')

        # Get the first element or elements in nested_elements
        if idx == 0:
            if not s:
                try:
                    if by == 'xpath':
                        element = self.driver.find_element_by_xpath(element_path)    
                    elif by == 'name':
                        element = self.driver.find_element_by_name(element_path)  
                    elif by == 'class':
                        element = self.driver.find_element_by_class_name(element_path)  
                    elif by == 'id':
                        element = self.driver.find_element_by_id(element_path)
                    elif by == 'tag':
                        element = self.driver.find_element_by_tag_name(element_path)
                # Raise Exception if element is not found
                except:
                    raise Exception(element_path, 'does not exist')
                all_elements = [element]
            else:
                if by == 'xpath':
                    element = self.driver.find_elements_by_xpath(element_path)    
                elif by == 'name':
                    element = self.driver.find_elements_by_name(element_path)  
                elif by == 'class':
                    element = self.driver.find_elements_by_class_name(element_path)  
                elif by == 'id':
                    element = self.driver.find_elements_by_id(element_path)
                elif by == 'tag':
                    element = self.driver.find_elements_by_tag_name(element_path)

                # Raise Exception if element is not found
                if not element:
                    raise Exception(element_path, 'does not exist')
                all_elements = element
        # Iterate through all the elements
        else:
            all_elements = []
            for el in elements:
                if not s:
                    try:
                        if by == 'xpath':
                            element = el.find_element_by_xpath(element_path)    
                        if by == 'name':
                            element = el.find_element_by_name(element_path)  
                        if by == 'class':
                            element = el.find_element_by_class_name(element_path)  
                        if by == 'id':
                            element = el.find_element_by_id(element_path)
                        if by == 'tag':
                            element = el.find_element_by_tag_name(element_path)
                    # Raise Exception if element is not found
                    except:
                        raise Exception(element_path, 'does not exist')
                    all_elements.append(element)
                else:
                    if by == 'xpath':
                        element = el.find_elements_by_xpath(element_path)    
                    if by == 'name':
                        element = el.find_elements_by_name(element_path)  
                    if by == 'class':
                        element = el.find_elements_by_class_name(element_path)  
                    if by == 'id':
                        element = el.find_elements_by_id(element_path)
                    if by == 'tag':
                        element = el.find_elements_by_tag_name(element_path)

                    # Raise Exception if element is not found
                    if not element:
                        raise Exception(element_path, 'does not exist')
                    all_elements.extend(element)
        return self.get_nested_elements(nested_elements, elements=all_elements, idx=idx+1)

    def find_element(self, element):
        by = element.get('by')
        prop = element.get('prop')
        elements = element.get('elements')
        
        if not elements:
            if by == 'xpath':
                element = self.driver.find_element_by_xpath(prop)    
            elif by == 'name':
                element = self.driver.find_element_by_name(prop)  
            elif by == 'class':
                element = self.driver.find_element_by_class_name(prop)  
            elif by == 'id':
                element = self.driver.find_element_by_id(prop)
            elif by == 'tag':
                element = self.driver.find_element_by_tag_name(prop)
            else:
                raise Exception('Invalid by: {}'.format(by))
        else:
            if by == 'xpath':
                element = self.driver.find_elements_by_xpath(prop)    
            elif by == 'name':
                element = self.driver.find_elements_by_name(prop)  
            elif by == 'class':
                element = self.driver.find_elements_by_class_name(prop)  
            elif by == 'id':
                element = self.driver.find_elements_by_id(prop)
            elif by == 'tag':
                element = self.driver.find_elements_by_tag_name(prop)
            else:
                raise Exception('Invalid by: {}'.format(by))

            # Raise Exception if element is not found
            if not element:
                raise Exception(prop, 'does not exist')
        return element

    def click_element(self, element_param):
        if isinstance(element_param, dict):
            element = self.find_element(element_param)
        else:
            element = element_param
        element.click()

    def send_keys_element(self, element_param, text):
        if isinstance(element_param, dict):
            element = self.find_element(element_param)
        else:
            element = element_param
        element.send_keys(text)
        
    def action_decider(self, action, v, elements, stored_data, iterate=None):
        if action == 'get':
            url = iterate if iterate else elements[v]
            self.driver.get(url)
        elif action == 'click':
            el = iterate if iterate else elements[v['element']]
            self.click_element(el)
        elif action == 'send_keys':
            el = iterate if iterate else elements[v['element']]
            self.send_keys_element(el, v['text'])
        elif action == 'wait':
            element_params = elements[v['element']]
            self.wait_on_element(element_params['prop'], 
                                    by=element_params['by'], 
                                    wait_type=v['wait_type'],
                                    wait_time=v['wait_time'])
        elif action == 'sleep':
            time.sleep(v['time'])
        elif action == 'word_count':
            element_params = elements[v['element']]
            count = self.article_path_word_count(element_params['prop'], by=element_params['by'])
            key = iterate if iterate else v['store']
            if not stored_data['word_count'].get(key):
                stored_data['word_count'][key] = count
            else:
                stored_data['word_count'][key] += count
        elif action == 'store':
            key = v['key']
            nested_elements = elements[v['element']]
            temp_list = self.get_nested_elements(nested_elements)
            if not stored_data.get(key):
                stored_data[key] = temp_list
            else:
                stored_data[key].extend(temp_list)
        elif action == 'while':
            while True:
                try:
                    for a in v:
                        act = list(a.keys())[0]
                        val = list(a.values())[0]
                        self.action_decider(act, val, elements, stored_data)
                except Exception as e:
                    print(e)
                    break
        elif action == 'for':
            key = v['iterate']
            if not stored_data.get(key):
                raise Exception('{} is not yet stored'.format(key))
            for x in stored_data[key]:
                for a in v['action']:
                    act = list(a.keys())[0]
                    val = list(a.values())[0]
                    iterate = None
                    if isinstance(val, str) and key == val:
                        iterate = x
                    if isinstance(val, dict) and key in val.values():
                        iterate = x
                    try:
                        self.action_decider(act, val, elements, stored_data, iterate=iterate)
                    except:
                        break
        else:
            raise Exception('Invalid action')

    def execute_yaml_instructions(self, yml):
        # load the yml
        with open(yml) as file:
            conf = yaml.load(file)
        elements = conf['elements']
        instructions = conf['instructions']
        
        stored_data = {
            'word_count': {}
        }
        for instruction in instructions:
            action, v = list(instruction.items())[0]
            self.action_decider(action, v, elements, stored_data)

        return stored_data