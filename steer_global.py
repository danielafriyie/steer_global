import warnings
import time
from threading import Thread, Lock, RLock
import logging
import random as rd
import os
import re

from selenium import webdriver
from selenium_stealth import stealth
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.firefox.service import Service as FoxService
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    TimeoutException, WebDriverException, InvalidSessionIdException, ElementClickInterceptedException,
    NoSuchElementException
)
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver as uc
from seleniumwire import webdriver as sw
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DRIVER_PATH = os.path.join(BASE_DIR, 'chromedriver')
ADS_BLOCKER_PATH = os.path.join(BASE_DIR, 'ads_blocker.crx')
GOST_PATH = os.path.join(BASE_DIR, 'ghost.crx')
RUNNING_CLARITY_BLOCKER = True
warnings.filterwarnings('ignore')


def get_driver():
    options = sw.ChromeOptions()
    service = Service(DRIVER_PATH)
    # options.add_argument('--headless')
    options.add_argument('--start-maximized')
    options.add_experimental_option("excludeSwitches", ["enable-automation", 'enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)
    chrome_driver = sw.Chrome(options=options)
    return chrome_driver


def block_clarity(driver: webdriver.Chrome):
    print("Blocker started")
    while RUNNING_CLARITY_BLOCKER:
        try:
            elms = driver_wait(driver, "//script[@async and contains(@src, 'clarity.ms')]", secs=0.01,
                               condition=EC.presence_of_all_elements_located)
            for e in elms:
                print("Blocking clarity scripts")
                try:
                    driver.execute_script(f"arguments[0].src=''", e)
                except WebDriverException:
                    pass
        except WebDriverException:
            continue
    print("Clarity blocker closing...")


def window_scroll_to(driver, loc):
    driver.execute_script(f"window.scrollTo(0, {loc});")


def scroll_into_view(driver, element, offset=200):
    window_scroll_to(driver, element.location['y'] - offset)


def driver_wait(driver, xpath, secs=10, condition=EC.element_to_be_clickable, action=None):
    wait = WebDriverWait(driver=driver, timeout=secs)
    element = wait.until(condition((By.XPATH, xpath)))
    if action:
        if hasattr(element, action):
            action_func = getattr(element, action)
            action_func()
    return element


def driver_or_js_click(driver, xpath, secs=5, condition=EC.element_to_be_clickable):
    try:
        driver_wait(driver, xpath, action='click', secs=secs, condition=condition)
    except (TimeoutException, ElementClickInterceptedException):
        elm = driver.find_element_by_xpath(xpath)
        driver.execute_script("arguments[0].click()", elm)


def manual_entry(driver, xpath, text, secs, *args, **kwargs):
    elm = driver_wait(driver, xpath, secs=secs)
    try:
        elm.clear()
    except WebDriverException:
        pass
    text = str(text)
    for letter in text:
        elm.send_keys(letter)
        time.sleep(0.05)
    elm.send_keys(*args, **kwargs)


def logger(name: str = None, fmt: str = None, filename: str = None):
    _logger = Logger(name, fmt, filename)
    return _logger()


class Logger:
    """
    Base Logger class
    """
    __loggers = {}

    def __init__(self, name: str = None, fmt: str = None, filename: str = None):
        self.name = name if name else __name__
        self.fmt = fmt if fmt else '%(asctime)s:%(levelname)s:%(message)s'
        self._root_path = BASE_DIR
        self.filename = filename if filename else self._get_log_file()

    def _get_log_file(self):
        log_path = os.path.join(self._root_path, 'logs/log.log')
        if not os.path.isfile(log_path):
            os.mkdir(os.path.join(self._root_path, 'logs'))
        return log_path

    def _create_logger(self):
        _logger = logging.getLogger(self.name)
        _logger.setLevel(level=logging.DEBUG)

        formatter = logging.Formatter(self.fmt)
        file_handler = logging.FileHandler(self.filename)
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)

        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        stream_handler.setLevel(logging.INFO)

        _logger.addHandler(file_handler)
        _logger.addHandler(stream_handler)

        return _logger

    def _log_file_manager(self):
        """
        check if the log file size is more than 10mb then deletes it
        """
        raise NotImplementedError(f'{self.__class__.__name__}._log_file_manager() method is not implemented!')

    def __call__(self):
        if self.name in self.__loggers:
            return self.__loggers.get(self.name)
        else:
            _logger = self._create_logger()
            self.__loggers[self.name] = _logger
            return _logger


class WorkFlow(Thread):

    def __init__(self):
        super().__init__()
        self._queue = []
        self._idx = 0
        self._running = True
        self._driver = get_driver()
        # Thread(target=block_clarity, args=(self._driver,)).start()
        self._log = logger()

    def log(self, msg, method='info'):
        log_method = getattr(self._log, method)
        log_method(f" {self.name}: {msg}")
        self.log_status_code()

    def log_status_code(self):
        for r in self._driver.requests:
            try:
                self._log.info(f"{self.name}: STATUS CODE - {r.response.status_code}")
            except AttributeError:
                pass

    def log_with_username(self, username, msg):
        self.log(f"{username} {msg}")

    def submit_task(self, task):
        self._queue.append(task)

    def quit(self):
        global RUNNING_CLARITY_BLOCKER
        if self.is_alive():
            self._running = False
            RUNNING_CLARITY_BLOCKER = False
            self._driver.quit()

    def login(self, username, password):
        try:
            driver_or_js_click(self._driver, "//div[text()='Logout']", secs=2)
        except WebDriverException:
            pass
        try:
            for i in range(3):
                self.log(f'Authenticating with username ({username}) attempt {i + 1}')
                driver_or_js_click(self._driver, "//div[@id='pupil-login-div']/div", secs=25)
                manual_entry(self._driver, "//input[@id='pupil-username']", username.strip(), 25)
                manual_entry(self._driver, "//input[@id='pupil-password']", password.strip(), 25)
                driver_or_js_click(self._driver, "//form[@id='pupil-login']/input[@type='submit']", secs=25)
                try:
                    err = driver_wait(self._driver, "//div[@class='new-login-error-msg-pupil']", secs=10,
                                      condition=EC.presence_of_element_located)
                    self.log(err.text)
                    self._driver.refresh()
                    continue
                except TimeoutException:
                    self.log(f"({username}) authenticated successfully!")
                    return True
        except WebDriverException:
            pass
        self.log(f"Couldn't authenticate {username}!")

    def get_question_count(self, username, resume=False):
        self.log("Getting current question number.")
        try:
            driver_or_js_click(self._driver, "//div[@class='reminder']/div/img", 5)
            time.sleep(2)
            driver_or_js_click(self._driver, "//div[@class='reminder']/div/img", 10)
            driver_or_js_click(self._driver, "//div[@class='intro_statement']/h1", 10)
            voice = rd.choice(['male-voice', 'female-voice'])
            driver_or_js_click(self._driver, f"//div[@class='{voice}']", 10)
        except WebDriverException:
            self.log_status_code()
            pass
        try:
            driver_wait(self._driver, "//span[@ng-click='playAudio()']", 5)
            self._driver.execute_script('document.getElementById("audio_statement").children[0].children[0].click()')
            # self.log_with_username(username, "Playing audio")
            self.log_status_code()
            resume = False
        except WebDriverException:
            self.log_status_code()
            pass
        for i in range(16):
            try:
                if resume:
                    self.log_with_username(username, "Resuming")
                    try:
                        qs_count = driver_wait(self._driver, f"(//div[@class='questionCount'])[{i + 1}]",
                                               condition=EC.presence_of_element_located, secs=10).text
                    except TimeoutException:
                        try:
                            qs_count = self._driver.find_element(By.XPATH,
                                                                 f"(//div[@class='questionCount'])[{i + 1}]").text
                        except NoSuchElementException:
                            continue
                else:
                    self.log_with_username(username, "Starting over")
                    resume = True
                    try:
                        qs_count = driver_wait(self._driver, f"(//div[@class='questionCount'])[{i + 1}]",
                                               condition=EC.presence_of_element_located, secs=480).text
                    except TimeoutException:
                        try:
                            qs_count = self._driver.find_element(By.XPATH,
                                                                 f"(//div[@class='questionCount'])[{i + 1}]").text
                        except NoSuchElementException:
                            continue
                qs_split = qs_count.split('/')
                questions_count = int(''.join(re.findall(r'[0-9]+', qs_split[-1])))
                current_question = int(''.join(re.findall(r'[0-9]+', qs_split[0])))
                return questions_count, current_question
            except ValueError:
                time.sleep(1)
                continue
        return None, None

    def answer(self, username, questions_count, current_question):
        self.log(f'Questions count {questions_count}')
        self.log(f"Current question for username - ({username})")
        actions = ActionChains(self._driver)
        for n in range(current_question, questions_count + 1):
            self.log_with_username(username, f"Attempting question {n}")
            xpath = f"(//div[@class='options_div'])[{n}]"
            options_div = driver_wait(self._driver, xpath, secs=25,
                                      condition=EC.visibility_of_element_located)
            options = options_div.find_elements(By.XPATH, ".//div")
            len_options = len(options)
            option_n = rd.randint(1, (len_options - 2))
            self.log(f"{username} selects option {option_n}")
            option = options_div.find_element(By.XPATH, f"(.//div)[{option_n}]")
            actions.move_to_element(option)
            time.sleep(1)
            actions.click()
            actions.perform()
            self.log_with_username(username, f"Done with question number: {n}")
            time.sleep(3)
            actions.reset_actions()

    def answer_questions(self, username, resume=False):
        start_time = time.time()
        self._driver.switch_to.default_content()
        self._driver.switch_to.frame(
            driver_wait(self._driver, "//iframe[@id='iframe']", secs=60, condition=EC.presence_of_element_located))
        time.sleep(1)
        questions_count, current_question = self.get_question_count(username, True)
        if questions_count is None and current_question is None:
            self.log(f"{username} couldn't get question count")
            return round((time.time() - start_time) / 60, 2)
        self.answer(username, questions_count, current_question)
        # for i in range(5):
        #     try:
        #         driver_or_js_click(self._driver, f"//span[@ng-click='getAssessment({i if i > 0 else ''})']", secs=2)
        #         self._driver.switch_to.default_content()
        #         return self.answer_questions(username, False)
        #     except (TimeoutException, NoSuchElementException):
        #         continue
        # driver_wait(self._driver, "//h1[contains(text(), 'You have now completed all parts of the assessment')]",
        #             secs=25, condition=EC.visibility_of_element_located)
        # driver_or_js_click(self._driver, "//span[@ng-click='getAssessment(1)']")
        time.sleep(1.5)
        self._driver.switch_to.default_content()
        self.log_with_username(username, "Done answering questions")
        end_time = (time.time() - start_time) / 60
        # try:
        #     text = driver_wait(self._driver,
        #                        "//strong[contains(text(), 'You have already completed the assessment today')]",
        #                        secs=15, condition=EC.visibility_of_element_located).text
        #     self.log(f"{username}, {text}")
        # except WebDriverException:
        #     pass
        return round(end_time, 2)

    def do_job(self, row):
        url = row['Login URL']
        sch_code = row['School Code']
        mis_id = row['MIS ID']
        fn = row['First Name']
        username = row['Username']
        password = row['Password']
        year = row['Year']
        gender = row['Gender']
        date = row['Date']
        house = row['House']
        form = row['Form']
        campus = row['Campus']
        self._driver.get(url)
        self.log_status_code()
        is_authenticated = self.login(username, password)
        if not is_authenticated:
            return
        try:
            driver_or_js_click(self._driver, "//strong[contains(text(), 'Click here to begin your new AS')]", 5)
        except (TimeoutException, NoSuchElementException):
            try:
                driver_or_js_click(self._driver, "//strong[contains(text(), 'Click here to continue your AS')]", 5)
                self.log_with_username(username, "Resuming answering questions")
                time_taken = self.answer_questions(username)
                self.log(f"Total time taken for ({username}) to answer questions {time_taken} minutes")
                return
            except (TimeoutException, NoSuchElementException):
                text = driver_wait(self._driver,
                                   "//strong[contains(text(), 'You have already completed the assessment today')]",
                                   secs=5, condition=EC.visibility_of_element_located).text
                self.log(f"{username}, {text}")
                try:
                    driver_or_js_click(self._driver, "//div[text()='Logout']", secs=5)
                except WebDriverException:
                    pass
                return
        self.log_with_username(username, "Answering questions from start")
        self._driver.switch_to.frame(
            driver_wait(self._driver, "//iframe[@id='iframe']", secs=60, condition=EC.presence_of_element_located))
        driver_or_js_click(self._driver, "//div[@class='reminder']/div/img", 60)
        time.sleep(2)
        driver_or_js_click(self._driver, "//div[@class='reminder']/div/img", 10)
        driver_or_js_click(self._driver, "//div[@class='intro_statement']/h1", 10)
        voice = rd.choice(['male-voice', 'female-voice'])
        driver_or_js_click(self._driver, f"//div[@class='{voice}']", 10)
        driver_or_js_click(self._driver, "//span[@ng-click='playAudio()']", 10)
        self._driver.switch_to.default_content()
        time_taken = self.answer_questions(username)
        self.log(f"Total time taken for ({username}) to answer questions {time_taken} minutes")

    def run(self):
        while self._running:
            try:
                row = self._queue.pop(0)
                self.do_job(row)
            except IndexError:
                self.log('Queue list exhausted!')
                break
            except Exception as e:
                self.log_status_code()
                pass
        self.quit()
        self.log(f"\n\n{'=' * 80}\n\n\t\t\t\t\tAll process done!\n\n{'=' * 80}")


if __name__ == '__main__':
    N_THREADS = 3
    df = pd.read_excel('data.xlsx', engine='openpyxl')
    len_df = len(df)
    counter = 0
    workers = []

    for _ in range(N_THREADS):
        worker = WorkFlow()
        workers.append(worker)

    while counter <= len_df:
        try:
            for worker in workers:
                worker.submit_task(df.iloc[counter])
                counter += 1
        except IndexError:
            break

    for worker in workers:
        worker.start()

    for worker in workers:
        worker.join()
