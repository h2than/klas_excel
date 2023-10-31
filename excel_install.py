import os
from selenium import webdriver
import chromedriver_autoinstaller


def connect_chrome_session(self):
    try :
        chromedriver_autoinstaller.install()
    except:
        pass
    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://klas.jeonju.go.kr/klas3/LoanReturn/NewLoanReturnPage")

if __name__ == '__main__' :
    try :
        chromedriver_autoinstaller.install()
    except:
        pass
    profile = find_chrome_profile()
    print(profile)
    # connect_chrome_session()
