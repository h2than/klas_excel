import os
from selenium import webdriver

def find_chrome_profile():
    user_data_path = os.path.expanduser("~")  # 현재 사용자 홈 디렉토리
    chrome_user_data_path = os.path.join(user_data_path, 'AppData', 'Local', 'Google', 'Chrome', 'User Data') 
    profile_name = os.listdir(chrome_user_data_path)[0] if os.path.exists(chrome_user_data_path) else "Default"    
    profile_path = os.path.join(chrome_user_data_path, profile_name)
    return profile_path

def connect_chrome_session(self):
    chrome_profile_path = self.find_chrome_profile()
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument(f'--user-data-dir={chrome_profile_path}')
    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://klas.jeonju.go.kr/klas3/Admin/")

if __name__ == '__main__' :
    profile = find_chrome_profile()
    print(profile)
    # connect_chrome_session()
