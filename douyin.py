from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions


def douyin():
    driver = webdriver.Chrome()
    driver.get('https://www.douyin.com/video/7286653740084514104')

    loginXpath = '//*[@id="login-pannel"]/div[2]'
    WebDriverWait(driver, 10, 0.5).until(expected_conditions.element_to_be_clickable((By.XPATH, loginXpath)))
    driver.find_element(By.XPATH, loginXpath).click()

    commentXpath = '//*[@id="douyin-right-container"]/div[2]/div/div[1]/div[5]/div/div/div[3]/div[position()<4]'
    commentXpath .= '/div/div[2]/div/div[2]/span/span/span/span'
    WebDriverWait(driver, 10, 0.5).until(expected_conditions.visibility_of(driver.find_element(By.XPATH, commentXpath)))
    comments = driver.find_elements(By.XPATH, commentXpath)

    for comment in comments:
        print(comment.text)


if __name__ == '__main__':
    douyin()
