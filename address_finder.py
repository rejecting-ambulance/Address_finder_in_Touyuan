import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook

def search_address(driver, wait, address):
    driver.get('https://addressrs.moi.gov.tw/address/index.cfm?city_id=68000')
    address_box = wait.until(EC.presence_of_element_located((By.ID, 'FreeText_ADDR')))
    submit_button = driver.find_element(By.ID, 'ext-comp-1010')

    address_box.clear()
    address_box.send_keys(address)
    submit_button.click()

    wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ext-gen107"]/div[1]/table/tbody/tr/td[2]/div'))
    )

    try:
        result = driver.find_element(By.XPATH, '//*[@id="ext-gen107"]/div/table/tbody/tr/td[2]/div')
        return result.text.strip()
    except:
        return "找不到結果"

def simplify_address(address):
    # 你要找的斷點字元
    split_chars = ['號', '及', '、', '.']
    
    # 找最早出現的位置
    split_indices = [(address.find(char), char) for char in split_chars if address.find(char) != -1]
    
    if split_indices:
        # 按位置排序，找最前面的字
        split_indices.sort()
        index, char = split_indices[0]
        
        if char == '號':
            short_address = address[:index + 1]  # 不包含「號」
            rest_address = address[index + 1:]
        else:
            short_address = address[:index ]  # 包含那個字
            rest_address = address[index:]   # 後面剩下的
    else:
        short_address = address
        rest_address = ''
    
    return short_address, rest_address

def fullwidth_to_halfwidth(text):
    half_text = ''
    for char in text:
        code = ord(char)
        if code == 0x3000:
            code = 0x0020
        elif 0xFF01 <= code <= 0xFF5E:
            code -= 0xFEE0
        half_text += chr(code)
    return half_text

def remove_ling_with_condition(full_address):
    if "高上里" in full_address:
        # 是高上里 → 不刪除，直接回傳原本
        return full_address
    else:
        # 不是高上里 → 刪掉「三位數字+鄰」
        return re.sub(r'\d{3}鄰', '', full_address)



if __name__ == '__main__':
    file_path = 'address_data.xlsx'

    df = pd.read_excel(file_path)
    addresses = df['查詢地址'].tolist()
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)

    full_addresses = []
    simplified_addresses = []

    for address in addresses:
        try:
            shorter_address, last_address = simplify_address(address)
            result_address = search_address(driver, wait, shorter_address)
            result_address = fullwidth_to_halfwidth(result_address)
            full_address = f'桃園市{result_address}{last_address}'
            print(f"{address} → {full_address}")

            full_addresses.append(full_address)
            simplified_addresses.append(remove_ling_with_condition(full_address))
        except Exception as e:
            print(f"{address} 查詢失敗：{e}")
            full_addresses.append("查詢失敗")
            simplified_addresses.append("查詢失敗")

    driver.quit()

    df['完整地址'] = full_addresses
    df['不含鄰的地址'] = simplified_addresses
    df.to_excel(file_path, index=False)
    print(f"已完成，請查看 {file_path}")
