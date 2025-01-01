import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook

def read_keywords_from_excel(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df['Keyword'].tolist()

def write_results_to_excel(file_path, sheet_name, results):
    workbook = load_workbook(file_path)
    sheet = workbook[sheet_name]

    for i, result in enumerate(results, start=2):
        sheet[f'B{i}'] = result['Longest']
        sheet[f'C{i}'] = result['Shortest']

    workbook.save(file_path)

def get_search_suggestions(keyword, driver):
    driver.get('https://www.google.com')
    search_box = driver.find_element_by_name('q')
    search_box.send_keys(keyword)
    driver.implicitly_wait(3)

    suggestions = driver.find_elements_by_css_selector('.sbl1 span')
    if suggestions:
        suggestion_texts = [s.text for s in suggestions]
        longest = max(suggestion_texts, key=len)
        shortest = min(suggestion_texts, key=len)
        return {'Longest': longest, 'Shortest': shortest}
    else:
        return {'Longest': '', 'Shortest': ''}

def main():
    file_path = 'keywords.xlsx'
    sheet_name = 'Monday'

    keywords = read_keywords_from_excel(file_path, sheet_name)

    service = Service('path_to_chromedriver')
    driver = webdriver.Chrome(service=service)

    results = []
    for keyword in keywords:
        result = get_search_suggestions(keyword, driver)
        results.append(result)

    driver.quit()

    write_results_to_excel(file_path, sheet_name, results)

if __name__ == '__main__':
    main()
