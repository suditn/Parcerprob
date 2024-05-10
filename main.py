from selenium import webdriver
import re
import pandas as pd
from bs4 import BeautifulSoup
import shutil
from io import StringIO
import os
from pathlib import Path
import requests
import time

pattern = re.compile(r'\w+')
# driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()))

driver = webdriver.Edge()
url = 'https://www.vishay.com/en/inductors/'
save_patch = str(Path(__file__).parent.resolve())
img_small_save_patch = save_patch+"\\image\\small_inductors\\"
datash_save_patch = save_patch+"\\Datasheet\\"
headers = {'User-Agent': "scrapping_script/1.0"}

def get_web(u):
    driver.get(u)
    print(save_patch)
    Path(img_small_save_patch).mkdir(parents=True, exist_ok=True)
    Path(datash_save_patch).mkdir(parents=True, exist_ok=True)
    option2 = driver.find_element('xpath', '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/div[2]/label/select/option[1]')
    max_ent = driver.find_element('xpath', '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div').text
# //*[@id="Table_vshGenTblPaginationInfo__29nYK"]/div/text()[6]
    driver.execute_script('arguments[0].value = arguments[1]',option2 , pattern.findall(max_ent)[5])
# print(option)
    option2.click()
    time.sleep(3)
    webso = driver.page_source
    table = []
#    df = pd.DataFrame()
#    df_datasheet = pd.DataFrame(['Datasheet'])
#    df = pd.read_html(StringIO(webso))
#    print(df[3])

    soup = BeautifulSoup(webso, "lxml")
    table = soup.find('table', {'id': 'poc'})
    img = table.findAll('img')
    img_src = []
    img_alt = []
    img_pr = ''
    img_pr2 =''
    datash_src= []
    datash_pr = ''
    columns = [i.get_text(strip=True) for i in table.find_all("th")]
    data = []

    for tr in table.find("tbody").find_all("tr"):
        data.append([td.get_text(strip=True) for td in tr.find_all("td")])

    df = pd.DataFrame(data, columns=columns, )
    i = 0
    for im in img:

        ser = df['Series▲▼'][i]
        if (im['src'].split('/')[-2]=='pt-small'):
            img_src.append(img_small_save_patch+im['alt']+'.png')
            img_alt.append(im['alt'])
            if(img_pr!=im['src'] and im['alt']!="Datasheet"):
                img_requ = requests.get('https://www.vishay.com/'+im['src'], stream=True)
                if os.path.exists(img_small_save_patch+im['alt']+'.png' or img_small_save_patch+im['alt']+'.jpg'):
                    print('файл '+im['alt']+' существует')
                else:
                    with open(img_small_save_patch+im['alt']+'.png', 'wb') as out_file:
                        shutil.copyfileobj(img_requ.raw, out_file)
                    print('файл '+im['alt']+'.png создан')

                img_pr = im['src']
                del img_requ
            datash_src.append(datash_save_patch + ser + '\\' + ser + '.pdf')
            i += 1



        if(im['alt'] != "Datasheet" and datash_pr != ser):
            datash_requ = requests.get('https://www.vishay.com/doc?'+im['alt'],headers=headers, stream = True)
            Path(datash_save_patch+'\\'+ser).mkdir(parents=True, exist_ok=True)
            print(i)
            print(datash_requ.url)
            if os.path.exists(datash_save_patch + '\\' + ser+'\\'+ser+'.pdf'):
                print('файл '+ser+' существует')
            else:
                with open(datash_save_patch + ser + '\\' + ser+'.pdf', 'wb') as out_file:
                    out_file.write(datash_requ.content)
                print('файл ' + ser + ' создан')

            del datash_requ

            datash_pr = ser
















    writer = pd.ExcelWriter(save_patch + '/' + url.split('/')[-2] + '.xlsx', engine='xlsxwriter')
    df_img = pd.DataFrame(img_src)
    df_datasheet = pd.DataFrame(datash_src, columns=['Datasheet'])
    df['Product Image'] = df_img

    df.join(df_datasheet).to_excel(writer, index=False, sheet_name='Inductors')

    worksheet = writer.sheets['Inductors']
    worksheet.autofit()



    writer.close()



#    return "Страница открыта"

print(get_web(url))
