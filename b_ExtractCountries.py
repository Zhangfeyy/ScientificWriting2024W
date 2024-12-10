import jieba
import pandas as pd
import re
import xlwt

# 1. get counyty list
ex_country_list = pd.read_excel('CountryList.xls')
country_list = ex_country_list.iloc[:, 0].tolist()

# 2. tokenization(word segmentation) & data analysis
ex_content_list = pd.read_excel('ChinaNewsIntList.xls', sheet_name=1)
content_list = ex_content_list.iloc[:, 5].tolist()

# 2.1 create the csv file
excl = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = excl.add_sheet('RelatedCountryList', cell_overwrite_ok=True)
sheet.write(0, 0, 'related country')
n = 1  # csv driver

# 2.2 write the data
for content in content_list:
    # Only save Chinese characters
    content = "".join(re.findall('[\u4e00-\u9fa5]+', content, re.S))
    # segmentation
    content = jieba.lcut(content)
    # extract country
    related_country = ""
    for word in country_list:
        related_country
        if word in content:
            related_country = related_country + word + ','
    # write it into the csv file
    sheet.write(n, 0, related_country)
    n += 1
    print('success')

# 2.3 save the csv file
excl.save("RelatedCountryList.xls")