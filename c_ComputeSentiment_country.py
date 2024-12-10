import pandas as pd

times_country_dict = {}
score_country_dict = {}
network_country_list = []

# 1. single-article related countries list & Full countries list
related_country_list = pd.read_excel('ChinaNewsIntList.xls', sheet_name=2).iloc[:, 4].tolist()
organized_country_list = []
organized_countries = {}
for countries in related_country_list:
    network_countries = []
    if ',' in countries:
        tempo_list = countries.split(',')
        for item in tempo_list:
            if item != '':
                if item not in times_country_dict.keys():
                    times_country_dict[item] = 1
                else:
                    times_country_dict[item] += 1
                organized_countries[item] = 0
                network_countries.append(item)
    else:
        organized_countries[countries] = 0
        if countries not in times_country_dict.keys():
            times_country_dict[countries] = 1
        else:
            times_country_dict[countries] += 1
        network_countries.append(item)
    organized_country_list.append(organized_countries)
    network_country_list.append(network_countries)

    organized_countries = {}

# 3.assign scores
index = 0
sentiment_scores = pd.read_excel('ChinaNewsIntList.xls', sheet_name=2).iloc[:, 5].tolist()
for countries in organized_country_list:
    for country in countries:
        countries[country] = sentiment_scores[index]
    index += 1

# 4.calculate the score of each country
score_country_dict = times_country_dict.copy()
for country_score in score_country_dict:
    score_country_dict[country_score] = 0
for countries in organized_country_list:
    for country_organized in countries:
        score_country_dict[country_organized] += countries[country_organized]

with open('Country.py', mode='w', encoding='utf-8') as file:
    file.write(f"score_country_dict_overall = {score_country_dict}\n")
    for country_score_overall in score_country_dict:
        score_country_dict[country_score_overall] = score_country_dict[country_score_overall] / times_country_dict[
            country_score_overall]
    file.write(f"score_country_dict = {score_country_dict}\n")
    file.write(f"times_country_dict = {times_country_dict}\n")
    file.write(f"network_country_list = {network_country_list}\n")
