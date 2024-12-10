import numpy as np
import pandas as pd
import xlwt
import seaborn as sb
import matplotlib.pyplot as plt


def statistics(list):
    mean = np.mean(list)
    std = np.std(list)
    print(mean, std)
    return (list - mean) / std


dict_sentiment = pd.read_excel('F:\Python2024\ScientificWriting\Excel\SentimentScores1.xls')
dict_sentiment_list = dict_sentiment.iloc[0:, 0].tolist()
z_dict_sentiment = statistics(dict_sentiment_list)
print(np.mean(z_dict_sentiment), np.std(z_dict_sentiment))

bayes_sentiment = pd.read_excel('F:\Python2024\ScientificWriting\Excel\SentimentScores3.xls')
bayes_sentiment_list = bayes_sentiment.iloc[0:, 0].tolist()
z_bayes_sentiment = statistics(bayes_sentiment_list)
print(np.mean(z_bayes_sentiment), np.std(z_bayes_sentiment))

# Results
# 1.194743374691192 1.707695919034323
# 4.603948611836502e-17 1.0
# 0.7276430688283081 0.20996515773498253
# 1.5346495372788342e-17 1.0

# excl = xlwt.Workbook(encoding='utf-8', style_compression=0)
# sheet = excl.add_sheet('SentimentScores', cell_overwrite_ok=True)
# sheet.write(0, 0, 'dict')
# sheet.write(0, 1, 'bayes')
# n = 1  # csv driver
#
# for number in z_dict_sentiment:
#     sheet.write(n, 0, number)
#     n += 1
#
# n = 1
#
# for number in z_bayes_sentiment:
#     sheet.write(n, 1, number)
#     n += 1
#
# excl.save("SentimentScores2.xls")


# draw the distrinution function
sb.distplot(z_dict_sentiment,bins=20,kde_kws={"color":"blue","linestyle":"-"},hist_kws={"color":"steelblue"},label="dict")
sb.distplot(z_bayes_sentiment,bins=20,kde_kws={"color":"purple","linestyle":"-"},hist_kws={"color":"pink"},label="bayes")

plt.title('Distribution of the sentiment scores under different algorithms')
plt.legend()
plt.show()