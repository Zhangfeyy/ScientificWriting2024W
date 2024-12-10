import xlwt
import pandas as pd
import jieba
import re
import snownlp

from sentimentanalysis.DegreeDict import degree_dict
from sentimentanalysis.NegationList import negation_list
from sentimentanalysis.PositiveLexi import positive_dict
from sentimentanalysis.NegativeLexi import negative_dict


# 0. function
def cut_sent(para):
    para = re.sub('([。！？\?])([^”’])', r"\1\n\2", para)  # 。！？
    para = re.sub('(\.{6})([^”’])', r"\1\n\2", para)  # ...
    para = re.sub('(\…{2})([^”’])', r"\1\n\2", para)  # 。。。
    para = re.sub('([。！？\?][”’])([^，。！？\?])', r'\1\n\2', para)
    para = para.rstrip()
    return para.split("\n")
    # refer to CSDN


def sentiment_score(sentence):
    # record the pos of the scanned word
    i = 0
    # record the pos of the (former) sentiment word
    a = 0
    # the first score of positive word
    poscount = 0
    # the reversed score of positive word
    poscount2 = 0
    # the final score of positive word
    poscount3 = 0
    # the first score of negative word
    negcount = 0
    # the reversed score of negative word
    negcount2 = 0
    # the final score of negative word
    negcount3 = 0
    word_count = 0
    for word in sentence:
        if word in positive_dict.keys():
            # positive sentiment analyse
            poscount += int(positive_dict[word])
            deny_count_positive = 0
            # check if the front word is degree word
            for front_word in sentence[a:i]:
                if front_word in degree_dict.keys():
                    poscount *= int(degree_dict[front_word])
                    # check if the front word is to deny
                elif front_word in negation_list:
                    deny_count_positive += 1
            # check if the count is even (to cancel the effect of negation)
            if deny_count_positive % 2 != 0:
                poscount *= -1.0
                poscount2 += poscount
                poscount = 0
                # + poscount
                poscount3 = + poscount2 + poscount3
                poscount2 = 0
            else:
                # + poscount2 = 0
                poscount3 = poscount + poscount3
                poscount = 0
            a = i + 1
            word_count += 1

        elif word in negative_dict.keys():
            # positive sentiment analyse
            negcount += int(negative_dict[word]) * (-1)
            deny_count_negative = 0
            # check if the front word is degree word
            for front_word in sentence[a:i]:
                if front_word in degree_dict.keys():
                    negcount *= int(degree_dict[front_word])
                    # check if the front word is to deny
                elif front_word in negation_list:
                    deny_count_negative += 1
            # check if the count is even (to cancel the effect of negation)
            if deny_count_negative % 2 != 0:
                negcount *= -1.0
                negcount2 += negcount
                negcount = 0
                # + negcount
                negcount3 = + negcount2 + negcount3
                negcount2 = 0
            else:
                # + poscount2 = 0
                negcount3 = negcount + negcount3
                negcount = 0
            a = i + 1
            word_count += 1
        i += 1
    if not word_count:
        return 0
    return ((poscount3 + negcount3)/word_count)


# 1. open the csv file
ex_content_list = pd.read_excel('ChinaNewsIntList.xls', sheet_name=2)
content_list = ex_content_list.iloc[0:, 5].tolist()

# 2. create stopwords list
file = open('sentimentanalysis/scu_stopwords.txt', encoding='utf-8')
stopword_list = []
for word in file.readline():
    stopword_list.append(word.strip())

# 3 create the csv file
excl = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = excl.add_sheet('SentimentScores', cell_overwrite_ok=True)
sheet.write(0, 0, 'sentiment score')
n = 1  # csv driver

# 4. analyse sentiment
for content in content_list:
    # sentence segmentation
    content_score = 0
    sentence_list = cut_sent(content)
    sentence_count = 0
    for sentence in sentence_list:
        # only keep chinese characters
        sentence = "".join(re.findall('[\u4e00-\u9fa5]+', sentence, re.S))
        if sentence != "":
            # sentence = jieba.lcut(sentence)
            # content_score += sentiment_score(sentence)
            content_score += snownlp.SnowNLP(sentence).sentiments
            sentence_count += 1
    sheet.write(n, 0, content_score / sentence_count)
    n += 1

# excl.save('SentimentScores1.xls')
excl.save('SentimentScores3.xls')
