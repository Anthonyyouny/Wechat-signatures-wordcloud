import xlrd
import jieba
import matplotlib.pyplot as plt
from wordcloud import WordCloud,STOPWORDS 

signatures = []
workbook = xlrd.open_workbook('Friends_in_wechat.xlsx')

sheet = workbook.sheet_by_index(1)

signatures = str(sheet.col_values(6))

#with open('friends.txt', mode='r', encoding='utf-8') as f:
#    rows = f.readlines()
#    for row in rows:
#        signature = row
#        if signature != '':
#            signatures.append(signature)

split = jieba.cut(str(signatures),cut_all = False)
words = ' '.join(split)
print(words)

stopwords = STOPWORDS.copy()
stopwords.add('span')
stopwords.add('class')
stopwords.add('emoji')
stopwords.add('emoji1f334')
stopwords.add('emoji270c')
stopwords.add('emoji1f645')
stopwords.add('emoji1f44a')
stopwords.add('emoji2764')

bg_image = plt.imread('bg2.jpg')

wc = WordCloud(background_color='white', mask=bg_image, font_path='msyh.ttc',stopwords=stopwords,
               max_font_size=400, random_state=50)

wc.generate_from_text(words)
plt.imshow(wc) 
plt.axis('off')

wc.to_file('sign.jpg')





