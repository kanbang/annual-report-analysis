'''
Descripttion: 
version: 0.x
Author: zhai
Date: 2022-01-04 15:49:31
LastEditors: zhai
LastEditTime: 2022-01-05 13:35:26
'''

import re                           # 正则表达式库
import os  
import docx
from cntext.stats import term_freq, readability
from cntext.sentiment import senti_by_hownet, senti_by_dutir
from cntext import wordcloud
import pyecharts.options as opts
from pyecharts.charts import Radar
import random
from pyecharts.charts import Bar
import names

# 个人总结
summarys = {}
# 文字汇总
text_all = ""

# 读取word文档
def readdoc(file_path):
    doc=docx.Document(file_path)
    text = ""
    for para in doc.paragraphs:
        # print(para.text)
        text += para.text
    return text
    
# 读取文件夹
def readdir(path):  
    for file in os.listdir(path):  
        file_path = os.path.join(path, file)  
        if os.path.splitext(file_path)[1]=='.docx':  
            # name = re.findall(r'[^()]+', file)[1] # 通过文件提取名字
            name = names.get_first_name() # 名字脱敏
            global text_all
            global summarys
            text = readdoc(file_path)
            text_all += text
            summarys[name] = text


readdir("E:\\pj\\python\\summary\\doc")

# 文本预处理
# pattern = re.compile(u'\t|\n|\.|-|:|;|\)|\(|\?|"') # 定义正则表达式匹配模式（空格等）
# text_all = re.sub(pattern, '', text_all)     # 将符合模式的字符去除

# 词频统计
print(term_freq(text_all))

# 使用大连理工大学情感本体库对文本进行情绪分析，统计各情绪词语出现次数
senti1 = senti_by_dutir(text_all)
print(senti1)

# 使用知网Hownet词典进行(中)文本数据的情感分析，统计正、负情感信息出现次数(得分)
senti2 = senti_by_hownet(text_all)
print(senti2)

# 输出词云图
wordcloud(text=text_all, 
          title='词云图', 
          html_path='output/词云图.html')


# 输出情绪雷达图
rad = Radar(init_opts=opts.InitOpts(width="1280px", height="720px", bg_color="#CCCCCC"))
rad.add_schema(
        schema=[
            opts.RadarIndicatorItem(name="好", max_=100),
            opts.RadarIndicatorItem(name="乐", max_=100),
            opts.RadarIndicatorItem(name="哀", max_=100),
            opts.RadarIndicatorItem(name="怒", max_=100),
            opts.RadarIndicatorItem(name="惧", max_=100),
            opts.RadarIndicatorItem(name="恶", max_=100),
            opts.RadarIndicatorItem(name="惊", max_=100),
        ],
        splitarea_opt=opts.SplitAreaOpts(
            is_show=True, areastyle_opts=opts.AreaStyleOpts(opacity=1)
        ),
        textstyle_opts=opts.TextStyleOpts(color="#fff"),
    )


def random_color():
     colors = '0123456789ABCDEF'
     clr = "#"
     for i in range(6):
         clr += random.choice(colors)
     return clr

for name in summarys.keys() :
    # print(name , summarys[name])
    senti = senti_by_dutir(summarys[name])
    sen = [list(senti.values())[3:]]
   
    rad.add(
            series_name=name,
            data=sen,
            linestyle_opts=opts.LineStyleOpts(color=random_color()),
        )
    

rad.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
rad.set_global_opts(
        title_opts=opts.TitleOpts(title="情绪雷达图"), legend_opts=opts.LegendOpts()
    )
rad.render("output/情绪雷达图.html")


# 输出情绪堆叠柱状图

pos = []
neg = []
for name in summarys.keys() :
    senti = senti_by_hownet(summarys[name])
    pos.append(senti['pos_word_num'])
    neg.append(-senti['neg_word_num'])

bar = Bar()
bar.add_xaxis(list(summarys.keys()))

bar.add_yaxis("pos", pos, stack="senti")
bar.add_yaxis("neg", neg, stack="senti")
bar.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
bar.set_global_opts(title_opts=opts.TitleOpts(title="情绪堆叠数据"))
bar.render("output/情绪堆叠柱状图.html")

