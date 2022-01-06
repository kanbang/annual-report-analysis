'''
Descripttion: 
version: 0.x
Author: zhai
Date: 2022-01-04 15:49:31
LastEditors: zhai
LastEditTime: 2022-01-06 13:34:17
'''

from collections import defaultdict
import re                           # 正则表达式库
import os
from cntext.dictionary.dictionary import DUTIR_Ais, DUTIR_Haos, DUTIR_Jings, DUTIR_Jus, DUTIR_Les, DUTIR_Nus, DUTIR_Wus, STOPWORDS_zh  
import docx
import cntext
from cntext.stats import term_freq, readability
from cntext.sentiment import senti_by_hownet, senti_by_dutir
from cntext import wordcloud
import jieba
from pyecharts.charts.basic_charts.wordcloud import WordCloud
import pyecharts.options as opts
from pyecharts.charts import Radar
from pyecharts.charts import Pie
from pyecharts.charts import Bar
import random
import names
import pandas as pd


# help(cntext)

# 个人总结
summarys = {}

# 文字汇总
text_all = ""


def senti_by_dutir_detail(text):
    """
    使用大连理工大学情感本体库DUTIR，仅计算文本中各个情绪词出现次数
    :param text:  中文文本字符串
    :return: 返回文本情感统计信息，类似于这样{'words': 22, 'sentences': 2, '好': 0, '乐': 4, '哀': 0, '怒': 0, '惧': 0, '恶': 0, '惊': 0}
    """
    wordnum, sentences, hao, le, ai, nu, ju, wu, jing, stopwords =0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    sentences = len(re.split('[\.。！!？\?\n;；]+', text))
    words = jieba.lcut(text)
    wordnum = len(words)

    # 统计词频
    dict_hao = defaultdict(lambda: 0)
    dict_le = defaultdict(lambda: 0)
    dict_ai = defaultdict(lambda: 0)
    dict_nu = defaultdict(lambda: 0)
    dict_ju = defaultdict(lambda: 0)
    dict_wu = defaultdict(lambda: 0)
    dict_jing = defaultdict(lambda: 0)

    for w in words:
        if w in STOPWORDS_zh:
            stopwords+=1
        if w in DUTIR_Haos:
            hao += 1
            dict_hao[w] += 1
        elif w in DUTIR_Les:
            le += 1
            dict_le[w] += 1
        elif w in DUTIR_Ais:
            ai += 1
            dict_ai[w] += 1
        elif w in DUTIR_Nus:
            nu += 1
            dict_nu[w] += 1
        elif w in DUTIR_Jus:
            ju += 1
            dict_ju[w] += 1
        elif w in DUTIR_Wus:
            wu += 1
            dict_wu[w] += 1
        elif w in DUTIR_Jings:
            jing += 1
            dict_jing[w] += 1
        else:
            pass

    result = {'word_num':wordnum,
            'sentence_num':sentences,
            'stopword_num':stopwords,
            '好_num':hao, '乐_num':le, '哀_num':ai, '怒_num':nu, '惧_num':ju, '恶_num': wu, '惊_num':jing,
            '好_dict':dict_hao,
            '乐_dict':dict_le,
            '哀_dict':dict_ai,
            '怒_dict':dict_nu,
            '惧_dict':dict_ju,
            '恶_dict':dict_wu,
            '惊_dict':dict_jing
            }
    return result

def wordcloud_by_dict(title, html_path, wordfreq_dict):
    wordfreqs = [(word, str(freq)) for word, freq in wordfreq_dict.items()]
    wc = WordCloud()
    wc.add(series_name="", data_pair=wordfreqs, word_size_range=[20, 100])
    wc.set_global_opts(
        title_opts=opts.TitleOpts(title=title,
                                  title_textstyle_opts=opts.TextStyleOpts(font_size=23)
                                  ),
        tooltip_opts=opts.TooltipOpts(is_show=True))
    wc.render(html_path)  #存储位置

    
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


readdir("./doc")

# 文本预处理
# pattern = re.compile(u'\t|\n|\.|-|:|;|\)|\(|\?|"') # 定义正则表达式匹配模式（空格等）
# text_all = re.sub(pattern, '', text_all)     # 将符合模式的字符去除

# 词频统计表
freq = term_freq(text_all)
freq_sorted=sorted(freq.items(),key=lambda x:x[1],reverse=True)  # 按字典集合中，每一个元组的第二个元素排列。

keys =[key for key,value in freq_sorted]
values =[value for key,value in freq_sorted]

df = pd.DataFrame({'词': keys, '频': values}) 

df.to_html("./output/词频.html")


# 使用大连理工大学情感本体库对文本进行情绪分析，统计各情绪词语出现次数
senti1 = senti_by_dutir_detail(text_all)
print(senti1)

# 使用知网Hownet词典进行(中)文本数据的情感分析，统计正、负情感信息出现次数(得分)
senti2 = senti_by_hownet(text_all)
print(senti2)


# 饼图
inner_x_data = ["正", "负"]
inner_y_data = [senti2['pos_word_num'], senti2['neg_word_num']]
inner_data_pair = [list(z) for z in zip(inner_x_data, inner_y_data)]

outer_x_data = ["好","乐","哀","怒","惧","恶","惊"]
outer_y_data = [senti1['好_num'], senti1['乐_num'], senti1['哀_num'], senti1['怒_num'], senti1['惧_num'], senti1['恶_num'], senti1['惊_num']]
outer_data_pair = [list(z) for z in zip(outer_x_data, outer_y_data)]


(
    Pie(init_opts=opts.InitOpts(width="1600px", height="800px"))
    .add(
        series_name="情感",
        data_pair=inner_data_pair,
        radius=[0, "30%"],
        label_opts=opts.LabelOpts(position="inner"),
    )
    .add(
        series_name="情绪",
        radius=["40%", "55%"],
        data_pair=outer_data_pair,
        label_opts=opts.LabelOpts(
            position="outside",
            formatter="{a|{a}}{abg|}\n{hr|}\n {b|{b}: }{c}  {per|{d}%}  ",
            background_color="#eee",
            border_color="#aaa",
            border_width=1,
            border_radius=4,
            rich={
                "a": {"color": "#999", "lineHeight": 22, "align": "center"},
                "abg": {
                    "backgroundColor": "#e3e3e3",
                    "width": "100%",
                    "align": "right",
                    "height": 22,
                    "borderRadius": [4, 4, 0, 0],
                },
                "hr": {
                    "borderColor": "#aaa",
                    "width": "100%",
                    "borderWidth": 0.5,
                    "height": 0,
                },
                "b": {"fontSize": 16, "lineHeight": 33},
                "per": {
                    "color": "#eee",
                    "backgroundColor": "#334455",
                    "padding": [2, 4],
                    "borderRadius": 2,
                },
            },
        ),
    )
    .set_global_opts(legend_opts=opts.LegendOpts(pos_left="left", orient="vertical"))
    .set_series_opts(
        tooltip_opts=opts.TooltipOpts(
            trigger="item", formatter="{a} <br/>{b}: {c} ({d}%)"
        )
    )
    .render("./output/情感分析.html")
)



# 情感云图
wordcloud_by_dict('好', './output/词云图_好.html', senti1['好_dict'])
wordcloud_by_dict('乐', './output/词云图_乐.html', senti1['乐_dict'])
wordcloud_by_dict('哀', './output/词云图_哀.html', senti1['哀_dict'])
wordcloud_by_dict('怒', './output/词云图_怒.html', senti1['怒_dict'])
wordcloud_by_dict('惧', './output/词云图_惧.html', senti1['惧_dict'])
wordcloud_by_dict('恶', './output/词云图_恶.html', senti1['恶_dict'])
wordcloud_by_dict('惊', './output/词云图_惊.html', senti1['惊_dict'])

# 输出词云图
wordcloud(text=text_all, 
          title='词云图', 
          html_path='./output/词云图-总.html')


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
    sen = [list(senti.values())[3:10]]
   
    rad.add(
            series_name=name,
            data=sen,
            linestyle_opts=opts.LineStyleOpts(color=random_color()),
        )
    

rad.set_series_opts(label_opts=opts.LabelOpts(is_show=False))
rad.set_global_opts(
        title_opts=opts.TitleOpts(title="情绪雷达图"), legend_opts=opts.LegendOpts()
    )
rad.render("./output/情绪雷达图.html")


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
bar.render("./output/情绪堆叠柱状图.html")

