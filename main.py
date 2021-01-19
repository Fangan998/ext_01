#import here
#\import{
import xlrd

#from PIL import Image
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import jieba
import numpy as np
from collections import Counter
import pandas as pd
#____________________________________________________________________\

book = xlrd.open_workbook("ww_01.xlsx")

list_book_txt=open("list_txt.txt","w",encoding='utf-8')

sheelt1 = book.sheets()[0]

nrows = sheelt1.nrows

'''
print(f"Row: {nrows}")
nrow_3 = sheelt1.col_values(3)

ncols = sheelt1.ncols
print(f"Col : {ncols}")


cell_x =sheelt1.cell(1,3).value
print(f'Cell_x \'s單元格的值：{cell_x}')
'''
#______________________________________________________________
#class here
#----------------------------------
def word():
    count_file = open('count.txt','w',encoding='utf')
    
    with open("cut_book.txt",'r',encoding='utf-8',errors='ignore') as f:
        text = f.read().strip().split()
    
    stx_freq = count_segment_freq(text)

    print(stx_freq.shape)
    print(stx_freq.describe())
    print(stx_freq.head(10))
    print(stx_freq.tail(10))
    print(stx_freq.info())
    stx_freq.to_excel('sss.xlsx')

#----------------------------------
def _wit_(a):
    t_lsit = f"{a}\n\n" 
    list_book_txt.write(t_lsit)

#-----------------------------------
def wil():
    j = 1
    for i in range(nrows):
        if j >= nrows:
            break;
        a = sheelt1.cell(j,3).value
        a = a.strip()
        _wit_(a)
        j+= 1 

#---------------------------------------
def cut():
    list_book = open('list_txt.txt','r',encoding='utf-8')
    cut_list=open('cut_book.txt',"w",encoding='utf-8')
    jieba.set_dictionary("dictionary/dict.txt")
    jieba.load_userdict("dictionary/yyy.txt")
    jieba.load_userdict("dictionary/universities.txt")
    for i in list_book.readlines():
        a=i
        b=jieba.cut(a.strip())
        c=' '.join(b)
        d = f"{c}\n"
        cut_list.write(d)
    list_book.close()
    cut_list.close()      
    
#---------------------------------------
def country():
    con = 0
    lsit_book = open('list_txt.txt','r',encoding='utf-8')
    country_key=["日本", "美國", "英國", "瑞士", "俄羅斯",
    "印度", "法國", "芬蘭", "加拿大", "紐西蘭",
    "拉脫維亞", "杜拜", "荷蘭", "泰國", "韓國",
    "香港", "大陸","中國","台灣","德國"]
    #print(len(country_key))
    #=======================================
    for i in lsit_book.readlines():
        a = i
        for j in range(len(country_key)):
            num = a.count(country_key[j])
            if num > 0:
                con+=num
    #=======================================
    lsit_book.close()
    return con
    
#----------------------------------------   
def nom_cloud():
    
    with open("cut_book.txt",'r',encoding='utf-8',errors='ignore') as f:
        text = f.read()
    font=r'msjh.ttc'
    wordcloud = WordCloud(font_path=font,background_color='white',width=4000,height=3000).generate(text)
    
    #plt.imshow(wordcloud)
    #plt.axis('off')
    #plt.show()
    
    wordcloud.to_file('output_nom.png')
#---------------------------------------
def count_segment_freq(seg_list):
  seg_df = pd.DataFrame(seg_list,columns=['seg'])
  seg_df['count'] = 1
  sef_freq = seg_df.groupby('seg')['count'].sum().sort_values(ascending=False)
  sef_freq = pd.DataFrame(sef_freq)
  return sef_freq     
#--------------------------------------

#__________________________________________________________________ 
#main

wil() 
cut()
word()
print(f"country: {country()}")
nom_cloud()


list_book_txt.close()