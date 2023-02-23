#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import json
import os
import kss # 한국 NLP라이브러리
import time # 시간측정 라이브러리
import re # 파이썬 정규식
from konlpy.tag import Okt # 품사조사
from konlpy.tag import Kkma #한국어 토큰화
from nltk.tokenize import word_tokenize #영어 토큰화
from nltk.tokenize import TreebankWordTokenizer
from tensorflow.keras.preprocessing.text import text_to_word_sequence
okt = Okt()
kkma = Kkma()
tokenizer = TreebankWordTokenizer()


# In[3]:
# 필수 변수 값 지정
# 파일명
fileName = "data.xlsx"
# 결과 파일명
resultFile = "result.xlsx"
# 데이터파일 위치
input__position = str(os.getcwd() + "\\input\\" + fileName)
# 결과파일 위치
output__position = str(os.getcwd() + "\\output\\" + resultFile)


# In[4]:
# 판다스:데이터 읽어오기
df = pd.read_excel(input__position)
# Check Data
df1=df.질문
# 총 데이터 개수 
index = df1.index.stop
print("총 질문 개수 : " ,index)

# In[7]:

# 배열만들기
a=[]
# 진행율 배열만들기
b=int()
# 진행 시간 측정 start
start = time.time()

for i in range(index):
    # 한국어NLP 토큰화
    knToken=kkma.nouns(df1[i])    
    # 영어NLP 토큰화        # 특수문자 정규화
    enTokenSub = re.sub('[^a-zA-Z0-9-=+,#/\:^.@*\"※~ㆍ!』‘|\(\)\[\]`\'…》\”\“\’·\ ]','',df1[i]).strip()
    enToken = text_to_word_sequence(enTokenSub)
#     print('OKT 명사 분석 :',token)
#     token = tokenizer.tokenize(df1[i])
    a = a + knToken
    a = a + enToken
    
# -------------- 진행률 --------------#
    b = b + 1
    percent = round(100/index * b,2)
    print(percent,'%')
# 진행 시간 측정 end
end = time.time()
print("진행시간",round(end - start,4), "s")
# -------------- 진행률 --------------#

# In[8]:

# 토큰 카운터 작업
e = dict()
for key in a:
    check=e.get(key)
    if check == None:        
        e[key]=1
    else:
        e[key]= e[key]+1

# 토큰 카운터 높은 순으로 필터
__result=sorted(e.items(), key=lambda item: item[1], reverse = True)

result = []
for key in list(__result):
    t = []
    t.append(key[0])
    t.append(key[1])
    t.append(str(round(key[1]/index*100,2)))
    result.append(t)

# In[9]:

# print(result)
columns = ['key','count','percentage ( % )']
# 데이터 프레임화
ddff=pd.DataFrame(result,columns=columns)


# In[10]:


# 액셀 생성
ddff.to_excel(output__position)
print("진행완료")
