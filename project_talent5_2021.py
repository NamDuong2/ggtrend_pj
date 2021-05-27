#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd 
from pytrends.request import TrendReq
pytrend = TrendReq()


# In[2]:


# đọc tự động file excel chứa keyword yêu cầu

excel_file = "keytrends.xlsx"
df = pd.read_excel(excel_file)
df.info()


# In[3]:


# xem dataframe xuất ra
df


# In[4]:


# tạo list các list chứa keyword của mỗi mục

all_kw = []
kw_list = []

for j in range(0,len(df.columns)):
    for i in range(0,len(df)):        
        kw_list.append(df.iloc[i,j])        
        kw_list = [x for x in kw_list if str(x) != 'nan']
    all_kw.append(kw_list)
    kw_list = []

print(all_kw)


# In[9]:


# Tạo dataframe cho mục news

keywords = []       #list tạm thời
d = {}
def check_trends():
    pytrend.build_payload(keywords, timeframe='2020-01-01 2020-12-31', geo='VN')
    d[i] = pytrend.interest_over_time()            #lưu từng dataframe tương ứng mỗi vòng lặp vào dict
    
i = 0
for kw in all_kw[0]:
    keywords.append(kw)
    check_trends()
    keywords.pop()           # mỗi vòng lặp xóa kí tự cuối để có thể chạy liên tục mà không bị giới hạn 5 tìm kiếm của ggtrend
    i += 1

newsframe = pd.concat(d, axis=1)         # nối các dataframe trong dict thành 1 bảng lớn đẩy đủ các keyword
newsframe.columns = newsframe.columns.droplevel(0)
newsframe = newsframe.drop('isPartial', axis = 1)

# những mục còn lại tương tự như mục news


# In[10]:


# Tạo dataframe cho mục person

keywords = [] #list tạm thời
d = {}
def check_trends():
    pytrend.build_payload(keywords, timeframe='2020-01-01 2020-12-31', geo='VN')
    d[i] = pytrend.interest_over_time()
    
i = 0
for kw in all_kw[1]:
    keywords.append(kw)
    check_trends()
    keywords.pop()               # mỗi vòng lặp xóa kí tự cuối để có thể chạy liên tục mà không bị giới hạn 5 tìm kiếm của ggtrend
    i += 1

person_frame = pd.concat(d, axis=1)
person_frame.columns = person_frame.columns.droplevel(0)
person_frame = person_frame.drop('isPartial', axis = 1)


# In[12]:


# Tạo dataframe cho mục film

keywords = []     #list tạm thời
d = {}
def check_trends():
    pytrend.build_payload(keywords, timeframe='2020-01-01 2020-12-31', geo='VN')
    d[i] = pytrend.interest_over_time()
i = 0
for kw in all_kw[2]:
    keywords.append(kw)
    check_trends()
    keywords.pop()             # mỗi vòng lặp xóa kí tự cuối để có thể chạy liên tục mà không bị giới hạn 5 tìm kiếm của ggtrend
    i += 1

film_frame = pd.concat(d, axis=1)
film_frame.columns = film_frame.columns.droplevel(0)
film_frame = film_frame.drop('isPartial', axis = 1)


# In[13]:


# Tạo dataframe cho mục elearning online

keywords = []         #list tạm thời
d = {}
def check_trends():
    pytrend.build_payload(keywords, timeframe='2020-01-01 2020-12-31', geo='VN')
    d[i] = pytrend.interest_over_time()
i = 0
for kw in all_kw[3]:
    keywords.append(kw)
    check_trends()
    keywords.pop()          # mỗi vòng lặp xóa kí tự cuối để có thể chạy liên tục mà không bị giới hạn 5 tìm kiếm của ggtrend
    i += 1

eonline_frame = pd.concat(d, axis=1)
eonline_frame.columns = eonline_frame.columns.droplevel(0)
eonline_frame = eonline_frame.drop('isPartial', axis = 1)


# In[14]:


# Tạo dataframe cho mục diseases

keywords = []           #list tạm thời
d = {}
def check_trends():
    pytrend.build_payload(keywords, timeframe='2020-01-01 2020-12-31', geo='VN')
    d[i] = pytrend.interest_over_time()
    
i = 0
for kw in all_kw[4]:
    keywords.append(kw)
    check_trends()
    keywords.pop()            # mỗi vòng lặp xóa kí tự cuối để có thể chạy liên tục mà không bị giới hạn 5 tìm kiếm của ggtrend
    i += 1

d_frame = pd.concat(d, axis=1)
d_frame.columns = d_frame.columns.droplevel(0)
d_frame = d_frame.drop('isPartial', axis = 1)

    


# In[15]:


# Tạo dataframe cho mục songs

keywords = []          #list tạm thời
d = {}
def check_trends():
    pytrend.build_payload(keywords, timeframe='2020-01-01 2020-12-31', geo='VN')
    d[i] = pytrend.interest_over_time()
    
i = 0
for kw in all_kw[5]:
    keywords.append(kw)
    check_trends()
    keywords.pop()            # mỗi vòng lặp xóa kí tự cuối để có thể chạy liên tục mà không bị giới hạn 5 tìm kiếm của ggtrend
    i += 1

s_frame = pd.concat(d, axis=1)
s_frame.columns = s_frame.columns.droplevel(0)
s_frame = s_frame.drop('isPartial', axis = 1)


# In[16]:


# Tạo dataframe cho mục travel

keywords = []          #list tạm thời
d = {}
def check_trends():
    pytrend.build_payload(keywords, timeframe='2020-01-01 2020-12-31', geo='VN')
    d[i] = pytrend.interest_over_time()
    
i = 0
for kw in all_kw[6]:
    keywords.append(kw)
    check_trends()
    keywords.pop()           # mỗi vòng lặp xóa kí tự cuối để có thể chạy liên tục mà không bị giới hạn 5 tìm kiếm của ggtrend
    i += 1

tr_frame = pd.concat(d, axis=1)
tr_frame.columns = tr_frame.columns.droplevel(0)
tr_frame = tr_frame.drop('isPartial', axis = 1)


# In[19]:


# chuyển tất cả dataframe thành sheet tương ứng trong file excel

path = r'C:\Users\ADMIN\vn_trend_2020.xlsx'

writer = pd.ExcelWriter(path, engine = 'xlsxwriter')

newsframe.to_excel(writer, sheet_name = 'news')
person_frame.to_excel(writer, sheet_name = 'person')
film_frame.to_excel(writer, sheet_name = 'film')
eonline_frame.to_excel(writer, sheet_name = 'elearning online')
d_frame.to_excel(writer, sheet_name = 'diseases')
s_frame.to_excel(writer, sheet_name = 'songs')
tr_frame.to_excel(writer, sheet_name = 'travel')
writer.save()


# In[ ]:




