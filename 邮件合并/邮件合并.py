#!/usr/bin/env python
# coding: utf-8

# In[18]:


import xlrd
from docxtpl import DocxTemplate
import pandas as pd
import numpy as np


# In[79]:


xls = pd.read_excel('C:/Users/51may/Desktop/信达证券/中科坤健终止交易台账(2)(1).xlsx',sheet_name='Sheet1')


# In[80]:


xls


# In[81]:


rows = xls.index[xls['日期'].isnull()][0]
rows


# 思想就是把word中想批量生成的数据替换为{{a}}的形式，然后在python中导入excel，对应填入word中即可，一键生成多个word

# In[82]:


for i in range(rows):
    id1 = xls['编号'][i]
    customer = xls['客户名称'][i]
    bank = xls['开户行'][i]
    account = xls['账号'][i]
    main_id = xls['主协议编号'][i]
    sup_id = xls['补充协议编号'][i]
    confirm_id = xls['交易确认书'][i]
    intent_id1 = xls['意向书编号1'][i]
    intent_id2 = xls['意向书编号2'][i]
    target_name = xls['标的名称'][i]
    target_id = xls['标的代码'][i]
    stop_num = xls['终止数量'][i]
    start_price = xls['初始交易价格'][i]
    end_date = xls['日期'][i].date()
    end_price = xls['终止均价'][i]
    end_pro_amt = xls['终止部分履约保障金'][i]
    flo_inc_amt = xls['浮动收益'][i]
    pre_stat_inc_amt = xls['收取固定收益'][i]
    end_amt = xls['结算净额'][i]
    # 打开一个模板
    doc = DocxTemplate("C:/Users/51may/Desktop/信达证券/模板.docx")
    data = {} #  构造填充模板需要的数据
    data['id'] = id1
    data['customer'] = customer
    data['bank'] = bank
    data['account'] = account
    data['main_id'] = main_id
    data['sup_id'] = sup_id
    data['confirm_id'] = confirm_id
    data['intent_id1'] = intent_id1
    data['intent_id2'] = intent_id2
    data['customer'] = customer
    data['target_name'] = target_name
    data['target_id'] = target_id
    data['end_date_myd'] = '%d年%d月%d日' %(end_date.year, end_date.month, end_date.day)
    data['stop_num'] = str("{:,.2f}".format(stop_num))
    data['start_price'] = str("{:,.2f}".format(stop_num))
    data['end_date'] = end_date
    data['end_price'] = str("{:,.2f}".format(end_price))
    data['end_pro_amt'] = str("{:,.2f}".format(end_pro_amt))
    data['flo_inc_amt'] = str("{:,.2f}".format(flo_inc_amt))
    data['pre_stat_inc_amt'] = str("{:,.2f}".format(pre_stat_inc_amt))
    data['end_amt'] = str("{:,.2f}".format(end_amt))
    doc.render(data) # 填充数据data到模板
    doc.save("{}.docx".format(target_name)) # 可以按照标的名称进行存储


# In[ ]:




