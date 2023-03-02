#!/usr/bin/env python
# coding: utf-8

# In[7]:


import win32com.client


# In[8]:


outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)


# In[9]:


mail.to = 'yawatanabe@micron.com;'
mail.cc = 'mosano@micron.com;ofushitani@micron.com;'
mail.subject = 'Daily_Report'
mail.bodyFormat = 1
mail.body = '''本日の作業日報を送付します。

1. 残業申請 Automation tool test
2. P-Tech向け wiki 作成

明日の予定

'''
# mail app が起動する（内容確認）
mail.display(True)


# In[ ]:




