#!/usr/bin/env python
# coding: utf-8

# In[1]:


import docx
import openpyxl


# In[2]:


doc = docx.Document('test.docx')


# In[3]:


para = doc.paragraphs[0]
t = para.text
t = t.replace(t,'置換')
para.text = t


# In[4]:


doc.save('test_save.docx')

