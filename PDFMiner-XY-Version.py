#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pdfminer.layout import LAParams, LTTextBox
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal
from io import StringIO
import re
import pandas as pd


# ## PDF converter function definition

# In[2]:


from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
from pdfminer.converter import PDFPageAggregator

# The Actual PDF
PDF = 'Test2.pdf'
# Output Excel File
exceloutput = ''

# PDF reader setup
fp = open(PDF, 'rb')
rsrcmgr = PDFResourceManager()
laparams = LAParams()
device = PDFPageAggregator(rsrcmgr, laparams=laparams)
interpreter = PDFPageInterpreter(rsrcmgr, device)
pages = PDFPage.get_pages(fp)

combinedList = []

for i, page in enumerate(pages, start=1):
    interpreter.process_page(page)
    layout = device.get_result()
    for element in layout:
        if isinstance(element, LTTextBoxHorizontal):
            x, y, text = element.bbox[0], element.bbox[3], element.get_text()
            x = round(x, 2)
            y = round(y, 2)
            text = text.replace('\n', ' ')
            combinedList.append((i, text, x, y))

# Print the combined list containing page number, text, and coordinates
for entry in combinedList:
    page_num, text, x, y = entry
    print(f'Page: {page_num}')
    print(f'Text: {text}')
    print(f'Coordinates: {x}')
    print(f'Coordinates: {y}')


# In[3]:


Excel = pd.DataFrame(combinedList)
pd.DataFrame.to_excel(Excel, r"C:\Users\Ben\Desktop\PDF_Testing\PDFtoExcel2.xlsx")


# In[4]:


# workingPage = combinedList[403:432]


# In[5]:


# #Grouping Header with Paragraphs
# threshold = 100
# i = 0
# j = 0
# for i in range(0,len(workingPage)-1):
#     for j in range(0,len(workingPage)):
#         differenceX = abs(workingPage[i][2] - workingPage[j][2])
#         differenceY = abs(workingPage[i][3] - workingPage[j][3])
        
#         if differenceX < threshold and differenceY < threshold:
#             merge = list(workingPage[i])
#             merge[1] += workingPage[i][1]
#             workingPage[i] = tuple(merge)
#             workingPage.pop(i+1)
# workingPage


# In[15]:


# from pdfminer.layout import LAParams
# from pdfminer.converter import PDFResourceManager, PDFPageAggregator
# from pdfminer.pdfpage import PDFPage
# from pdfminer.layout import LTTextBoxHorizontal

# document = open('', 'rb')
# #Create resource manager
# rsrcmgr = PDFResourceManager()
# # Set parameters for analysis.
# laparams = LAParams()
# # Create a PDF page aggregator object.
# device = PDFPageAggregator(rsrcmgr, laparams=laparams)
# interpreter = PDFPageInterpreter(rsrcmgr, device)
# for page in PDFPage.get_pages(document):
#     interpreter.process_page(page)
#     # receive the LTPage object for the page.
#     layout = device.get_result()
#     for element in layout:
#         if isinstance(element, LTTextBoxHorizontal):
#             print(element.get_text())


# In[ ]:




