#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import re
import pandas as pd


# ## PDF converter function definition

# In[2]:


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    setattr(laparams, 'all_texts', True)
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text


# ## Category Dictionary and Column Generation

# In[3]:


#dictionary and categorizing
def categorize_strings(strings):
    categories = {
        'must': 'Test',
        'should': 'recommendation',
        'can': 'option',
        'may': 'permission'
    }
    
    categorized_strings = []
    
    for string in strings:
        categorized_string = {'category': None}
        
        for keyword, category in categories.items():
            if keyword in string.lower():
                categorized_string['category'] = category
                break
        
        categorized_strings.append(categorized_string)
    
    return categorized_strings


# ## Division column Generation

# In[4]:


def Division(strings):
    divided = []
    current_division = 'None'
    pattern = r'^\d+\.\d+'  # Regex pattern for "number . number" at the start of the string
    
    for string in strings:
        if re.match(pattern, string):
            current_division = string.split(' ',1)[0]
            current_division = {'Section': current_division}
        divided.append(current_division)
    return divided


# ## Random Symbol Cleaning 

# In[5]:


def symbolClean(text):
    #Cleaned up super random symbols
    uselessPatterns = ['""',"‘‘", "-'", ",‘", "--", "-‘","--'',","',",",,", ',-']
    for i in uselessPatterns:
        cleanedtext = text.replace(i, "")
    spaced = cleanedtext.split()
    spaced = ' '.join(spaced)
    return spaced


# ## Lemmatization

# In[6]:


def lemmatization(text): 
    ##Testing Lemmatization
    #Not sure if it really worked
    from nltk.stem import WordNetLemmatizer
    stemmer = WordNetLemmatizer()

    text = text.split()

    lemmatizedtext = [stemmer.lemmatize(word) for word in text]
    lemmatizedtext = ' '.join(text)
    # Text.append(text)
    return lemmatizedtext


# ## Useless short portions of text removal

# In[7]:


def shortparaRemove(text):
    #Remove Useless short sentences less than 200 character??
    for sentence in text:
        if len(sentence) > 200:
            text.remove(sentence)
    return text


# ## Compiler

# In[13]:


def compile(CleanText, Categories, Sections):
    end_text = []
    for string in CleanText:
        stringy = {'Text': string}
        end_text.append(stringy)
    #okay this is the long way, not sure how 
    #I want to do it better but there is a way
    finished = []

    for index in range(len(CleanText)):
        element = {}

        if isinstance(end_text[index], dict):
            element.update(end_text[index])

        if isinstance(Categories[index], dict):
            element.update(Categories[index])

        if isinstance(Sections[index], dict):
            element.update(Sections[index])

        finished.append(element)
    return finished


# ## Creates Excel document using Pandas and exports to desktop labeled "PDFtoExcel.xlsx"

# In[34]:


import os

def toExcel(table, fileName):
    output_dir = r'C:\Users\Ben\Desktop\PDF_Testing'
    output_path = os.path.join(output_dir, f'{fileName}.xlsx')
    
    Excel = pd.DataFrame(table)
    pd.DataFrame.to_excel(Excel, output_path)
    
    return


# # Actual Code ran proccessing ("Test.pdf")

# In[35]:


#PDF to test
PDF = "Test4.pdf"

#orgText = convert_pdf_to_txt(PDF)
text = orgText.replace("\n", " ")
cleantext0 = symbolClean(text)
cleantext1 = lemmatization(cleantext0)
#Splits sentences into seperate strings
split = re.split(r'(?<=\.)[ \n]', cleantext1)
cleantext2 = shortparaRemove(split)
## Section covers categorization of strings
categories = categorize_strings(cleantext2)
## Section division
sections = Division(cleantext2)
#Creating the list for excel upload
table = compile(cleantext2, categories, sections)
#Sending table to Excel
toExcel(table, 'work')


# ## Testing

# In[ ]:


import re

def split_combined_string(string):
    words = re.findall(r'\b[a-z]+\b', string)
    return ' '.join(words)

# Example usage
input_string = cleanedtext
split_string = split_combined_string(input_string)
print(split_string)


# In[ ]:




