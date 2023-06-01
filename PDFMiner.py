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


# # Actual Code ran proccessing ("Test.pdf")

# In[27]:


#PDF to test
PDF = "Test2.pdf"

orgText = convert_pdf_to_txt(PDF)
text = orgText.replace("\n", " ")


# In[16]:


#Testing Lemmatization
#Not sure if it really worked
from nltk.stem import WordNetLemmatizer
Text = []
stemmer = WordNetLemmatizer()

text = text.split()

text = [stemmer.lemmatize(word) for word in text]
text = ' '.join(text)
# Text.append(text)


# In[32]:


#Cleaned up super random symbols
uselessPatterns = ['""',"‘‘", "-'", ",‘", "--", "-‘","--'',","',",",,", ',-']
for i in uselessPatterns:
    cleanedtext = text.replace(i, "")
spaced = cleanedtext.split()
spaced = ' '.join(spaced)
spaced


# In[26]:


#Splits sentences into seperate strings
cleaned2text = re.split(r'(?<=\.)[ \n]', cleanedtext)
cleaned2text


# In[24]:


#Remove Useless short sentences less than 200 character??
for sentence in cleaned2text:
    if len(sentence) > 200:
        cleaned2text.remove(sentence)


# ## Section covers categorization of strings

# In[18]:


Categorized = categorize_strings(cleaned2text)


# ## Testing division

# In[11]:


divided = Division(cleaned2text)


# ## Compiles info together for export

# In[12]:


end_text = []
for string in cleaned2text:
    stringy = {'Text': string}
    end_text.append(stringy)
#okay this is the long way, not sure how 
#I want to do it better but there is a way
finished = []

for index in range(len(cleaned2text)):
    element = {}
    
    if isinstance(end_text[index], dict):
        element.update(end_text[index])
        
    if isinstance(Categorized[index], dict):
        element.update(Categorized[index])
        
    if isinstance(divided[index], dict):
        element.update(divided[index])
    
    finished.append(element)


# ## Creates Excel document using Pandas and exports to desktop labeled "PDFtoExcel.xlsx"

# In[13]:


Excel = pd.DataFrame(finished)


# In[14]:


pd.DataFrame.to_excel(Excel, r"C:\Users\Ben\Desktop\PDF_Testing\PDFtoExcel.xlsx")


# ## Testing

# In[36]:


import re

def split_combined_string(string):
    words = re.findall(r'\b[a-z]+\b', string)
    return ' '.join(words)

# Example usage
input_string = cleanedtext
split_string = split_combined_string(input_string)
print(split_string)


# In[ ]:




