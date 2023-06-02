from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import re
import pandas as pd
import time

# ### PDF converter function definition

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


# #### Category Dictionary and Column Generation

#dictionary and categorizing
def categorize_strings(strings):
    categories = {
        'must': 'Test',
        'should': 'recommendation',
        'can': 'option',
        'may': 'permission'
    }
    
    headerPattern0 = r'\d+[a-zA-Z]|[a-zA-Z]+\d+'
    categorized_strings = []
    
    for string in strings:
        categorized_string = {'category': None}
        count = 0
        dig = False
        for char in string:
            if char.isdigit():
                dig = True
            if char.isupper():
                count += 1
        if count > 12 or '..' in string or '--' in string or re.search(headerPattern0, string) is not None:
            categorized_string = {'category': 'Header'}
        else:
            for keyword, category in categories.items():
                if keyword in string.lower():
                    categorized_string['category'] = category
                    break

        categorized_strings.append(categorized_string)
    
    return categorized_strings


# #### Division column Generation

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


# #### Random Symbol Cleaning 

def symbolClean(text):
    #Cleaned up super random symbols
    uselessPatterns = ['""',"‘‘", "-'", ",‘", "--", "-‘","--'',","',",",,", ',-']
    for i in uselessPatterns:
        cleanedtext = text.replace(i, "")
    spaced = cleanedtext.split()
    spaced = ' '.join(spaced)
    return spaced


# #### Lemmatization

# def lemmatization(text): 
#     ##Testing Lemmatization
#     #Not sure if it really worked
#     from nltk.stem import WordNetLemmatizer
#     stemmer = WordNetLemmatizer()

#     text = text.split()

#     lemmatizedtext = [stemmer.lemmatize(word) for word in text]
#     lemmatizedtext = ' '.join(text)
#     # Text.append(text)
#     return lemmatizedtext


# #### Useless short portions of text removal


def shortparaRemove(text):
    #Remove Useless short sentences less than 200 character??
    for sentence in text:
        if len(sentence) > 200:
            text.remove(sentence)
    return text

def compile(Source, CleanText, Categories, Sections):
    end_text = []
    for string in CleanText:
        stringy = {'Text': string}
        end_text.append(stringy)
    #okay this is the long way, not sure how 
    #I want to do it better but there is a way
    finished = []

    for index in range(len(CleanText)):
        element = {}
        
        identifier = {'ID': index}
        element.update(identifier)
        
        if isinstance(Source, dict):
            element.update(Source)
            
        if isinstance(Sections[index], dict):
            element.update(Sections[index])
            
        if isinstance(end_text[index], dict):
            element.update(end_text[index])

        if isinstance(Categories[index], dict):
            element.update(Categories[index])


        finished.append(element)
    return finished


import os

def toExcel(table, fileName):
    #output_dir = r'C:\Users\Ben\Desktop\PDF_Testing'
    #output_path = os.path.join(output_dir, f'{fileName}.xlsx')
    
    Excel = pd.DataFrame(table)
    pd.DataFrame.to_excel(Excel, fileName)
    
    return

#The fileName must have '_' between all of the words
def sourceName(fileName):
    words = fileName.split('_')
    sourceWord = ''
    for word in words:
        if word.isdigit():
            sourceWord += word
        else:
            sourceWord += word[0]
    source = {'Source': sourceWord}
    return source



from tkinter import Tk, Button, Label, Entry, filedialog

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    file_path_entry.delete(0, 'end')
    file_path_entry.insert(0, file_path)

def choose_output_location():
    output_location = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    output_location_entry.delete(0, 'end')
    output_location_entry.insert(0, output_location)

def process_file():
    input_file_path = file_path_entry.get()
    output_file_path = output_location_entry.get()

        #PDF to test
    PDF = input_file_path
    orgText = "That's not a file worth doing this in. Sorry you waited patiently though."
    orgText = convert_pdf_to_txt(PDF)
    text = orgText.replace("\n", " ")
        #Tons of random symbols like ';._' that need to be removed for a cleaner loo
    cleantext0 = symbolClean(text)
        #Makes words simpiler, not sure if helpful but useful for ML later
        #cleantext1 = lemmatization(cleantext0)
        #Splits sentences into seperate strings
    split = re.split(r'(?<=\.)[ \n]', cleantext0)
    cleantext2 = shortparaRemove(split)
        ## Section covers categorization of strings
    categories = categorize_strings(cleantext2)
        ## Section division
    sections = Division(cleantext2)
        #creates Source name
    source = sourceName(PDF)
        #Creating the list for excel upload
    table = compile(source, cleantext2, categories, sections)
        #Sending table to Excel
    toExcel(table, output_file_path)
    print('done')

    success_label = Label(window, text="File processing successful!")
    success_label.pack()
    window.destroy()


# Create the main window
window = Tk()
window.title("PDF to Excel Converter")

# File path label and entry field
file_path_label = Label(window, text="Choose a PDF:")
file_path_label.pack()

file_path_entry = Entry(window, width=50)
file_path_entry.pack()

choose_file_button = Button(window, text="Browse", command=choose_file)
choose_file_button.pack()

# Output location label and entry field
output_location_label = Label(window, text="Output file location:")
output_location_label.pack()

output_location_entry = Entry(window, width=50)
output_location_entry.pack()

choose_output_button = Button(window, text="Browse", command=choose_output_location)
choose_output_button.pack()

# Process button
process_button = Button(window, text="Convert", command=process_file)
process_button.pack()

# Run the main window event loop
window.mainloop()
