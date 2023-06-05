from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import re
import pandas as pd
import time
import threading
import os
from tkinter import Tk, Button, Label, Entry, filedialog, Checkbutton, IntVar


#General Header formats to look at in documents
header1 = r'^\d+\.\d+' #number.number format for header


#Main/What happens after convert is hit
def process_file():
    input_file_path = file_path_entry.get() #Gathers the file input location
    output_file_path = output_location_entry.get() #Gathers the file output location
    # custom_text = custom_text_entry.get()
    header1Bool = header1_var.get() #Gathers the boolean on header format 1

    if header1Bool == True: #sets the header pattern to header 1 if box is ticked
        headerpattern = header1


###MAIN###

    PDF = input_file_path #path to PDF for conversion
    orgText = convert_pdf_to_txt(PDF) #converts the pdf to text
    text = orgText.replace("\n", " ") #replaces newlines with spaces
    cleantext0 = symbolClean(text) #removes extra nonsense symbols
    #cleantext1 = lemmatization(cleantext0) #makes text into simple words, lemmatizes it
    split = re.split(r'(?<=\.)[ \n]', cleantext0) #splits string into list by sentences
    cleantext2 = shortparaRemove(split) #removes short paragraphs that have no use or are pointless
    categories = categorize_strings(cleantext2) #categorize the text
    sections = Division(cleantext2, headerpattern) #labels text w/ the sections they correspond to
    source = sourceName(PDF) #creates source name 
    table = compile(source, cleantext2, categories, sections, orgText, cleantext0, cleantext2) #takes all the previous data and puts it into a 2D list for export
    toExcel(table, output_file_path) #sends 2D list to excel
    print('done')

    window.after(0, update_gui) #uses seperate cpu core and closes program when finished


# ### PDF converter function definition


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager() #sets resource manager
    retstr = StringIO() #sets retstr
    codec = 'utf-8' #sets the processing codec
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

def Division(strings, pattern):
    divided = []
    current_division = 'None'
    
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

def compile(Source, CleanText, Categories, Sections,orgText,cleantext0,cleantext2):
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

        if index < 1:
            element.update({'Original Text': orgText})
            element.update({'Better Original Text': orgText.replace("\n", " ")})
            element.update({'First Clean (Random Symbol Removal)': cleantext0 })
            element.update({'Second Clean (Short Sentence Removal\ Sentence Splitting)': cleantext2})

        if isinstance(end_text[index], dict):
            element.update(end_text[index])

        if isinstance(Categories[index], dict):
            element.update(Categories[index])


        finished.append(element)
    return finished

def toExcel(table, fileName):
    #output_dir = r'C:\Users\Ben\Desktop\PDF_Testing'
    #output_path = os.path.join(output_dir, f'{fileName}.xlsx')
    
    Excel = pd.DataFrame(table)
    pd.DataFrame.to_excel(Excel, fileName)
    
    return

#The fileName must have '_' between all of the words
def sourceName(fileName):
    source = ''
    name = fileName.rfind('\\')
    if fileName != -1:
        source = fileName[name + 1:]
    source = {'Source': source}
    return source


def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    file_path_entry.delete(0, 'end')
    file_path_entry.insert(0, file_path)

def choose_output_location():
    output_location = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    output_location_entry.delete(0, 'end')
    output_location_entry.insert(0, output_location)
 
def update_gui():
        # Update the GUI once the task is completed
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
choose_file_button.pack(pady = 10)

# Output location label and entry field
output_location_label = Label(window, text="Output file location:")
output_location_label.pack()

output_location_entry = Entry(window, width=50)
output_location_entry.pack()

choose_output_button = Button(window, text="Browse", command=choose_output_location)
choose_output_button.pack()
choose_output_button.pack(pady = 10)

# Custom text label and entry field
# custom_text_label = Label(window, text="Header Format:")
# custom_text_label.pack()

# custom_text_entry = Entry(window, width=50)
# custom_text_entry.pack()
# custom_text_entry.pack(pady = 10)
#Checkboxes for headers

header1_label = Label(window, text='Header Format: 1.1')
header1_label.pack(pady = 5)
header1_var = IntVar()
header1_button = Checkbutton(window, text="Correct?", variable=header1_var)
header1_button.pack()
# Process button
process_button = Button(window, text="Convert", command=process_file)
process_button.pack()


# Run the main window event loop
window.mainloop()
