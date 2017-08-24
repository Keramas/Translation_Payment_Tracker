import docx, sys, re, openpyxl, datetime, csv
from datetime import datetime

filename = sys.argv[1]

#Get text from specified Word document.

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

#Find total word count.

def wordcount(value):
    list = re.findall("(\S+)", value)
    return len(list)

value = getText(filename)

#Display total word count as well as total cost for the job.

print("Total number of words:", wordcount(value))
print("Total cost:", (wordcount(value))*5,"yen")    \\value currently set to 5 yen per word.

#Output date and payment amount to CSV

now = datetime.now()

Date = now.strftime('%m/%d/%Y')
Cost = (wordcount(value))*5

with open(r"FILENAME.csv", 'a', newline='') as f:   \\change FILENAME to the path and name of your CSV
    w = csv.writer(f, delimiter=',')
    w.writerow([Date, Cost])

print("CSV has been updated!")
