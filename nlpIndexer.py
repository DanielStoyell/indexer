
# TODO:
# Detect doc type, convert to pdf using comtype if docx for reading
# Use pypdf2 to grab words by page
# Rewrite construction algorithm to work with new format (scrap article delimiter approach)
# Prettify output (in word doc?)

import sys
import comtypes.client
import PyPDF2
import os
import csv
import string
import nltk


def index():
  print "Parsing input..."
  if len(sys.argv) == 1:
    print "Please enter a file to index like so: 'python indexer.py <filename.docx OR filename.pdf>'"
    return
  file = sys.argv[1]
  if not os.path.isfile(file):
    print "File not found! Please enter a valid file path to the desired document, with the .docx extension included"
    return

  filename, filetype = os.path.splitext(file)
  pdfIsTemp = False
  if (filetype == ".pdf"):
    fileToProcess = file
  elif filetype == ".docx":
    print "Creating temporary pdf for processing..."
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(os.path.abspath(file))
    doc.SaveAs(os.path.abspath(filename + ".pdf"), FileFormat=17)
    doc.Close()
    word.Quit()
    fileToProcess = filename + ".pdf"
    pdfIsTemp = True
  elif filetype == ".doc":
    print ".doc to .pdf conversion is untested. Please convert your Word Document to .docx and try again"
    return
  else:
    print "Can only process word documents or pdfs! A filetype of " + filetype + " cannot be processed"
    return 

  print "Reading and processing excluded words..."
  excluded_words = []
  with open('exclude.csv', 'rb') as f:
    excludeFile = csv.reader(f, delimiter=",")
    for row in excludeFile :
      excluded_words = excluded_words + row

  # gotta get dat constant lookup time
  excluded_words = set(excluded_words)

  print "Opening PDF and preparing to process..."
  doc = open(fileToProcess, 'rb')
  reader = PyPDF2.PdfFileReader(doc)
  pages = reader.numPages

  print "Constructing index..."

  index = {}

  width = 30
  sys.stdout.write("[%s]" % (" " * width))
  sys.stdout.flush()
  sys.stdout.write("\b" * (width+1)) # return to start of line, after '['

  ticks = 0
  for pageNum in xrange(pages):
    while(pageNum >= (float(ticks-1)/width)*pages):
      ticks += 1
      sys.stdout.write("#")
      sys.stdout.flush()

    page = reader.getPage(pageNum)
    text = page.extractText()
    text = text.encode('ascii','ignore')
    parsed_words = nltk.pos_tag(nltk.word_tokenize(text))
    
    for parsed_word in parsed_words:
      if "NN" in parsed_word[1]:
        word = parsed_word[0].strip()
        word = word.lower()
        word = word.translate(string.maketrans("",""), string.punctuation)
        if not word in excluded_words and word != "" and not is_number(word):
          if word in index:
            index[word].add(pageNum+1)
          else:
            index[word] = set([pageNum+1])

  sys.stdout.write("\n")
  doc.close()

  print "Writing to output file..."
  sortedKeys = index.keys()
  sortedKeys.sort()
  indexName = "nlp_" + filename + "_index.txt"
  with open(indexName, 'wb') as iFile:
    for key in sortedKeys:
      string_rep = key + ": "
      for p in index[key]:
        string_rep += str(p) + ", "
      string_rep = string_rep[:-2]
      iFile.write(string_rep + "\n")

  if pdfIsTemp:
    print "Deleting temporary pdf..."
    os.remove(fileToProcess)

  print "Complete!"

def is_number(s):
  try:
    float(s)
    return True
  except ValueError:
    return False

if __name__==  "__main__":
  index()