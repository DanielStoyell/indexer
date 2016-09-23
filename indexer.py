import sys
import docx
import os
import csv
import string

ARTICLE_DELIMITER = "ARTICLEDELIMITER"

def index():
  print "Parsing input..."
  if len(sys.argv) == 1:
    print "Please enter a file to index like so: 'python indexer.py <filename.docx>'"
    return
  file = sys.argv[1]
  if not os.path.isfile(file):
    print "File not found! Please enter a valid file path to the desired document, with the .docx extension included"
    return

  "Reading and processing excluded words..."
  excluded_words = []
  with open('exclude.csv', 'rb') as f:
    excludeFile = csv.reader(f, delimiter=",")
    for row in excludeFile :
      excluded_words = excluded_words + row

  # gotta get dat constant lookup time
  excluded_words = set(excluded_words)

  print "Reading and processing document..."
  doc = docx.Document(file)
  paragraphs = doc.paragraphs

  words = []
  for p in paragraphs:
    p_words = p.text.split()
    words += p_words

  print "Constructing index..."
  progress = -1
  index = {}
  article = 0
  for i in xrange(len(words)):
    percentComplete = float(i)/len(words)
    percentComplete = int(round(percentComplete*10))*10
    if percentComplete > progress:
      print str(int(percentComplete)) + "%"
      progress = percentComplete

    word = words[i]
    word = word.strip() #Remove trailing whitespcae
    word = word.encode('ascii', 'ignore') #Strip Word Unicode bullshit
    word = word.translate(string.maketrans("",""), string.punctuation) #strip punctuation

    if word == ARTICLE_DELIMITER:
      article += 1
    else:
      if article > 0:
        word = word.lower() #Make lowercase
        if word not in excluded_words:
          if word in index:
            index[word].add(article)
          else:
            index[word] = set([article])
        

  print "Writing to output file..."
  sortedKeys = index.keys()
  sortedKeys.sort()
  indexName = file[:-5] + "_index.txt"
  with open(indexName, 'wb') as iFile:
    for key in sortedKeys:
      string_rep = key + ": "
      for p in index[key]:
        string_rep += str(p) + ", "
      string_rep = string_rep[:-2]
      iFile.write(string_rep + "\n")

  print "Complete!"

if __name__==  "__main__":
  index()