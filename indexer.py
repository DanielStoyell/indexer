import sys
import docx
import os
import csv
import string

def index():
  print "Parsing input..."
  if len(sys.argv) == 1:
    print "Please enter a file to index like so: 'python indexer.py <filename.docx>'"
    return
  file = sys.argv[1]
  if not os.path.isfile(file):
    print "File not found! Please enter a valid file name located within the indexer.py folder."
    return

  "Reading and processing excluded words..."
  excluded_words = []
  with open('exclude.csv', 'rb') as f:
    excludeFile = csv.reader(f, delimiter=",")
    for row in excludeFile :
      excluded_words = excluded_words + row

  # gotta get dat constant lookup time
  excluded_words = set(excluded_words)

  print "Reading document..."
  doc = docx.Document(file)
  paragraphs = doc.paragraphs

  print "Constructing index..."
  progress = -1
  index = {}
  for i in range(len(paragraphs)):
    percentComplete = float(i)/len(paragraphs)
    percentComplete = int(round(percentComplete*10))*10
    if percentComplete > progress:
      print str(int(percentComplete)) + "%"
      progress = percentComplete

    text = paragraphs[i].text
    words = text.split()

    for word in words:
      word = word.strip() #Remove trailing whitespcae
      word = word.encode('ascii', 'ignore') #Strip Word bullshit
      word = word.lower() #Make lowercase
      word = word.translate(string.maketrans("",""), string.punctuation) #strip punctuation
      if word not in excluded_words:
        if word in index:
          index[word].add(i)
        else:
          index[word] = set([i])

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