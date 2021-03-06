#!/usr/bin/env python

__author__ = "Martin Rydén"
__copyright__ = "Copyright 2016, Martin Rydén"
__license__ = "MIT"
__version__ = "1.0.1"
__email__ = "pemryd@gmail.com"

import re
import requests
from bs4 import BeautifulSoup

avail_lang = {
'cz':["cestina","czech"],
'dk':["dansk","danish"],
'nl':["nederlands","dutch"],
'en':["english","english"],
'es':["espanol","spanish"],
'fr':["francais","french"],
'it':["italiano","italian"],
'pl':["jezyk-polski","polish"],
'no':["norsk","norwegian"],
'fi':["suomi","finnish"],
'hu':["magyar","hungarian"],
'po':["portugues","portugese"],
'bp':["portugues-brasileiro","brazilian-portuguese"],
'sv':["svenska","swedish"]}

print("Available languages:\tCode:\n")
for k, v in avail_lang.items():
     st = ("%s:\t%s" % (v[1], k))
     padding = 25
     print(st.expandtabs(padding))

# Choose language to translate from
while True:
     langf = input("\nTranslate function from: ").lower()
     if(langf in avail_lang.keys()):
          break
     else:
          print("Try again.")

langt = "en" # Translate to English by default
# Choose language to translate to
if(langf == "en"):
     while True:
          langt = input("\nTranslate function to: ").lower()
          if(langt in avail_lang.keys()):
               break
          else:
               print("Try again.")
# Input Excel function
function = input("\nFunction: ")

# Find all tables in selected languages
url = "http://www.piuha.fi/excel-function-name-translation/index.php?page=%s-%s.html" % (avail_lang[langf][0], avail_lang[langt][1])
r = requests.get(url)
soup = BeautifulSoup(r.content, "html.parser")

# Create dict and enumerate over table values
tdict = dict((i,t) for i,t in enumerate(soup.find_all('td'))) 

# Split by and keep any non-alphanumeric delimiter (except .)
# the full string including delimters is added at the end
spfunction = re.split('([^\\w.])', function)

# Split by and remove delimiters, filter out empty elements
function = list(filter(None, re.split(r'[\W]+', function)))

# Only keep elements longer than 2 chars, in order to limit matching
function = [f for f in function if len(f) > 2]

# Iterate over table values, function parts
# Add original and translated values to dict
trdict = {}
#for i, t in tdict.items():
for i, t in tdict.items():
    for x in function:
         if(str(x) in str(t.getText())):
              if(langf == "en"):
                   if(t.getText() not in trdict.keys()): # Prevent duplicates
                       fr = (tdict[i+1].getText().split(','))[0]
                       to = (t.getText().split(','))[0]
                       trdict[to] = fr # If translating from English, set key to English
              else: 
                   fr = (t.getText().split(','))[0]
                   to = (tdict[i+1].getText().split(','))[0]
                   trdict[fr] = to # If translating from non-English, set key to non-English


new = [trdict.get(x,x) for x in spfunction] # Get translated value from dict

if(langf != "en"):
     new = '=' + ''.join(new).replace(';',',')
elif(langf == "en"):
     new = '=' + ''.join(new).replace(',',';')

print("\n%s"%new)
