import csv
import re
import string
import nltk
nltk.download('stopwords')
nltk.download('wordnet')
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer

comp = []
processed = []
out = [['Title','Description','Media','Price','Tax','SKU','Type','Vendor','Tags']]

with open('Barrera-Store-CSV.csv','r') as file:
    reader = csv.reader(file)
    for d in reader:
        comp.append(d)

for entry in comp:
    temp = entry
    e1 = re.findall('<div class="value" itemprop="description">.*', entry[2])
    if len(e1) > 0:
        e1 = e1[0]
        e1 = re.sub('<div class="value" itemprop="description">', '', e1)
        e1 = re.sub('</div>', '', e1)
    else:
        e1 = ''
    e2 = re.findall('<div class="value"><p>.*', entry[2])
    if len(e2) > 0:
        e2 = e2[0]
        e2 = re.sub('<div.*><p>', '', e2)
        e2 = re.sub('</p></div>', '', e2)
        e2 = re.sub('</p>','',e2)
    else:
        e2 = ''
    temp[2] = [e1,e2]
    temp[25] = re.sub('^\s+','',temp[25])
    temp[26] = re.sub('^\n+','',temp[26])
    temp[26] = re.sub('^\s+','',temp[26])
    if temp[0] != "":
        processed.append(temp)

#for x in processed[2]:
#    print(x)
#    print('----------------------------')

#print('----------------------------')
#print('----------------------------')

for p in processed:
    text = '<p>'+'\n'.join(p[2])+'</p>\n<h2>ADDITIONAL INFO</h2>\n<table width="100%"><tbody><tr><td>&nbsp;<span data-mce-fragment="1">Benefits</span></td><td><span>'+p[18]+'</span></td>\n</tr>\n<tr>\n<td><span>FAQs</span></td>\n<td><span>'+p[19]+'</span></td>\n</tr>\n<tr>\n<td><span>Skin Type</span></td>\n<td><span>'+p[22]+'</span></td>\n</tr>\n<tr>\n<td><span>Size</span></td>\n<td><span>'+p[23]+'</span></td>\n</tr>\n<tr>\n<td><span>Brand</span></td>\n<td><span>'+p[24]+'</span></td>\n</tr>\n</tbody>\n</table>'
    tokens = nltk.word_tokenize('\n'.join(p[2]))
    stop = stopwords.words('english')
    exclude = ['','skin',"'s"]
    clean_tokens = [t.lower() for t in tokens if t.lower() not in string.punctuation and t.lower() not in stop and t.lower() not in exclude]
    lemmatizer = WordNetLemmatizer()
    lem_tokens = [lemmatizer.lemmatize(t) for t in clean_tokens]
    word_list = {word: sum([1 for w in lem_tokens if w == word]) for word in lem_tokens}
    tags = [k for k,v in word_list.items() if v > 1]
    temp = [p[1],'"{}"'.format(text.replace('\n','<br>')),p[28],p[10],"TRUE",p[9],"Non-Surgical",p[3],','.join(tags)]
    out.append(temp)
out.remove(out[1])
f = open('Barrera-Shopify-CSV.csv', 'w')
for x in out:
    f.write('/\c'.join(x))
    f.write('\n')
f.close()
print('complete')


