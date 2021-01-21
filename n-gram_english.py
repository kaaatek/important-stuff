import pandas as pd
import os, nltk, re, time
from nltk.util import ngrams
from nltk.corpus import stopwords
from collections import Counter
# global variable for the stop word list, with language definition
STOP_WORDS = set(stopwords.words('english'))
# time stamp format
timestr = time.strftime("%Y-%m-%d_%H-%M-%S")

# location of the file from which to read data
path = os.path.dirname(os.path.realpath(__file__))
filename = "input.xlsx"
# create data frame from Excel
df = pd.read_excel("{}/{}".format(path, filename))
print ("Reading input file...")

def generate_ngrams(row, n):
    text = row["text"]
    # replace punctuation (apart from "-") with a space
    text = re.sub(r'[^\w\s\-]',' ',text)
    # remove duplicate spaces
    text = re.sub(' +',' ',text)
    # remove leading/trailing spaces
    text = re.sub(r'^\s+|\s+$', '',text)
    # split text into individual words
    words = text.split(" ")
    # create new list for clean text
    clean_words = []
    # loop to clean stopwords and add to list
    for w in words:
        if w.lower() not in STOP_WORDS:
            clean_words.append(w)
    # loop to create n-gram tuples
    for i in n:
        n_grams = ngrams(clean_words, i)
        # change tuples into a string
        n_grams = [" ".join(n_gram) for n_gram in n_grams]
        row[str(i)] = n_grams
    return row

# list of numbers for n-gram creation    
n = [1, 2, 3]

# apply user-defined function to data frame
df = df.apply(generate_ngrams, axis = "columns", n = n)
print ("Removing stopwords...")
print ("Creating n-grams...")
#df.to_excel(str(path+"/"+"all_grams.xlsx"), index = False)
# write n-grams into an Excel sheet
with pd.ExcelWriter(timestr+'_n_gram_output.xlsx') as writer:
    for i in n:
        # creating list of n-grams
        column_values = df[str(i)].tolist()
        # creating new list for all n-grams
        all_grams = []
        # adding n-grams to list
        for j in column_values:
           all_grams.extend(j)
        # count n-gram frequency
        all_grams = Counter(all_grams)
        # write n-grams to columns
        final_df = pd.DataFrame.from_dict(all_grams, orient ='index').reset_index()
        # rename column names
        final_df = final_df.rename(columns={'index':str(i)+' grams', 0:'count'})
        # write to specific sheet
        final_df.to_excel(writer, sheet_name = str(i)+"_gram", index = False)
    print ("Counting n-grams...")
    
print("Finished")

