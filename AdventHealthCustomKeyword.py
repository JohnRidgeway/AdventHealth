import nltk
import string
from nltk.stem import LancasterStemmer
from nltk.stem import PorterStemmer
import pandas as pd
import os
import numpy as np
from afinn import Afinn
afinn = Afinn()

#Read in excel sheet
data = pd.read_excel("3Q2020 Survey details.xlsx")

#Select only "Please provide additional feedback column"
df = pd.DataFrame(data, columns= ['Please provide additional feedback. (Optional)'])

#Rename only column Feedback and get rid of any null values
df.columns = ['Feedback']
df = df.dropna()

#Contains afinn score
afinnList = []

#Assign affinity score to each respose and save in affinList
for i in range(len(df)):
    afinnList.append(afinn.score_with_pattern(str(df.iloc[i])))

#Add affinity score as column to the dataframe
df["Affinity Score"] = afinnList   

#Take in user keywords
userKeywords = str (input ("Enter comma separated keywords to search for: "))
userList = userKeywords.split(",")
keywordList = []

#Function to tokenize and stem text as well as convert to lower case using Lancaster
def tokenize_and_stem_Lancaster(text):
    tokens = nltk.tokenize.word_tokenize(text)
    tokens = [token.lower().strip(string.punctuation)
              for token in tokens if token.isalnum()]
    tokens = [LancasterStemmer().stem(token) for token in tokens]
    return tokens

#Function to tokenize and stem text as well as convert to lower case using Porter
def tokenize_and_stem_Porter(text):
    tokens = nltk.tokenize.word_tokenize(text)
    tokens = [token.lower().strip(string.punctuation)
              for token in tokens if token.isalnum()]
    tokens = [PorterStemmer().stem(token) for token in tokens]
    return tokens

#Function to tokenize and convert to lower case
def tokenize_lower(text):
    tokens = nltk.tokenize.word_tokenize(text)
    tokens = [token.lower().strip(string.punctuation)
              for token in tokens if token.isalnum()]
    return tokens

#Function to tokenize
def tokenize(text):
    tokens = nltk.tokenize.word_tokenize(text)
    return tokens

#Checks for words in each category, assigning a list a 1 if word is found, a 0 otherwise
for i in range(len(df)):
    tokens1 = tokenize_and_stem_Lancaster(str(df['Feedback'].iloc[i]))
    tokens2 = tokenize_and_stem_Porter(str(df['Feedback'].iloc[i]))
    tokens3 = tokenize_lower(str(df['Feedback'].iloc[i]))
    tokens4 = tokenize(str(df['Feedback'].iloc[i]))

    #Check for keywords
    if any(s in userList for s in tokens1):
        keywordList.append(1)
    elif any(s in userList for s in tokens2):
        keywordList.append(1)
    elif any(s in userList for s in tokens3):
        keywordList.append(1)
    elif any(s in userList for s in tokens4):
        keywordList.append(1)
    else:
        keywordList.append(0)

#Add columns to dataframe to display whether a given comment falls in a certain category
df["User Keywords"] = keywordList

#Print out results to Excel
df.to_excel("Output.xlsx")

