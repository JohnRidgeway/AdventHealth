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
data = pd.read_excel("/Users/johnridgeway/Desktop/AdventHealth/Participant Packet AdventHealth Case-Study/Excel Documents with Project Data/3Q2020 Survey details.xlsx")

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


#Key word lists and lists to hold 1 if keyword is found and 0 otherwise
keyWordsShipping = ["ship","fedex","usps","mail","delivery","mailbox","ups","order"]
shippingList = []

keyWordsService = ["staff","everyone","service","representative","people","reps","person","manager","team","member"]
serviceList = []

keyWordsWebsite = ["website","online","on-line"]
websiteList = []

keyWordsTimeliness = ["wait","delay","time","slow","soon","quick","long","fast","forever","late"]
timelinessList = []

keyWordsRefill = ["auto-refill","refill","auto-fill"]
refillList = []


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

    #Check for shipping words
    if any(s in keyWordsShipping for s in tokens1):
        shippingList.append(1)
    elif any(s in keyWordsShipping for s in tokens2):
        shippingList.append(1)
    elif any(s in keyWordsShipping for s in tokens3):
        shippingList.append(1)
    elif any(s in keyWordsShipping for s in tokens4):
        shippingList.append(1)
    else:
        shippingList.append(0)

    #Check for service words
    if any(s in keyWordsService for s in tokens1):
        serviceList.append(1)
    elif any(s in keyWordsService for s in tokens2):
        serviceList.append(1)
    elif any(s in keyWordsService for s in tokens3):
        serviceList.append(1)
    elif any(s in keyWordsService for s in tokens4):
        serviceList.append(1)
    else:
        serviceList.append(0)
    
    #Check for website words
    if any(s in keyWordsWebsite for s in tokens1):
        websiteList.append(1)
    elif any(s in keyWordsWebsite for s in tokens2):
        websiteList.append(1)
    elif any(s in keyWordsWebsite for s in tokens3):
        websiteList.append(1)
    elif any(s in keyWordsWebsite for s in tokens4):
        websiteList.append(1)
    else:
        websiteList.append(0)

    #Check for timeliness words
    if any(s in keyWordsTimeliness for s in tokens1):
        timelinessList.append(1)
    elif any(s in keyWordsTimeliness for s in tokens2):
        timelinessList.append(1)
    elif any(s in keyWordsTimeliness for s in tokens3):
        timelinessList.append(1)
    elif any(s in keyWordsTimeliness for s in tokens4):
        timelinessList.append(1)
    else:
        timelinessList.append(0)

    #Check for refill words
    if any(s in keyWordsRefill for s in tokens1):
        refillList.append(1)
    elif any(s in keyWordsRefill for s in tokens2):
        refillList.append(1)
    elif any(s in keyWordsRefill for s in tokens3):
        refillList.append(1)
    elif any(s in keyWordsRefill for s in tokens4):
        refillList.append(1)
    else:
        refillList.append(0)

#Add columns to dataframe to display whether a given comment falls in a certain category
df["Shipping"] = shippingList
df["Service"] = serviceList
df["Website"] = websiteList
df["Timeliness"] = timelinessList
df["Refill"] = refillList

#Find entries that couldn't be sorted
df_noValue = df.loc[(df['Shipping'] == 0) & (df['Service'] == 0) & (df['Website'] == 0) & (df['Timeliness'] == 0) & (df['Refill'] == 0)]

#Print out results to Excel
df.to_excel("/Users/johnridgeway/Desktop/AdventHealth/Participant Packet AdventHealth Case-Study/Excel Documents with Project Data/Output.xlsx")

