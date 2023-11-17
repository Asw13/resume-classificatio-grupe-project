import streamlit as st
import pickle
import base64
import matplotlib.pyplot as plt
import tensorflow as tf
from tensorflow import keras
from glob import glob
import pandas as pd
import docx2txt
from win32com import client
import os
from docx import Document 
import PyPDF2
import textract
import spacy
import shutil
import json
import re
import matplotlib.pyplot as plt
import seaborn as sns
from wordcloud import WordCloud
from nltk.tokenize import TweetTokenizer
import string
from collections import Counter
import string
from nltk.tokenize import word_tokenize
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import sweetviz
from sklearn.feature_extraction.text import TfidfVectorizer
from scipy.sparse import hstack
from sklearn.preprocessing import LabelEncoder
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import accuracy_score,recall_score,precision_score,f1_score,roc_auc_score,confusion_matrix


st.set_page_config(page_title="Resume Extractor", page_icon=":smiley:", layout="wide")
# Set background image
def add_bg_from_local(image_file):
    with open(image_file, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read())
    st.markdown(f"""
    <style>
    .stApp {{
        background-image: url(data:image/{"png"};base64,{encoded_string.decode()});
        background-size: cover
    }}
    </style>
    """,
    
    unsafe_allow_html=True
    )
add_bg_from_local('o.jpg')
# Add title with CSS styles
st.markdown( """
<style>
h1 {
    text-align: center;
    color: #333333;
}
</style>
""",unsafe_allow_html=True)

st.title("Resume category Project")
file_path=glob("D:\\Astik Project\\Project 2\\New folder\\strimlate\\New folder\\*")
def docReader(doc_file_name): 
    ## 1) Initiate an object that interfaces to Word
    word = client.Dispatch("Word.Application")
    word.Visible = False 
    
    
    ## 2) Open the Word document to read in
    _ = word.Documents.Open(doc_file_name)

    ## 3) Extract the paragraphs and close the connections
    doc = word.ActiveDocument
    paras = doc.Range().text    
    doc.Close()
    word.Quit()
    return paras 

def read_doc_file(file):
    
    
    # Save the file to a specific directory
    file_dir = "D:\\Astik Project\\Project 2\\New folder\\strimlate\\New folder\\"


    
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    with open(os.path.join(file_dir, file.name), 'wb') as f:
        shutil.copyfileobj(file, f)
    text = docReader(file)
    return text

def read_docx_file(file):
    text = docx2txt.process(file)
    return text

def read_pdf_file(file):
    pdf_reader = PyPDF2.PdfFileReader(file)
    text = ''
    for page in range(pdf_reader.getNumPages()):
        text += pdf_reader.getPage(page).extractText()
    return text

def get_resume_data(file):
    data1=[]
    name1 = []
    names1 = []
    file_type = os.path.splitext(file.name)[1]
    if file_type == '.docx':
        text = read_docx_file(file)
        data1.append(text)
    elif file_type == '.pdf':
        text = read_pdf_file(file)
        data1.append(text)
    elif file_type== '.doc':
        text= read_doc_file(file);
        data1.append(text)
    data1 = pd.DataFrame(data=data1,columns=['data'])
   
    names1 = pd.DataFrame(data = name1,columns=["Name"])
    data1 = pd.concat([data1,names1],axis=1)
    # Extract relevant information from text (e.g. name, contact information, education, experience)
    # and store in a Pandas dataframe
    # ...
    return data1


def skile(file):
    df=get_resume_data(file)
    test = spacy.load('en_core_web_sm')
    skills = []
    for i in range(len(df.data)):
        ts = test(" ".join(df.data[i].split('\n'))) # we have splitted our data with '\n' and rejoined with space. 
        tt = []
        for ent in ts.ents:
            if ent.label_.upper() == 'ORG':
                tt.append(ent.text)
        skills.append(tt) # appending all skills to the list skills
    df['skills']=0  # creating new columns skills and assiging 0 to every column
    for i in range(len(df.skills)):
            df.skills[i] = skills[i]
    df.to_csv('dataS.csv')
    return df

def eda(file):
    df=skile(file)
    for i in range(len(df.skills)):
        lower_words=[Text.lower() for Text in df.skills[i]]
    df.skills[i] = lower_words
    
    for i in range(len(df.skills)):
        ab =[]
    
    # finding the duplicat values
    return df
   
def duplica(file):
    df=eda(file)
    # finding the duplicat values
    for i in range(len(df.skills)):
        numbers = df.skills[i]
        counts = dict(Counter(numbers))
        duplicates = {key:value for key, value in counts.items() if value > 1}
    for i in range(len(df.skills)):
        lm = set(df.skills[i])
        df.skills[i] = list(lm)
    return df
def remoo(file):
    df=duplica(file)
    
        # Removing the unwanted data like '',' ','s','cs' which contains length upto 2
    for i in range(len(df.skills)):
        er = []
        for j in range(len(df.skills[i])):
            if (len(df.skills[i][j]) >= 3) :
                ab = df.skills[i][j]
                er = er + [ab]
        df.skills[i] = er
    # Removing all punctuation

    for i in range(len(df.skills)):
        for j in range(len(df.skills[i])):
            df.skills[i][j] = df.skills[i][j].translate(str.maketrans('','',string.punctuation))
    for i in range(len(df.skills)):
        for j in range(len(df.skills[i])):
            df.skills[i][j] = ''.join([i for i in df.skills[i][j] if not i.isdigit()])
    # Removing all spaces 
   
    for i in range(len(df.skills)):
        for j in range(len(df.skills[i])):
            df.skills[i][j] = word_tokenize(df.skills[i][j])
            df.skills[i][j] =  ' '.join(df.skills[i][j])
    # Removing the unwanted data like '',' ','s','cs' which contains length upto 2
    
    for i in range(len(df.skills)):
        er = []
        for j in range(len(df.skills[i])):
            if (len(df.skills[i][j]) >= 3) :
                ab = df.skills[i][j]
                er = er + [ab]
        df.skills[i] = er
    
    
    
    return df

def coppy(file):
    df=remoo(file)
    df1 = df.copy(deep=True) # it will take copy of df 
    for i in range(len(df1.skills)):
        df1.skills[i] = " ".join(df1.skills[i]) # converting list into string
    for i in range(len(df1.skills)):
        df1.skills[i] = word_tokenize(df1.skills[i]) # tokenization
    nltk.download('stopwords') # importing stop words

    my_stop_words = stopwords.words('english')
    my_stop_words.append(' ')
    my_stop_words.append('&') # adding reqiued stop words
    # removing stop words
    for i in range(len(df1.skills)):    
        df1.skills[i] = [word for word in df1.skills[i] if not word in my_stop_words ]
    # joining the words into single document (removing the tokenization)
    for i in range(len(df1.skills)):
        df1.skills[i] =  ' '.join(df1.skills[i])
    # Lemmatization
    Lemmatizer = WordNetLemmatizer()
    for i in range(len(df1.skills)):
        lemmas = []
        for token in df1.skills[i].split():
            lemmas.append(Lemmatizer.lemmatize(token))
        df1.skills[i] = lemmas
    # joining the words into single document (removing the tokenization)
    for i in range(len(df1.skills)):
        df1.skills[i] =  ' '.join(df1.skills[i])

    df3 = df1['skills'].str.split(' ').explode().str.strip().value_counts()
    return df1 ,df3

# Define the KNN model
def knn_model():
    # Load the training data
    df = pd.read_csv('classified_resumes.csv')
    x = df['skills']
    y = df['category']

    # Vectorize the skills using TF-IDF
    word_vectorizer = TfidfVectorizer(sublinear_tf=True)
    word_vectorizer.fit(x)
    x = word_vectorizer.transform(x)

    # Encode the categories
    LE = LabelEncoder()
    y = LE.fit_transform(y)

    # Train the KNN model
    knn = KNeighborsClassifier(n_neighbors=5, p=2)
    knn.fit(x, y)

    return knn, word_vectorizer, LE


def classify_resume(knn, word_vectorizer, LE, df3):
    # Extract the skills from the resume
        
        
        
    skills = df3['skills']
    
    # Vectorize the skills using TF-IDF
    x = word_vectorizer.transform(skills)

    # Predict the category using the trained KNN model
    y_pred = knn.predict(x)

        # Decode the predicted category
    category = LE.inverse_transform(y_pred)

    return category

def app():

    
    
    file = st.file_uploader('Upload your resume', type=['doc', 'docx', 'pdf'])
    
    
    # Define the classification function
  

    
    if st.button('Classify'):
        # Check if a file or text is uploaded
        if file is not None:
            df3,df2 = coppy(file)
        else :
            st.write("Uplode file")
        # Load the KNN model and other preprocessed data
        knn, word_vectorizer, LE = knn_model()
        st.write('skils:',df2)
        # Classify the resume and display the predicted category
        category = classify_resume(knn, word_vectorizer, LE, df3)
        
        st.write('Predicted category:', category)


if __name__ == '__main__':
    app()
