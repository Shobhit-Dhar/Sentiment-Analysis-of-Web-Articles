import os

#import nltk
import requests
from bs4 import BeautifulSoup
import pandas as pd
from nltk.corpus import stopwords

#nltk.download('stopwords')

# Function to load stop words from multiple files
def load_stop_words(file_paths):
    stop_words = set()
    for file_path in file_paths:
        with open(file_path, 'r') as file:
            stop_words.update(word.strip().lower() for word in file.read().splitlines())
        stop_words.update(stopwords.words('english'))
    return stop_words

# Function to remove stop words from text
def remove_stop_words(text, stop_words):
    words = text.split()
    filtered_words = [word for word in words if word.lower() not in stop_words]
    return ' '.join(filtered_words)

df=pd.read_excel('input.xlsx')

# List of stop words files
stop_words_files = ['StopWords_Auditor.txt', 'StopWords_Currencies.txt', 'StopWords_DatesandNumbers.txt',
                    'StopWords_Generic.txt', 'StopWords_GenericLong.txt', 'StopWords_Geographic.txt',
                    'StopWords_Names.txt']

# Load stop words from the stop words files
stop_words = load_stop_words(stop_words_files)

for i in range(100):
    # URL of the article
    url = df['URL'].iloc[i]

    # Send a GET request to the URL
    response = requests.get(url)

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')
    res = soup.title

    # Find the main content of the article
    article_content = soup.find('div', class_='td-post-content')

    # Extract the text from the article content
    try:
        article_text = res.get_text() + '\n' + article_content.get_text(separator='\n', strip=True)

    except AttributeError:
        continue

    # Remove stop words from the article text
    filtered_text = remove_stop_words(article_text, stop_words)

    # Specify the directory to save the text file
    directory = os.getcwd() + r'\txtfiles'

    # Specify the file name and path
    file_name = 'blackassign000'+ str(i+1)+ '.txt'
    file_path = os.path.join(directory, file_name)

    # Write the filtered text to the file
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(filtered_text)

    print(f"Filtered article text has been saved to {file_path}")
