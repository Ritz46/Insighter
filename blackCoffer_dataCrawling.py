import os
import requests
import pandas as pd
from bs4 import BeautifulSoup

def extract_article_text(url):
    try:
        response = requests.get(url)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')

        title = soup.find('title').text.strip()
        article_text = ""
        for paragraph in soup.find_all('p'):
            article_text += paragraph.text.strip() + '\n'

        return title, article_text

    except Exception as e:
        print(f"Error extracting data from {url}: {str(e)}")
        return None, None

input_file = "input.xlsx"
df = pd.read_excel(input_file)

output_directory = "extracted_articles"
os.makedirs(output_directory, exist_ok=True)

for index, row in df.iterrows():
    url_id = str(row['URL_ID'])
    url = row['URL']

    title, article_text = extract_article_text(url)

    if title and article_text:

        output_filename = os.path.join(output_directory, f"{url_id}.txt")
        with open(output_filename, 'w', encoding='utf-8') as f:
            
            f.write(f"Title: {title}\n\n")
            f.write(article_text)

print("Extraction completed. Text files are saved in the 'extracted_articles' directory.")
