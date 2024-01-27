import os
import requests
from bs4 import BeautifulSoup
import pandas as pd

# Load the input data from Excel
input_file_path = 'input_file_path'
output_folder = 'output'

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Read the input data
df = pd.read_excel(input_file_path, engine='openpyxl')

title_class = 'inspect and find blog title class' 
content_class = 'inspect and find blog content'
# Function to extract article text from a URL
def extract_article_text(url):
    try:
        response = requests.get(url)
        response.raise_for_status()

        # Use BeautifulSoup to parse HTML content
        soup = BeautifulSoup(response.text, 'html.parser')



        # Extract title and article text
        title_element = soup.find(class_= title_class)

        if title_element:
            heading_element = title_element.find('h1', class_='your class')

            heading_text = heading_element.text.strip() if heading_element else ' '
        else:
            title_element = soup.find(class_= 'class found after inspection')
            if(title_element):

                heading_element = title_element.find('h1', class_='title class')
                heading_text = heading_element.text.strip() if heading_element else ' '
            else:
                title_element = soup.find(class_= 'inner classes')
                heading_element = title_element.find('h1', class_='title class')
                heading_text = heading_element.text.strip() if heading_element else ' '
        content_elements = soup.find(class_=content_class)

        if content_elements:
            # Extract text from paragraph tags within the specified class
            paragraphs_in_class = content_elements.find_all('p')
            paragraph_texts = [paragraph.get_text(strip=True) for paragraph in paragraphs_in_class]

            # Combine paragraph texts into a single string
            content_text = ' '.join(paragraph_texts)
        else:
            content_elements = soup.find(class_='content class')
            paragraphs_in_class = content_elements.find_all('p')
            paragraph_texts = [paragraph.get_text(strip=True) for paragraph in paragraphs_in_class]
            content_text = ' '.join(paragraph_texts)
        return heading_text, content_text
    except Exception as e:
        print(" Error, file saved as no text")
        return None, None

# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    # Extract article text
    title, content_text = extract_article_text(url)
   
    # Save the extracted text to a text file
    if title and content_text:
        output_file_path = os.path.join(output_folder, f'{url_id}.txt')
        with open(output_file_path, 'w', encoding='utf-8') as file:
            file.write(f'{title}\n\n{content_text}')

        print(f"Article extracted and saved for URL_ID {url_id}.")

print("Extraction complete.")
