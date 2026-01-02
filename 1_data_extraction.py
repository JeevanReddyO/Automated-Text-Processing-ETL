import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

# Create directory for text files
if not os.path.exists('extracted_articles'):
    os.makedirs('extracted_articles')

def extract_article(url):
    try:
        # Mimic a browser to avoid 403 Forbidden errors
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36'
        }
        page = requests.get(url, headers=headers)
        
        # Parse HTML
        soup = BeautifulSoup(page.content, 'html.parser')
        
        # 1. Extract Title
        title_tag = soup.find('h1')
        title = title_tag.get_text().strip() if title_tag else "No Title"
        
        # 2. Extract Article Text
        # Attempt to find the main article content div
        article_div = soup.find('div', class_='td-post-content')
        
        # Fallback for different site layouts
        if not article_div:
             article_div = soup.find('div', class_='article-content')
             
        if article_div:
            # Get text and separate lines
            article_text = article_div.get_text(separator=' \n').strip()
        else:
            # Last resort: grab all paragraphs (can be messy)
            paragraphs = soup.find_all('p')
            article_text = ' '.join([p.get_text() for p in paragraphs])

        return title + "\n" + article_text
        
    except Exception as e:
        print(f"Error extracting {url}: {e}")
        return None

def main():
    print("Loading Input.xlsx...")
    # Load Input file
    try:
        input_df = pd.read_excel('Input.xlsx')
    except FileNotFoundError:
        print("Error: 'Input.xlsx' not found.")
        return

    # Loop through each URL
    for index, row in input_df.iterrows():
        url_id = str(row['URL_ID'])
        url = row['URL']
        
        print(f"Processing {url_id}...")
        text_content = extract_article(url)
        
        if text_content:
            # Save to text file
            with open(f'extracted_articles/{url_id}.txt', 'w', encoding='utf-8') as file:
                file.write(text_content)
        else:
            print(f"Failed to extract {url_id}")

    print("Extraction complete.")

if __name__ == "__main__":
    main()