import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import time

# --- CONFIGURATION ---
INPUT_FILENAME = 'Input.xlsx'
OUTPUT_FOLDER = 'extracted_articles'

class WebScraper:
    def __init__(self):
        # Create output directory if it doesn't exist
        if not os.path.exists(OUTPUT_FOLDER):
            os.makedirs(OUTPUT_FOLDER)
        
        # Setup a session with headers to look like a real browser
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })

    def get_page_content(self, target_url):
        """Fetches the URL and parses title and text."""
        try:
            response = self.session.get(target_url, timeout=15)
            if response.status_code != 200:
                print(f"   [!] Failed to connect. Status: {response.status_code}")
                return None
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Extract Title
            head_tag = soup.find('h1')
            title_text = head_tag.get_text().strip() if head_tag else "No Title Found"

            # Extract Main Article Body
            # List of common container classes used in WordPress/Websites
            content_classes = ['td-post-content', 'article-content', 'entry-content', 'post-content']
            body_text = ""
            
            for css_class in content_classes:
                container = soup.find('div', class_=css_class)
                if container:
                    body_text = container.get_text(separator='\n').strip()
                    break
            
            # Fallback strategy: grab all paragraphs if specific div missing
            if not body_text:
                paragraphs = soup.find_all('p')
                body_text = " ".join([p.get_text() for p in paragraphs])

            return f"{title_text}\n\n{body_text}"

        except Exception as err:
            print(f"   [!] Error occurred: {err}")
            return None

    def run(self):
        print(f"--- Starting Extraction from {INPUT_FILENAME} ---")
        try:
            df = pd.read_excel(INPUT_FILENAME)
        except FileNotFoundError:
            print(f"Error: Could not find {INPUT_FILENAME}")
            return

        success_count = 0
        
        for row in df.itertuples():
            # Adjust column access based on your Excel headers
            art_id = str(row.URL_ID)
            art_url = row.URL
            
            print(f"Processing ID: {art_id}")
            content = self.get_page_content(art_url)
            
            if content:
                save_path = os.path.join(OUTPUT_FOLDER, f"{art_id}.txt")
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                success_count += 1
            else:
                print(f"   [x] Skipped {art_id}")

        print(f"\n--- Completed. Successfully extracted {success_count} articles. ---")

if __name__ == "__main__":
    bot = WebScraper()
    bot.run()