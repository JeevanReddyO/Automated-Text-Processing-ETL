Automated-Text-Processing-ETL


# NLP Article Analyzer & Web Scraper

A robust Python-based pipeline that extracts article text from URLs and performs detailed linguistic and sentiment analysis. This project automates the ETL (Extract, Transform, Load) process to derive insights such as polarity scores, fog index, and complexity metrics from unstructured text data.

##  Features

* **Automated Web Scraping:** Fetches title and body content from a list of URLs using `BeautifulSoup` and `requests`.
* **Resilient Extraction:** Handles various HTML structures and CSS class variations to ensure data integrity.
* **Sentiment Analysis:** Calculates Positive/Negative scores, Polarity, and Subjectivity using custom dictionaries.
* **Readability Metrics:** Computes Gunning Fox Index, Average Sentence Length, and Complex Word percentages.
* **Batch Processing:** Capable of handling hundreds of URLs via Excel input/output.

## Tech Stack

* **Language:** Python 3.x
* **Libraries:**
    * `pandas` (Data manipulation)
    * `requests` (HTTP Session handling)
    * `beautifulsoup4` (HTML Parsing)
    * `nltk` (Natural Language Processing & Tokenization)
    * `openpyxl` (Excel I/O)

##  Project Structure

text
├── crawler_main.py          # Script 1: Web Scraper
├── nlp_processor.py         # Script 2: Text Analysis Engine
├── Input.xlsx               # Source file containing URLs
├── Output_Data_Structure.xlsx # Final generated report
├── extracted_articles/      # Folder storing raw text files
├── StopWords/               # Directory for stop word lists
└── MasterDictionary/        # Directory for Positive/Negative dictionaries



**** Metrics Calculated *****
The tool computes the following variables for each article:

Sentiment: Positive Score, Negative Score, Polarity Score, Subjectivity Score.

Readability: Average Sentence Length, Percentage of Complex Words, Fog Index.

Linguistic: Average Number of Words Per Sentence, Complex Word Count, Word Count, Syllable Per Word, Personal Pronouns, Average Word Length.

###**Disclaimer
This project is for educational and assessment purposes. Ensure you have the right to scrape target websites before running the crawler on large datasets.
