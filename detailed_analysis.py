import pandas as pd
import os
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import re

# Download required NLTK data
nltk.download('punkt')
try:
    nltk.download('punkt_tab')
except:
    pass

# --- CONFIGURATION ---
STOPWORDS_DIR = 'StopWords'
MASTER_DICT_DIR = 'MasterDictionary'
TEXT_DIR = 'extracted_articles'
INPUT_FILE = 'Input.xlsx'
OUTPUT_FILE = 'Output_Data_Structure_Updated.xlsx'

# --- LOADING FUNCTIONS ---
def load_stop_words():
    stop_words = set()
    if not os.path.exists(STOPWORDS_DIR):
        print(f"WARNING: Directory '{STOPWORDS_DIR}' not found.")
        return stop_words
        
    for filename in os.listdir(STOPWORDS_DIR):
        path = os.path.join(STOPWORDS_DIR, filename)
        try:
            with open(path, 'r', encoding='latin-1') as f:
                words = f.read().splitlines()
                for word in words:
                    if '|' in word:
                        word = word.split('|')[0]
                    stop_words.add(word.strip().lower())
        except Exception as e:
            print(f"Error reading {filename}: {e}")
    return stop_words

def load_master_dictionary(stop_words):
    positive_words = set()
    negative_words = set()
    
    if not os.path.exists(MASTER_DICT_DIR):
        print(f"WARNING: Directory '{MASTER_DICT_DIR}' not found.")
        return positive_words, negative_words

    # Load Positive Words
    try:
        with open(os.path.join(MASTER_DICT_DIR, 'positive-words.txt'), 'r', encoding='latin-1') as f:
            for word in f.read().splitlines():
                if word not in stop_words:
                    positive_words.add(word.strip().lower())
    except FileNotFoundError:
        print("positive-words.txt not found.")

    # Load Negative Words
    try:
        with open(os.path.join(MASTER_DICT_DIR, 'negative-words.txt'), 'r', encoding='latin-1') as f:
            for word in f.read().splitlines():
                if word not in stop_words:
                    negative_words.add(word.strip().lower())
    except FileNotFoundError:
        print("negative-words.txt not found.")
        
    return positive_words, negative_words

# --- ANALYSIS HELPER FUNCTIONS ---
def count_syllables(word):
    word = word.lower()
    if word.endswith('es') or word.endswith('ed'):
        return 0 
    
    vowels = "aeiouy"
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    return max(1, count)

def analyze_text(text, stop_words, positive_dict, negative_dict):
    sentences = sent_tokenize(text)
    tokens = word_tokenize(text)
    
    cleaned_tokens_raw = [word.lower() for word in tokens if word.isalpha()]
    cleaned_tokens_sentiment = [word for word in cleaned_tokens_raw if word not in stop_words]
    
    # 1. Sentiment Scores
    positive_score = sum(1 for word in cleaned_tokens_sentiment if word in positive_dict)
    negative_score = sum(1 for word in cleaned_tokens_sentiment if word in negative_dict)
    
    # 2. Polarity Score
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    
    # 3. Subjectivity Score
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_tokens_sentiment) + 0.000001)
    
    # 4. Readability Analysis
    total_words = len(cleaned_tokens_raw)
    total_sentences = len(sentences)
    avg_sentence_length = total_words / total_sentences if total_sentences > 0 else 0
    
    complex_words = [word for word in cleaned_tokens_raw if count_syllables(word) > 2]
    percentage_complex_words = (len(complex_words) / total_words) if total_words > 0 else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    # 5. Other Metrics
    avg_words_per_sentence = avg_sentence_length
    complex_word_count = len(complex_words)
    word_count = len(cleaned_tokens_sentiment)
    
    total_syllables = sum(count_syllables(word) for word in cleaned_tokens_raw)
    syllable_per_word = total_syllables / len(cleaned_tokens_raw) if len(cleaned_tokens_raw) > 0 else 0
    
    # Personal Pronouns (Exclude 'US')
    pronoun_regex = re.compile(r'\b(I|we|my|ours|us)\b', re.IGNORECASE)
    pronoun_count = 0
    for word in tokens:
        if pronoun_regex.match(word) and word != 'US':
            pronoun_count += 1
            
    # Avg Word Length
    sum_chars = sum(len(word) for word in cleaned_tokens_raw)
    avg_word_length = sum_chars / len(cleaned_tokens_raw) if len(cleaned_tokens_raw) > 0 else 0
    
    return [positive_score, negative_score, polarity_score, subjectivity_score, 
            avg_sentence_length, percentage_complex_words, fog_index, 
            avg_words_per_sentence, complex_word_count, word_count, 
            syllable_per_word, pronoun_count, avg_word_length]

# --- MAIN EXECUTION ---
def main():
    print("Loading resources...")
    stop_words = load_stop_words()
    pos_dict, neg_dict = load_master_dictionary(stop_words)
    
    print("Reading Input...")
    try:
        input_df = pd.read_excel(INPUT_FILE)
    except Exception as e:
        print(f"Error reading {INPUT_FILE}: {e}")
        return

    output_data = []
    
    columns = ['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 
               'SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 
               'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 
               'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']

    for index, row in input_df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        file_path = os.path.join(TEXT_DIR, f"{url_id}.txt")
        
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
            scores = analyze_text(text, stop_words, pos_dict, neg_dict)
            output_data.append([url_id, url] + scores)
            print(f"Analyzed {url_id}")
        else:
            print(f"File missing: {url_id}")
            output_data.append([url_id, url] + [0]*13)

    output_df = pd.DataFrame(output_data, columns=columns)
    output_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Success! Output saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()