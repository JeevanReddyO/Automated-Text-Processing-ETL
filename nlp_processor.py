import pandas as pd
import os
import re
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize

# --- PRE-FLIGHT CHECK ---
def init_nltk():
    """Ensures necessary NLTK datasets are downloaded."""
    resources = ['punkt', 'punkt_tab']
    for r in resources:
        try:
            nltk.download(r, quiet=True)
        except:
            pass
init_nltk()

# --- CONSTANTS ---
DIR_STOPWORDS = 'StopWords'
DIR_DICTIONARY = 'MasterDictionary'
DIR_TEXTS = 'extracted_articles'
FILE_INPUT = 'Input.xlsx'
FILE_OUTPUT = 'Output_Data_Structure_Final.xlsx'

class TextAnalyzer:
    def __init__(self):
        self.stop_words = set()
        self.pos_words = set()
        self.neg_words = set()
        self.load_resources()

    def load_resources(self):
        """Loads Stop Words and Master Dictionary into memory."""
        print("Loading NLP dictionaries...")
        
        # 1. Load Stop Words
        if os.path.exists(DIR_STOPWORDS):
            for fname in os.listdir(DIR_STOPWORDS):
                fpath = os.path.join(DIR_STOPWORDS, fname)
                try:
                    with open(fpath, 'r', encoding='latin-1') as f:
                        lines = f.read().splitlines()
                        for line in lines:
                            # Handle currency format "USD|Dollar"
                            term = line.split('|')[0].strip().lower()
                            self.stop_words.add(term)
                except Exception as e:
                    print(f"Warning: Failed to read {fname}")

        # 2. Load Master Dictionary
        for cat_file, target_set in [('positive-words.txt', self.pos_words), 
                                     ('negative-words.txt', self.neg_words)]:
            path = os.path.join(DIR_DICTIONARY, cat_file)
            if os.path.exists(path):
                with open(path, 'r', encoding='latin-1') as f:
                    words = f.read().splitlines()
                    for w in words:
                        if w not in self.stop_words:
                            target_set.add(w.strip().lower())

    def get_syllable_count(self, word):
        """Counts syllables based on vowels, handling es/ed endings."""
        word = word.lower()
        if word.endswith(('es', 'ed')):
            return 0
        
        vowels = "aeiouy"
        count = 0
        if word[0] in vowels:
            count += 1
        for i in range(1, len(word)):
            if word[i] in vowels and word[i-1] not in vowels:
                count += 1
        if word.endswith('e'):
            count -= 1
        return max(1, count)

    def process_text(self, raw_text):
        """Main logic to calculate all metrics."""
        # Tokenization
        sentences = sent_tokenize(raw_text)
        tokens = word_tokenize(raw_text)
        
        # Filtering
        # Keep only alphabetic words for metrics
        words_alpha = [t.lower() for t in tokens if t.isalpha()]
        # [cite_start]Remove stop words for sentiment analysis [cite: 18-19]
        words_clean = [w for w in words_alpha if w not in self.stop_words]

        # [cite_start]--- A. Sentiment Metrics --- [cite: 25-34]
        score_pos = sum(1 for w in words_clean if w in self.pos_words)
        score_neg = sum(1 for w in words_clean if w in self.neg_words)
        
        # Polarity = (Pos - Neg) / (Pos + Neg + 0.000001)
        score_pol = (score_pos - score_neg) / ((score_pos + score_neg) + 0.000001)
        
        # Subjectivity = (Pos + Neg) / (Total Clean Words + 0.000001)
        score_sub = (score_pos + score_neg) / (len(words_clean) + 0.000001)

        # [cite_start]--- B. Readability Metrics --- [cite: 36-40]
        count_words = len(words_alpha)
        count_sents = len(sentences)
        avg_sent_len = count_words / count_sents if count_sents > 0 else 0
        
        complex_list = [w for w in words_alpha if self.get_syllable_count(w) > 2]
        percent_complex = len(complex_list) / count_words if count_words > 0 else 0
        
        fog_idx = 0.4 * (avg_sent_len + percent_complex)

        # --- C. Other Stats ---
        # [cite_start]Complex Word Count [cite: 44]
        count_complex = len(complex_list)
        
        # [cite_start]Word Count (Cleaned) [cite: 46]
        final_word_count = len(words_clean)
        
        # [cite_start]Syllable Per Word [cite: 50]
        total_syllables = sum(self.get_syllable_count(w) for w in words_alpha)
        avg_syllables = total_syllables / count_words if count_words > 0 else 0

        # [cite_start]Personal Pronouns [cite: 53]
        # Regex to find I, we, my, ours, us (boundaries ensured). Exclude US.
        pronoun_pat = re.compile(r'\b(I|we|my|ours|us)\b', re.IGNORECASE)
        pronoun_count = 0
        for t in tokens:
            if pronoun_pat.match(t) and t != 'US':
                pronoun_count += 1
        
        # [cite_start]Avg Word Length [cite: 56]
        char_count = sum(len(w) for w in words_alpha)
        avg_word_len = char_count / count_words if count_words > 0 else 0

        return [score_pos, score_neg, score_pol, score_sub, 
                avg_sent_len, percent_complex, fog_idx, 
                avg_sent_len, count_complex, final_word_count, 
                avg_syllables, pronoun_count, avg_word_len]

    def execute_analysis(self):
        print(f"Reading {FILE_INPUT}...")
        try:
            df = pd.read_excel(FILE_INPUT)
        except Exception as e:
            print(f"Fatal Error: {e}")
            return

        results = []
        # Define output columns strictly as required
        cols = ['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 
                'SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 
                'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 
                'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']

        for row in df.itertuples():
            uid = row.URL_ID
            url = row.URL
            file_path = os.path.join(DIR_TEXTS, f"{uid}.txt")
            
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    text_data = f.read()
                metrics = self.process_text(text_data)
                results.append([uid, url] + metrics)
                print(f"Analyzed {uid}")
            else:
                print(f"Missing file for {uid}, filling zeros.")
                results.append([uid, url] + [0]*13)

        # Save
        out_df = pd.DataFrame(results, columns=cols)
        out_df.to_excel(FILE_OUTPUT, index=False)
        print(f"Done! Results saved to {FILE_OUTPUT}")

if __name__ == "__main__":
    engine = TextAnalyzer()
    engine.execute_analysis()