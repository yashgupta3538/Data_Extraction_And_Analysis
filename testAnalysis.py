import os
import nltk
import pandas as pd
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import cmudict
import re

# Load stop words
def load_stop_words(file_path):
    with open(file_path, 'r') as file:
        stop_words = set(line.strip() for line in file)
    return stop_words

def calculate_readability_analysis(text):
    words = word_tokenize(text)
    sentences = sent_tokenize(text)
    total_words = len(words)
    total_sentences = len(sentences)
    avg_sentence_length = total_words / total_sentences
    
    # Count complex words
    cmu_dict = cmudict.dict()
    complex_word_count = 0
    for word in words:
        syllable_count = syllable_count_per_word(word, cmu_dict)
        if syllable_count > 2:
            complex_word_count += 1
    
    percentage_complex_words = (complex_word_count / total_words) * 100
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    avg_words_per_sentence = total_words / total_sentences
    syllable_per_word = sum(syllable_count_per_word(word, cmu_dict) for word in words) / total_words
    
    personal_pronouns = len(re.findall(r'\b(I|me|my|mine|we|us|our|ours|you|your|yours)\b', text, re.IGNORECASE))
    
    avg_word_length = sum(len(word) for word in words) / total_words
    
    return avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, total_words, syllable_per_word, personal_pronouns, avg_word_length

def syllable_count_per_word(word, cmu_dict):
    if word.lower() in cmu_dict:
        return max([len(list(y for y in x if y[-1].isdigit())) for x in cmu_dict[word.lower()]])
    return 0

def calculate_sentiment_scores(positive_count, negative_count, total_words):
    positive_score = positive_count
    negative_score = negative_count
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score

def load_stop_words_from_folder(folder_path):
    stop_words = {}
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r') as file:
            stop_words[file_name.split('.')[0]] = set(file.read().splitlines())
    return stop_words

local_folder_path = "C:\\Users\\DELL\\OneDrive\\Desktop\\Pyhton\\stop_words"
stop_words = load_stop_words_from_folder(local_folder_path)

def load_master_dictionary(folder_path):
    master_dictionaries = {}
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        with open(file_path) as file:
            master_dict = {}
            for line in file:
                # Check if the line contains at least two values
                if ',' in line:
                    word, sentiment = line.strip().split(',', 1)
                    master_dict[word.lower()] = sentiment
            master_dictionaries[file_name] = master_dict
    return master_dictionaries

master_dictionaries = load_master_dictionary("master_dictionary_file")

# Access each dictionary separately
with open("C:\\Users\\DELL\\OneDrive\\Desktop\\Pyhton\\master_dictionary_file\\positive-words.txt", 'r') as pos_file:
    positive_dict = pos_file.read().splitlines()
with open("C:\\Users\\DELL\\OneDrive\\Desktop\\Pyhton\\master_dictionary_file\\negative-words.txt", 'r') as neg_file:
    negative_dict = neg_file.read().splitlines()

# Tokenize text
def tokenize_text(text):
    return word_tokenize(text.lower())

# Remove stop words
def remove_stop_words(tokens, stop_words):
    return [token for token in tokens if token not in stop_words]

# Count positive and negative words
def count_sentiment_words(tokens, positive_dict, negative_dict):
    negative_word_count = sum(1 for word in tokens if word in negative_dict)
    positive_word_count = sum(1 for word in tokens if word in positive_dict)
    return positive_word_count, negative_word_count

# Perform analysis on each file
results = []
files_folder = "C:\\Users\\DELL\\OneDrive\\Desktop\\Pyhton\\extract"

# Load positive and negative dictionaries
with open("C:\\Users\\DELL\\OneDrive\\Desktop\\Pyhton\\master_dictionary_file\\positive-words.txt", 'r') as pos_file:
    positive_dict = set(pos_file.read().splitlines())
with open("C:\\Users\\DELL\\OneDrive\\Desktop\\Pyhton\\master_dictionary_file\\negative-words.txt", 'r') as neg_file:
    negative_dict = set(neg_file.read().splitlines())

for filename in os.listdir(files_folder):
    if filename.endswith('.txt'):
        file_path = os.path.join(files_folder, filename)
        with open(file_path, 'r', encoding='utf-8') as file: 
            text = file.read()
            tokens = tokenize_text(text)
            cleaned_tokens = remove_stop_words(tokens, stop_words)
            positive_count, negative_count = count_sentiment_words(cleaned_tokens, positive_dict, negative_dict)
            total_words = len(cleaned_tokens)
            positive_score, negative_score, polarity_score, subjectivity_score = calculate_sentiment_scores(positive_count, negative_count, total_words)
            avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, word_count, syllable_per_word, personal_pronouns, avg_word_length = calculate_readability_analysis(text)
            result = {
                'Filename': filename,
                'Positive Score': positive_score,
                'Negative Score': negative_score,
                'Polarity Score': polarity_score,
                'Subjectivity Score': subjectivity_score,
                'Avg Sentence Length': avg_sentence_length,
                'Percentage of Complex Words': percentage_complex_words,
                'FOG Index': fog_index,
                'Avg Words Per Sentence': avg_words_per_sentence,
                'Complex Word Count': complex_word_count,
                'Word Count': word_count,
                'Syllable Per Word': syllable_per_word,
                'Personal Pronouns': personal_pronouns,
                'Avg Word Length': avg_word_length
            }
            results.append(result)

df = pd.DataFrame(results)

# Save DataFrame to Excel file
output_excel_file = "C:\\Users\\DELL\\OneDrive\\Desktop\\Pyhton\\output_data_structure.xlsx"
df.to_excel(output_excel_file, index=False)
print(f"Analysis results saved to your Excel File!!")