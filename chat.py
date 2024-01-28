import os
import string
from nltk.tokenize import sent_tokenize,word_tokenize
import re
import pandas as pd
import nltk
nltk.download('punkt')


def load_stopwords(folder_path):
    stopwords_list = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r') as file:
            for line in file:
                line = line.strip()
                if '|' in line:
                    words = [word.strip().lower() for word in line.split('|')]
                    stopwords_list.extend(words)
                else:
                    stopwords_list.append(line.lower())
    return set(stopwords_list)
    
def load_master_dictionary(file_path):
    with open(file_path, 'r') as file:
        words = [line.strip() for line in file.readlines()]
    return set(words)

def clean_text(text, stop_words):
    lower_case = text.lower()
    cleaned_text = lower_case.translate(str.maketrans('', '', string.punctuation))
   
    tokens = word_tokenize(cleaned_text)
    filtered_tokens = [word for word in tokens if word not in stop_words]
    word_count = len(filtered_tokens)
    return filtered_tokens,word_count

def calculate_scores(filtered_tokens, positive_dict, negative_dict):
    positive_score = sum(1 for word in filtered_tokens if word in positive_dict)
    negative_score = sum(1 for word in filtered_tokens if word in negative_dict)
    
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score)/ (len(filtered_tokens) + 0.000001)

    return positive_score, negative_score, polarity_score, subjectivity_score

def calculate_gunning_fog_index(text):
    # Tokenize sentences
    sentences = sent_tokenize(text)

    # Tokenize words
    words = word_tokenize(text)

    # Calculate average sentence length
    sen_len = len(sentences)
    if(sen_len == 0):
        sen_len = 1
    avg_sentence_length = len(words) / sen_len

    # Count the number of complex words (words with more than 3 syllables)
    complex_words = [word for word in words if syllable_count(word) > 2]
    complex_word_count = len(complex_words)
    word_len = len(words)
    if(word_len == 0):
        word_len = 1
    percentage_complex_words = len(complex_words) / word_len * 100

    # Calculate Gunning Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    return avg_sentence_length, percentage_complex_words,fog_index,complex_word_count
# Set your folder paths

def syllable_count(word):
    # Simple syllable count approximation
    vowels = "aeiouy"
    count = 0
    prev_char = ' '

    for char in word:
        char = char.lower()
        if char in vowels and prev_char not in vowels:
            count += 1
        prev_char = char

    # Adjust for certain word endings
    if word.endswith(('es', 'ed')) and len(word) > 2 and word[-3] not in vowels:
        count -= 1

    return max(1, count)

def calculate_average_words_per_sentence(text):
    # Tokenize sentences
    sentences = sent_tokenize(text)

    # Tokenize words
    words = word_tokenize(text)
    total_char_count = sum(len(word) for word in words)
    # Calculate average number of words per sentence
    word_len = len(words)
    if(word_len == 0):
        word_len = 1
    avg_words_per_sentence = len(words) / len(sentences) if len(sentences) > 0 else 0
    avg_word_length = total_char_count/word_len
    return avg_words_per_sentence, avg_word_length

def count_personal_pronouns(text):
    # Define a regex pattern for personal pronouns
    pronoun_pattern = re.compile(r'\b(?:I|we|my|ours|us)\b', flags=re.IGNORECASE)

    # Find all matches in the text
    pronoun_matches = pronoun_pattern.findall(text)
    pronoun_matches = [pronoun for pronoun in pronoun_matches if pronoun != 'US']
    # Count the occurrences
    pronoun_count = len(pronoun_matches)

    return pronoun_count

stopwords_folder = 'StopWords'
master_dict_folder = 'MasterDictionary'

# Load stopwords
stop_words = load_stopwords(stopwords_folder)


# Load positive and negative dictionaries

positive_dict = load_master_dictionary(os.path.join(master_dict_folder, 'positive-words.txt'))
negative_dict = load_master_dictionary(os.path.join(master_dict_folder, 'negative-words.txt'))
index = 0
# Clean the text
output_folder = 'output'
excel_file_path =  r"Output Data Structure.xlsx"
df = pd.read_excel(excel_file_path, engine='openpyxl')
output_excel_path = r'Output Data Structure.xlsx'
index=0
for file_name in os.listdir('output'):
    file_path = os.path.join('output', file_name)
    
    # Read the content of the file
    with open(file_path, 'r', encoding='utf-8') as file:
        financial_text = file.read()

 

    # Clean the text
    cleaned_tokens,word_count = clean_text(financial_text, stop_words)


    # Calculate scores
    pos_score, neg_score, pol_score, subj_score = calculate_scores(cleaned_tokens , positive_dict, negative_dict)
    avg_words_per_sentence,avg_word_length = calculate_average_words_per_sentence(financial_text)
    avg_sentence_length, percentage_complex_words, fog_index ,complex_word_count= calculate_gunning_fog_index(financial_text)
    syllables = syllable_count(financial_text)
    pronouns = count_personal_pronouns(financial_text)

   
    # Access and update values in each column
    df.at[index, 'POSITIVE SCORE'] = pos_score
    df.at[index, 'NEGATIVE SCORE'] = neg_score
    df.at[index, 'POLARITY SCORE'] = pol_score
    df.at[index, 'SUBJECTIVITY SCORE'] = subj_score
    df.at[index, 'AVG SENTENCE LENGTH'] = avg_sentence_length
    df.at[index, 'PERCENTAGE OF COMPLEX WORDS'] = percentage_complex_words
    df.at[index, 'FOG INDEX'] = fog_index
    df.at[index, 'AVG NUMBER OF WORDS PER SENTENCE'] = avg_words_per_sentence
    df.at[index, 'COMPLEX WORD COUNT'] = complex_word_count
    df.at[index, 'WORD COUNT'] = word_count
    df.at[index, 'SYLLABLE PER WORD'] = syllables
    df.at[index, 'PERSONAL PRONOUNS'] = pronouns
    df.at[index, 'AVG WORD LENGTH'] = avg_word_length

    df.to_excel(output_excel_path, index=False,engine = 'openpyxl')
    index = index+1


# Load the DataFrame from the Excel file
df = pd.read_excel(output_excel_path, engine='openpyxl')

# Remove blank columns (columns with all NaN values)
df_cleaned = df.dropna(axis=1, how='all')

# Save the cleaned DataFrame back to Excel
df_cleaned.to_excel(output_excel_path, index=False, engine='openpyxl')

print('Complete')
