import os
import io
import pandas as pd
import re
import nltk
from nltk.corpus import stopwords



main_dir = r"C:\Users\Rithick\Desktop\Codes\python"
directory_path = r"C:\Users\Rithick\Desktop\Codes\python\extracted_articles"
stop_words_dictory = r"C:\Users\Rithick\Desktop\Codes\python\StopWords"
master_dict_directory = r"C:\Users\Rithick\Desktop\Codes\python\MasterDictionary"

stop_words_list =[]
positive_list = []
negative_list = []

output_file = "output.xlsx"

df = pd.read_excel(output_file)

for filename3 in os.listdir(master_dict_directory):
    file_path = os.path.join(master_dict_directory, filename3)
    with io.open(file_path, "r") as file:
        if(filename3 == "negative-words.txt"):
            text = file.read()
            negative_list = text.split()
        elif(filename3 == "positive-words.txt"):
            text= file.read()
            positive_list = text.split()

for filename2 in os.listdir(stop_words_dictory):
    file_path = os.path.join(stop_words_dictory, filename2)
    with io.open(file_path, "r") as file:
        text = file.read()
        temp_list = text.split()
    stop_words_list.extend(temp_list)



for filename in os.listdir(directory_path):
    pos_score = 0
    neg_score = 0
    for index, row in df.iterrows():   
        if(str(row['URL_ID'])+".txt" == filename):
            break 
    if filename.endswith(".txt"):
        file_path = os.path.join(directory_path, filename)
        with io.open(file_path, "r", encoding="utf-8") as file:
            text = file.read()
            cleaned_text = ''.join([char if ord(char) < 128 else ' ' for char in text])
            cleaned_text = cleaned_text.split()            
            more_cleaned = ""
            
            #1
                       
            for i in range(0, len(cleaned_text)):
                if (cleaned_text[i]=='We' and cleaned_text[i+1]=='provide' and cleaned_text[i+2]=='intelligence,' and cleaned_text[i+3]=='accelerate'):
                    break
                if cleaned_text[i] not in stop_words_list:               
                    more_cleaned += cleaned_text[i]
                    more_cleaned += " "            
                  
            count = 0
            for i in more_cleaned.split():
                count +=1
                if i in positive_list:
                    pos_score+=1
                elif i in negative_list:
                    neg_score +=1       
            polarity = (pos_score - neg_score)/((pos_score + neg_score + 0.000001))       
            subjectitivity = (pos_score+neg_score)/(count + 0.000001)
            
            # 2
            count = 0
            vowels = ['a', 'e', 'i', 'o', 'u', 'A', 'E', 'I', 'O', 'U']
            complex_Count = 0
            for i in range(0, len(cleaned_text)):
                if (cleaned_text[i]=='We' and cleaned_text[i+1]=='provide' and cleaned_text[i+2]=='intelligence,' and cleaned_text[i+3]=='accelerate'):
                    break               
                more_cleaned += cleaned_text[i]
                more_cleaned += " "
                count+=1
                if len((set(list(cleaned_text[i])) & set(vowels)))>=2 and (cleaned_text[i][-1:-3:-1]!= 'ed' or cleaned_text[i][-1:-3:-1]!= 'es') :
                    complex_Count+=1
            # print(complex_Count)
            num_senten = 0
            for i in more_cleaned.split("."):
                num_senten+=1
            stop_words = set(stopwords.words('english'))
            more_cleaned_nltk = ''
            
            length_more_cleaned_nltk = 0
            for i in more_cleaned.split():
                if (i in stop_words):
                    continue
                if (i in ['!', '?', '.', ',']):
                    continue
                else:
                    more_cleaned_nltk += i
                    more_cleaned_nltk += " "
                    length_more_cleaned_nltk+=1
            tot_Syll_count = 0
            for i in more_cleaned.split():
                if (i[-1:-3:-1]!='ed' and i[-1:-3:-1]!='es'):
                    list_i = list(i)
                    final_list = list(set(list_i) & set(vowels))
                    if (len(final_list)>=1):
                        tot_Syll_count += len(final_list)+1
                        
            pronoun_regex = r'\b(?:I|we|my|ours|us)(?![\w.])\b'
            pronoun_matches = re.findall(pronoun_regex, more_cleaned, flags=re.IGNORECASE)
            
            num_of_charac = len(more_cleaned)
            
            avg_word_len = num_of_charac/count
            
            num_pronouns = len(pronoun_matches)
            syll_count_per_word = tot_Syll_count/count
            avg_senten_length = count/num_senten
            percent_complex = complex_Count/count
            fog_index = 0.4 * (avg_senten_length+percent_complex)
            
                
            
            
            
                
                        
                
    df.at[index, 'POSITIVE SCORE'] = pos_score
    df.at[index, 'NEGATIVE SCORE'] = neg_score
    df.at[index, 'POLARITY SCORE'] = polarity
    df.at[index, 'SUBJECTIVITY SCORE'] = subjectitivity
    df.at[index, 'AVG SENTENCE LENGTH'] = avg_senten_length
    df.at[index, 'PERCENTAGE OF COMPLEX WORDS'] = percent_complex
    df.at[index, 'FOG INDEX'] = fog_index
    df.at[index, 'AVG NUMBER OF WORDS PER SENTENCE'] = avg_senten_length
    df.at[index, 'COMPLEX WORD COUNT'] = complex_Count
    df.at[index, 'WORD COUNT'] = length_more_cleaned_nltk
    df.at[index, 'SYLLABLE PER WORD'] = syll_count_per_word
    df.at[index, 'PERSONAL PRONOUNS'] = num_pronouns
    df.at[index, 'AVG WORD LENGTH'] = avg_word_len
    
    
      
        
        
        
excel_file_path = 'output.xlsx'
sheet_name = 'Sheet1'

# Write the DataFrame to the Excel sheet
df.to_excel(excel_file_path, sheet_name=sheet_name, index=False)
print(f'DataFrame written to {excel_file_path} in sheet "{sheet_name}"')
            
                    
            
            
            
                    
            
            
            
        
            
