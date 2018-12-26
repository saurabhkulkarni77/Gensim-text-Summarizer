from gensim.summarization import summarize 
import logging 
import pandas as pd
import re
import csv
import os
import nltk
logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.INFO) 
file_name = 'activity_responses.xls'
question_number = input('Please share the question number you want summary for [e.g, Q1 for question1] please use uppercase Q : ')
#print(question_number)

def text_clean(file_name, question_number):
	df = pd.read_excel(file_name)
	#print(df.head())
	text = df[question_number].dropna()
	#print("Before",text)
	excel_clean = re.sub("^\d+\s|\s\d+\s|\s\d+$", " ", str(text))
	s = re.sub("\n","",str(excel_clean))
	r = re.sub("\\d","",str(s))
	cleaned_text = re.sub("\\n","",str(r))
	give_summary_from_gensim = summarize(cleaned_text, ratio = 0.2)
	print("generating summary.........")
	print("Summary::::::::",give_summary_from_gensim)
	print("Also available on summary_output.xls file")
	
	with open("summary_output.xls","w") as f:
		file = f.write(give_summary_from_gensim)
	print("check summary_output.xls")
        
#text_clean(file_name, question_number)
if __name__ == "__main__":
    text_clean(file_name, question_number)


