import pandas as pd
import re
import logging
import os
from transformers import AutoTokenizer, AutoModel
import torch
from sklearn.metrics.pairwise import cosine_similarity
import json

# Load LaBSE model and tokenizer
tokenizer = AutoTokenizer.from_pretrained('sentence-transformers/LaBSE')
model = AutoModel.from_pretrained('sentence-transformers/LaBSE')

# Function to generate embeddings
def generate_embeddings(names):
    inputs = tokenizer(names, return_tensors='pt', padding=True, truncation=True)
    with torch.no_grad():
        embeddings = model(**inputs).last_hidden_state.mean(dim=1)
    return embeddings

# Function to separate male and female students
def separate_names(file_path):
    try:
        # Read all sheets from the Excel file
        sheets = pd.read_excel(file_path, sheet_name=None)
    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
        return [], []
    except Exception as e:
        logging.error(f"An error occurred while reading the Excel file: {e}")
        return [], []

    male_students = []
    female_students = []

    # Iterate through each sheet
    for sheet_name, df in sheets.items():
        # Check if the DataFrame has 'Student Name' and 'Gender' columns
        if 'Student Name' not in df.columns or 'Gender' not in df.columns:
            logging.error(f"The sheet '{sheet_name}' must contain 'Student Name' and 'Gender' columns.")
            continue
        
        # Ensure case insensitivity and strip whitespace
        df['Gender'] = df['Gender'].str.lower().str.strip().replace({'m': 'male', 'f': 'female'})
        
        # Separate male and female students
        males = df[df['Gender'] == 'male']['Student Name'].tolist()
        females = df[df['Gender'] == 'female']['Student Name'].tolist()
        
        # Append each male student to the male_students list
        male_students.extend(males)
        
        # Append each female student to the female_students list
        female_students.extend(females)
    
    return male_students, female_students

# Generate embeddings
file_path = r"C:\Users\Tevin\Links\test_files.xlsx"
male_students, female_students = separate_names(file_path)

# Debug: Print the lists of male and female students
print("Male Students:", male_students)
print("Female Students:", female_students)

male_embeddings = generate_embeddings(male_students)
female_embeddings = generate_embeddings(female_students)

# Debug: Print the shape of the embeddings
print("Male Embeddings Shape:", male_embeddings.shape)
print("Female Embeddings Shape:", female_embeddings.shape)

# Compute cosine similarity
similarity_matrix = cosine_similarity(male_embeddings, female_embeddings)

# Debug: Print the similarity matrix
print("Similarity Matrix:", similarity_matrix)

# Filter results with at least 50% similarity
threshold = 0.5
results = []

for i, male_name in enumerate(male_students):
    for j, female_name in enumerate(female_students):
        similarity = similarity_matrix[i][j]
        # Debug: Print each similarity score
        print(f"Similarity between {male_name} and {female_name}: {similarity}")
        if similarity >= threshold:
            results.append({
                "male_name": male_name,
                "female_name": female_name,
                "similarity": float(similarity)  # Convert to standard Python float
            })

# Debug: Print the results
print("Filtered Results:", results)

# Save results to JSON file
with open('similarity_results.json', 'w') as f:
    json.dump(results, f, indent=4)

print("Similarity results saved to similarity_results.json")