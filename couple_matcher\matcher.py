import pandas as pd
import numpy as np
import os

# Get the current working directory
current_directory = os.getcwd()

# Specify the subfolder and join it with the file name
subfolder = 'random_selector' 
file_name = 'staff.xlsx'
file_path = os.path.join(current_directory, subfolder, file_name)

# Define the sheet name and the column to use for coupling
sheet_name = 'Sheet1'  # Replace with the actual sheet name
coupling_column = 'LOCATION'  # Replace with the actual column name

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Separate males and females
males = df[df['SEX'] == 'Male']
females = df[df['SEX'] == 'Female']

# Group females and males by location and shuffle the order within each group
females = females.groupby(coupling_column).apply(lambda x: x.sample(frac=1)).reset_index(drop=True)
males = males.groupby(coupling_column).apply(lambda x: x.sample(frac=1)).reset_index(drop=True)

# Initialize an empty list to store file paths
output_file_paths = []

# Create a new DataFrame to store all couples
all_couples = pd.DataFrame(columns=['Female_FirstName', 'Female_Surname', 'Female_Location', 'Male_FirstName', 'Male_Surname', 'Male_Location'])

# Create a DataFrame to store unmatched individuals
unmatched_individuals = pd.DataFrame(columns=['FirstName', 'Surname', 'Location', 'Sex'])

# Iterate over each location
for location, females_loc, males_loc in zip(females[coupling_column].unique(), females.groupby(coupling_column), males.groupby(coupling_column)):
    couples_loc = pd.DataFrame(columns=['Female_FirstName', 'Female_Surname', 'Female_Location', 'Male_FirstName', 'Male_Surname', 'Male_Location'])
    
    # Iterate through males and females in the same location
    for i in range(min(len(females_loc[1]), len(males_loc[1]))):
        female = females_loc[1].iloc[i]
        male = males_loc[1].iloc[i]
        
        couples_loc = pd.concat([couples_loc, pd.DataFrame({
            'Female_FirstName': [female['FIRSTNAME']],
            'Female_Surname': [female['SURNAME']],
            'Female_Location': [location],
            'Male_FirstName': [male['FIRSTNAME']],
            'Male_Surname': [male['SURNAME']],
            'Male_Location': [location]
        })], ignore_index=True)

    # Append the couples_loc DataFrame to all_couples
    all_couples = pd.concat([all_couples, couples_loc], ignore_index=True)

    # Get unmatched individuals in this location
    unmatched_females_loc = females_loc[1].iloc[min(len(females_loc[1]), len(males_loc[1])):]
    unmatched_males_loc = males_loc[1].iloc[min(len(females_loc[1]), len(males_loc[1])):]

    # Append unmatched individuals to unmatched_individuals
    unmatched_individuals = pd.concat([unmatched_individuals, unmatched_females_loc, unmatched_males_loc], ignore_index=True)

# Sort the combined couples DataFrame by the location
all_couples = all_couples.sort_values(by='Female_Location').reset_index(drop=True)

# Write the result to a new Excel sheet for all couples
output_file_path_all_couples = os.path.join(current_directory, subfolder, f'output_all_couples_{sheet_name}.xlsx')
with pd.ExcelWriter(output_file_path_all_couples) as writer:
    all_couples.to_excel(writer, sheet_name='All_Couples', index=False)
    print(f"output_all_couples_{sheet_name}.xlsx has been successfully populated")

# Sort the unmatched individuals DataFrame by location
unmatched_individuals = unmatched_individuals.sort_values(by='LOCATION').reset_index(drop=True)

# Write the result to a new Excel sheet for unmatched individuals
output_file_path_unmatched = os.path.join(current_directory, subfolder, f'unmatched_individuals_{sheet_name}.xlsx')
with pd.ExcelWriter(output_file_path_unmatched) as writer:
    unmatched_individuals.to_excel(writer, sheet_name='Unmatched_Individuals', index=False)
    print(f"unmatched_individuals_{sheet_name}.xlsx has been successfully populated")
