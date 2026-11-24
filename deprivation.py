'''
I think the best way for this to work is to have a separate script that takes 
an input (FROM A csv OR excel ?) of a single column with postcodes or LSOAs, 
converts to dataframe and then matches to LSOA and then to IMD.
The output will be a csv with the original input and the IMD fields added as 
separate columns.
'''

# Things to do:
# Clean input for postcode - decide upon whitespace handling and case sensitivity
# Allow for input of LSOA also
# Handle cases where input is neither a valid postcode nor an LSOA
# Specify which IMD version (2019, 2025)
# Read in IMD data for the specified version - this will involve reading in the IMD data file and storing it in 
# a suitable data structure for quick lookup.
# Map postcode to LSOA using ONS Postcode Directory
# Use LSOA to get IMD
# Add IMD fields to final output as separate csv
# Add to pipeline as a separate flag to run deprivation analysis
# Make it such that any error in postcodes get flagged but do not completely block the pipeline

#make it such that the postcode lookup file can be updated to newer versions when they are published
#and the code will still work without needing to change the code itself.

# mapping file location: \mapping_files\PCD_OA21_LSOA21_MSOA21_LAD_FEB25_UK_LU.zip
# bear in mind this is a zip file
# this can be updated to newer versions when they are published

# Columns that we need from the mapping file:
# PCD (postcode) - comes in different field lengths 7/8/9 with different whitespace, so we match correctly
# LSOA21CD (LSOA code) - this is what we need to get the IMD
# Different postcodes have different lengths too, this needs to be taken into account when matching to the mapping file.

'''
Docstring for deprivation.py
'''

import sys
import re
from typing import Optional
import pandas as pd
from pathlib import Path
from datetime import date

# Read in flattened test data


# Read in the mapping file for postcode to LSOA mapping
#path = \mapping_files\deprivation_indices\PCD_OA21_LSOA21_MSOA21_LAD_FEB25_UK_LU.csv


# All paths are resolved relative to this script's folder, so the pipeline
# works no matter which folder the terminal is in.
BASE_DIR = Path(__file__).resolve().parent
TEST_DATA_DIR = BASE_DIR / 'evaluation_data/raw/'
FLAT_DATA_DIR = BASE_DIR / 'evaluation_data/flattened/'
MAPPING_DIR = BASE_DIR / 'mapping_files'
DEPRIVATION_MAPPING_DIR = MAPPING_DIR / 'deprivation_indices'

df_test_data = pd.read_csv(FLAT_DATA_DIR / 'TEST_GP00.csv')
print(df_test_data.columns)

df_PC_LSOA_mapping = pd.read_csv(DEPRIVATION_MAPPING_DIR / 'PCD_OA21_LSOA21_MSOA21_LAD_FEB25_UK_LU.csv')
df_LSOA_IMD_mapping = pd.read_csv(DEPRIVATION_MAPPING_DIR / 'IoD2025_All_Ranks_Scores_Deciles_Population_Denominators.csv')

#print(df_PC_LSOA_mapping.head())

print(df_test_data['Postcode'].head())

#strip whitespace from the postcode column in the test data
df_test_data['Postcode'] = df_test_data['Postcode'].str.strip() #this gives nans and seems to give 8 character postcodes

#count number of characters in the postcode column in the test data
df_test_data['Postcode_length'] = df_test_data['Postcode'].str.len()

print(df_test_data['Postcode_length'].value_counts())

#find the LSOA for each postcode in the test data using the mapping file
#remove nan postcodes from the test data before mapping
df_test_data = df_test_data[df_test_data['Postcode'].notna()]

# Map postcodes to LSOAs based on their length
for length in [7, 8]:
    df_test_data.loc[df_test_data['Postcode_length'] == length, 'LSOA'] = df_test_data.loc[df_test_data['Postcode_length'] == length, 'Postcode'].map(df_PC_LSOA_mapping.set_index(f'pcd{length}')['lsoa21cd'])

print(df_test_data['LSOA'].head())

# match lsoa to LSOA code (2021) and IMD using other mapping file
df_test_data['IMD'] = df_test_data['LSOA'].map(df_LSOA_IMD_mapping.set_index('LSOA code (2021)')['Index of Multiple Deprivation (IMD) Score'])
df_test_data['IMD_decile'] = df_test_data['LSOA'].map(df_LSOA_IMD_mapping.set_index('LSOA code (2021)')['Index of Multiple Deprivation (IMD) Decile (where 1 is most deprived 10% of LSOAs)'])
print(df_test_data[['Postcode', 'LSOA', 'IMD', 'IMD_decile']].head())
