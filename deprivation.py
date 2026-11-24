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


