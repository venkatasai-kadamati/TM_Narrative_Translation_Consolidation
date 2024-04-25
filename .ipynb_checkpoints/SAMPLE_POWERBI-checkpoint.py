# Import necessary libraries
import pandas as pd
import numpy as np
from scipy import stats
import os

# # Clear all variables
# %reset -f

# Provide name for report output
report_name = 'CNB Report Narratives - BTL new.doc'

# Specify location of the tuning tracker
fp = 'C:/Users/SprongJ/OneDrive - Crowe LLP/CNB Tuning 2023/Tuning/R Scripts/Tuning Report Narratives Script/'
file_name = 'Production BTL Tuning Tracker - With Calculations.xlsx'

# Specify how to format the different parameters seen (change threshold names as needed)
currencyF = ['Minimal Sum','Minimum Value','Minimal Transaction Amount','Sum Lower Bound', 'Minimal Current Month Sum', 'Minimal Transaction Value', 'Transaction Amount Lower Bound', 'Sum Amount Lower Bound'] # Values should be formatted as money
numberF = ['No. of Occurrences','Minimum Volume', 'Min Value'] # Values should be formatted as integers
percentF = ['Ratio Lower Bound','Ratio Upper Bound'] # Values should be formatted as percentage
decimalF = ['STDEV exceeds Historical Average Sum', 'STDEV exceeds Historical Average Count'] # Values should be formatted to 2 decimals

# Load data
data = pd.read_excel(os.path.join(fp, file_name), sheet_name = 0)

# Enter the column name of the Rule IDs, if different
ruleIDs = data['Rule ID'].value_counts()

# Create lookup for values below 10 that need to be written out (no change)
numbers = ['0','1','2','3','4','5','6','7','8','9']
alpha_numbers = ['zero (0)', 'one (1)', 'two (2)', 'three (3)', 'four (4)', 'five (5)', 'six (6)', 'seven (7)', 'eight (8)', 'nine (9)']
alpha_numbers_cap = ['Zero (0)', 'One (1)', 'Two (2)', 'Three (3)', 'Four (4)', 'Five (5)', 'Six (6)', 'Seven (7)', 'Eight (8)', 'Nine (9)']
numbers_df = pd.DataFrame({'numbers': numbers, 'alpha_numbers': alpha_numbers, 'alpha_numbers_cap': alpha_numbers_cap})

# Create an empty data frame to hold narratives
narratives = pd.DataFrame(columns = ['col1', 'col2', 'col3', 'col4', 'col5'])


# ___________________________________________________________


