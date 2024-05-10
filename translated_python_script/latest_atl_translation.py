# Data Load ---------------------------------------------------------------
import os
import pandas as pd
import numpy as np
from numpy import number
from pandas import DataFrame

# Specify location of the tuning tracker
fp = "C:/Users/KadamatiV/OneDrive - Crowe LLP/Documents/PROJECTHUB/TM TUNING/TM_Narrative_Codebase_Consolidation/required_processing_data/"
file_name = 'Tuning Tracker - ATL Calculationsv11.xlsx'

report_name = 'output/CNB - Production Report Narratives.docx' # Enter the desired name for the report that will be generated (must end in '.doc')

# Specify how to format the different parameters seen (change threshold names as needed)
currencyF = ['Minimal Sum','Minimum Value','Minimal Transaction Amount','Sum Lower Bound', 'Min Value', 'Minimal Current Month Sum', 'Minimal Transaction Value', 'Transaction Amount Lower Bound'] # Values should be formatted as money
numberF = ['No. of Occurrences','Minimum Volume', 'Min Value'] # Values should be formatted as integers
percentF = ['Ratio Lower Bound','Ratio Upper Bound'] # Values should be formatted as percentage
decimalF = ['STDEV exceeds Historical Average Sum', 'STDEV exceeds Historical Average Count'] # Values should be formatted to 2 decimals

# Data is loaded
data = pd.read_excel(os.path.join(fp, file_name), sheet_name=1)

# Enter the column name of the Rule IDs, if different
ruleIDs = data['Rule ID'].value_counts().index

# Create lookup for values below 10 that need to be written out (no change)
numbers = ['0','1','2','3','4','5','6','7','8','9']
alpha_numbers = ['zero (0)', 'one (1)', 'two (2)', 'three (3)', 'four (4)', 'five (5)', 'six (6)', 'seven (7)', 'eight (8)', 'nine (9)']
alpha_numbers_cap = ['Zero (0)', 'One (1)', 'Two (2)', 'Three (3)', 'Four (4)', 'Five (5)', 'Six (6)', 'Seven (7)', 'Eight (8)', 'Nine (9)']
numbers_df = DataFrame({'numbers': numbers, 'alpha_numbers': alpha_numbers, 'alpha_numbers_cap': alpha_numbers_cap})

# Create an empty data frame to hold narratives
narratives = DataFrame(columns=['Rule ID', 'Rule Name', 'Population Group', 'Parameter', 'Summary', 'Analysis', 'Conclusion'])

# Begin Iterations --------------------------------------------------------

# Iterate over Rule IDs, Population Groups, and Parameters
for x in ruleIDs:
    data_Rule = data[data['Rule ID'] == x]
    popGroups = data_Rule['Population Group'].value_counts().index

    for pop in popGroups:
        data_Rule_Pop = data_Rule[data_Rule['Population Group'] == pop]
        params = data_Rule_Pop['Parameter'].value_counts().index

        for threshold in params:
            data_Rule_Pop_Parameter = data_Rule_Pop[data_Rule_Pop['Parameter'] == threshold]

            # Summary Paragraph -------------------------------------------------------
            ### Create paragraph to be populated before parameter analysis

            # Time frame of alerts that generated; if no alerts generated -> no analysis performed
            temp = data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0]
            date_range = data_Rule_Pop_Parameter['Date Range'].values[0].split('-')

            temp2 = data_Rule_Pop_Parameter['Parameter'].value_counts().index
            if len(temp2) == 1:
                params_used = temp2[0] + ' parameter'
            elif len(temp2) == 2:
                params_used = temp2[0] + ' and ' + temp2[1] + ' parameters'
            else:
                params_used = ', '.join(temp2[:-1]) + ', and ' + temp2[-1] + ' parameters'

            line_one = ''
            if temp == 0:
                line_one = f"No alerts generated in the Actimize environment for the {data_Rule_Pop_Parameter['Population Group'].values[0]} population group between {date_range[0]} and {date_range[1]}; therefore, no analysis was performed. The current thresholds are recommended to be maintained."
            else:
                line_one = f"Production tuning analysis was performed on the {params_used} for the {data_Rule_Pop_Pop['Population Group'].values[0]} population group."
            # _________________________________________ added phase 2 ___________________________
            # No. of rule breaks generated
            temp = data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0]

            line_two = ''
            if temp == 0:
                pass
            elif temp == 1:
                line_two = f"{numbers_df[numbers_df['numbers'] == str(temp)]['alpha_numbers'].values[0]} rule break was generated between {date_range[0]} and {date_range[1]}."
            elif temp < 10:
                line_two = f"{numbers_df[numbers_df['numbers'] == str(temp)]['alpha_numbers'].values[0]} rule breaks were generated between {date_range[0]} and {date_range[1]}."
            else:
                line_two = f"{temp:,} rule breaks were generated between {date_range[0]} and {date_range[1]}."

            # Data quality alerts identified, if any
            temp = data_Rule_Pop_Parameter['Data Quality Alerts'].values[0]

            line_three = ''
            if temp == 0:
                pass
            elif temp == 1:
                line_three = f"The Bank identified {numbers_df[numbers_df['numbers'] == str(temp)]['alpha_numbers'].values[0]} Data Quality rule break that generated on duplicated or incorrectly mapped transactional activity. As a result, this rule break was marked as Data Quality and excluded from analysis."
            elif temp < 10:
                line_three = f"The Bank identified {numbers_df[numbers_df['numbers'] == str(temp)]['alpha_numbers'].values[0]} Data Quality rule breaks that generated on duplicated or incorrectly mapped transactional activity. As a result, these rule breaks were marked as Data Quality and excluded from analysis."
            else:
                line_three = f"The Bank identified {temp:,} Data Quality rule breaks that generated on duplicated or incorrectly mapped transactional activity. As a result, these rule breaks were marked as Data Quality and excluded from analysis."

            # Interesting alerts / effectiveness
            temp_total = data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] - data_Rule_Pop_Parameter['Data Quality Alerts'].values[0]
            temp_sar = data_Rule_Pop_Parameter['SARs Filed'].values[0]
            temp_int = data_Rule_Pop_Parameter['Interesting Alerts'].values[0] + data_Rule_Pop_Parameter['SARs Filed'].values[0]
            temp_eff = round(data_Rule_Pop_Parameter['Effectiveness'].values[0], 2)
            temp_sar_yield = round(data_Rule_Pop_Parameter['SAR Yield'].values[0], 2)

            line_four = ''
            if temp_total == 0:
                pass
            elif temp_total == 1 and temp_int == 1 and temp_sar == 1:
                line_four = f"The one (1) reviwed rule break was determined to be Interesting and led to a SAR filing, resulting in a production effectiveness and SAR yield of 100.00%."
            elif temp_total == 1 and temp_int == 1 and temp_sar == 0:
                line_four = f"The one (1) reviwed rule break was determined to be Interesting but did not end in a SAR filing, resulting in a production effectiveness of 100.00% and a SAR yield of 0.00%."
            elif temp_total == 1 and temp_int != 1:
                line_four = f"The one (1) reviwed rule break was not determined to be Interesting, resulting in a production effectiveness of 0.00%."
            elif temp_total < 10 and temp_int == 1:
                line_four = f"Of the total {numbers_df[numbers_df['numbers'] == str(temp_total)]['alpha_numbers'].values[0]} reviewed rule breaks, {numbers_df[numbers_df['numbers'] == str(temp_int)]['alpha_numbers'].values[0]} rule break was determined to be Interesting, resulting in a production effectiveness of {temp_eff}%."
            elif temp_total < 10 and temp_int < 10:
                line_four = f"Of the total {numbers_df[numbers_df['numbers'] == str(temp_total)]['alpha_numbers'].values[0]} reviewed rule breaks, {numbers_df[numbers_df['numbers'] == str(temp_int)]['alpha_numbers'].values[0]} rule breaks were determined to be Interesting, resulting in a production effectiveness of {temp_eff}%."
            elif temp_total >= 10 and temp_int == 1:
                line_four = f"Of the total {temp_total:,} reviewed rule breaks, {numbers_df[numbers_df['numbers'] == str(temp_int)]['alpha_numbers'].values[0]} rule break was determined to be Interesting, resulting in a production effectiveness of {temp_eff}%."
            elif temp_total >= 10 and temp_int < 10:
                line_four = f"Of the total {temp_total:,} reviewed rule breaks, {numbers_df[numbers_df['numbers'] == str(temp_int)]['alpha_numbers'].values[0]} rule breaks were determined to be Interesting, resulting in a production effectiveness of {temp_eff}%."
            elif temp_total >= 10 and temp_int >= 10:
                line_four = f"Of the total {temp_total:,} reviewed rule breaks, {temp_int:,} rule breaks were determined to be Interesting, resulting in a production effectiveness of {temp_eff}%."

            # SARs filed / SAR yield (using temp variables from line_four)
            line_five = ''
            if temp_total == 0:
                pass
            elif temp_int == 0:
                pass
            elif temp_sar < 10 and temp_int < 10:
                line_five = f"{numbers_df[numbers_df['numbers'] == str(temp_sar)]['alpha_numbers_cap'].values[0]} of the {numbers_df[numbers_df['numbers'] == str(temp_int)]['alpha_numbers'].values[0]} Interesting rule breaks led to a SAR filing, resulting in a SAR yield of {temp_sar_yield}%. Additional detail on the analysis results per population group is provided below."
            elif temp_sar < 10 and temp_int >= 10:
                line_five = f"{numbers_df[numbers_df['numbers'] == str(temp_sar)]['alpha_numbers_cap'].values[0]} of the {temp_int:,} Interesting rule breaks led to a SAR filing, resulting in a SAR yield of {temp_sar_yield}%. Additional detail on the analysis results per population group is provided below."
            elif temp_sar >= 10:
                line_five = f"{temp_sar:,} of the {temp_int:,} Interesting rule breaks led to a SAR filing, resulting in a SAR yield of {temp_sar_yield}%. Additional detail on the analysis results per population group is provided below."

            # Analysis Paragraph ------------------------------------------------------
            ### Create paragraph to be populated under parameter analysis

            # Threshold tuned and value
            temp = data_Rule_Pop_Parameter['Current Threshold'].values[0]
            if threshold in currencyF:
                temp_formatted = f"${temp:,}"
            elif threshold in numberF:
                if temp < 10:
                    temp_formatted = numbers_df[numbers_df['numbers'] == str(temp)]['alpha_numbers'].values[0]
                else:
                    temp_formatted = f"{temp:,}"
            elif threshold in percentF:
                temp_formatted = f"{temp:.2%}"
            elif threshold in decimalF:
                temp_formatted = f"{temp:,.2f}"
            else:
                temp_formatted = temp

            line_six = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            else:
                line_six = f"Above-the-line tuning was conducted on the {data_Rule_Pop_Parameter['Parameter'].values[0]} threshold, which was set at a value of {temp_formatted}."

            # Values that rule breaks generated at
            temp = data_Rule_Pop_Parameter['Max Val'].values[0] - data_Rule_Pop_Parameter['Min Val'].values[0]

            if threshold in currencyF:
                val_formatted = [f"${data_Rule_Pop_Parameter['Min Val'].values[0]:,}", f"${data_Rule_Pop_Parameter['Max Val'].values[0]:,}"]
            elif threshold in numberF:
                if np.isnan(data_Rule_Pop_Parameter['Max Val'].values[0]):
                    val_formatted = ['0','0']
                elif data_Rule_Pop_Parameter['Max Val'].values[0] < 10:
                    val_formatted = [numbers_df[numbers_df['numbers'] == str(data_Rule_Pop_Parameter['Min Val'].values[0])]['alpha_numbers'].values[0], numbers_df[numbers_df['numbers'] == str(data_Rule_Pop_Parameter['Max Val'].values[0])]['alpha_numbers'].values[0]]
                elif data_Rule_Pop_Parameter['Max Val'].values[0] >= 10 and data_Rule_Pop_Parameter['Min Val'].values[0] < 10:
                    val_formatted = [numbers_df[numbers_df['numbers'] == str(data_Rule_Pop_Parameter['Min Val'].values[0])]['alpha_numbers'].values[0], f"{data_Rule_Pop_Parameter['Max Val'].values[0]:,}"]
                else:
                    val_formatted = [f"{data_Rule_Pop_Parameter['Min Val'].values[0]:,}", f"{data_Rule_Pop_Parameter['Max Val'].values[0]:,}"]
            elif threshold in percentF:
                val_formatted = [f"{data_Rule_Pop_Parameter['Min Val'].values[0]:.2%}", f"{data_Rule_Pop_Parameter['Max Val'].values[0]:.2%}"]
            elif threshold in decimalF:
                val_formatted = [f"{data_Rule_Pop_Parameter['Min Val'].values[0]:,.2f}", f"{data_Rule_Pop_Parameter['Max Val'].values[0]:,.2f}"]
            else:
                val_formatted = [data_Rule_Pop_Parameter['Min Val'].values[0], data_Rule_Pop_Parameter['Max Val'].values[0]]

            line_seven = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            elif temp == 0:
                line_seven = f"Rule breaks were generated solely at a value of {val_formatted[1]} within the {data_Rule_Pop_Parameter['Population Group'].values[0]} population segment."
            elif temp != 0:
                line_seven = f"Rule breaks were generated for values ranging between {val_formatted[0]} and {val_formatted[1]} within the {data_Rule_Pop_Parameter['Population Group'].values[0]} population segment."

            # No. of interesting rule breaks in rule population
            temp_sar = data_Rule_Pop_Parameter['SARs Filed'].values[0]
            temp_int = data_Rule_Pop_Parameter['Interesting Alerts'].values[0] + data_Rule_Pop_Parameter['SARs Filed'].values[0]
            temp_eff = data_Rule_Pop_Parameter['Effectiveness'].values[0]
            temp_sar_yield = data_Rule_Pop_Parameter['SAR Yield'].values[0]

            line_eight = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            elif temp_int == 1 and temp_sar == 1:
                line_eight = "One (1) Interesting rule break was noted in the production population which also resulted in a SAR filing."
            elif temp_int == 1 and temp_sar != 1:
                line_eight = "One (1) Interesting rule break was noted in the production population which did not result in a SAR filing."
            elif temp_int < 10 and temp_sar == 1:
                line_eight = f"{numbers_df[numbers_df['numbers'] == str(temp_int)]['alpha_numbers_cap'].values[0]} Interesting rule breaks were noted in the production population, of which one (1) rule break resulted in a SAR filing."
            elif temp_int < 10:
                line_eight = f"{numbers_df[numbers_df['numbers'] == str(temp_int)]['alpha_numbers_cap'].values[0]} Interesting rule breaks were noted in the production population, of which {numbers_df[numbers_df['numbers'] == str(temp_sar)]['alpha_numbers'].values[0]} rule breaks resulted in SAR filings."
            elif temp_int >= 10 and temp_sar == 1:
                line_eight = f"{temp_int:,} Interesting rule breaks were noted in the production population, of which one (1) rule break resulted in a SAR filing."
            elif temp_int >= 10 and temp_sar < 10:
                line_eight = f"{temp_int:,} Interesting rule breaks were noted in the production population, of which {numbers_df[numbers_df['numbers'] == str(temp_sar)]['alpha_numbers'].values[0]} rule breaks resulted in SAR filings."
            elif temp_int >= 10 and temp_sar >= 10:
                line_eight = f"{temp_int:,} Interesting rule breaks were noted in the production population, of which {temp_sar:,} rule breaks resulted in SAR filings."

            # Placeholder for recommendation rationale
            line_nine = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] != 0:
                line_nine = '###INSERT TUNING DECISION###'

            # Tuning recommendation
            temp = [data_Rule_Pop_Parameter['Current Threshold'].values[0], data_Rule_Pop_Parameter['Recommended Threshold'].values[0]]
            if threshold in currencyF:
                temp_formatted = [f"${temp[0]:,}", f"${temp[1]:,}"]
            elif threshold in numberF:
                if temp[0] < 10 and temp[1] < 10:
                    temp_formatted = [numbers_df[numbers_df['numbers'] == str(temp[0])]['alpha_numbers'].values[0], numbers_df[numbers_df['numbers'] == str(temp[1])]['alpha_numbers'].values[0]]
                elif temp[0] < 10 and temp[1] >= 10:
                    temp_formatted = [numbers_df[numbers_df['numbers'] == str(temp[0])]['alpha_numbers'].values[0], f"{temp[1]:,}"]
                else:
                    temp_formatted = [f"{temp[0]:,}", f"{temp[1]:,}"]
            elif threshold in percentF:
                temp_formatted = [f"{temp[0]:.2%}", f"{temp[1]:.2%}"]
            elif threshold in decimalF:
                temp_formatted = [f"{temp[0]:,.2f}", f"{temp[1]:,.2f}"]
            else:
                temp_formatted = temp

            line_ten = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            elif data_Rule_Pop_Parameter['Current Threshold'].values[0] == data_Rule_Pop_Parameter['Recommended Threshold'].values[0]:
                line_ten = f"Therefore, it is recommended to maintain the {data_Rule_Pop_Parameter['Parameter'].values[0]} threshold at {temp_formatted[0]}."
            elif data_Rule_Pop_Parameter['Current Threshold'].values[0] != data_Rule_Pop_Parameter['Recommended Threshold'].values[0]:
                line_ten = f"Therefore, it is recommended to increase the {data_Rule_Pop_Parameter['Parameter'].values[0]} threshold from {temp_formatted[0]} to {temp_formatted[1]}."

            # Change in effectiveness
            line_eleven = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0 or data_Rule_Pop_Parameter['Effectiveness'].values[0] > data_Rule_Pop_Parameter['Prop Effectiveness'].values[0]:
                pass
            elif data_Rule_Pop_Parameter['SAR Yield'].values[0] > data_Rule_Pop_Parameter['Prop SAR Yield'].values[0] and (data_Rule_Pop_Parameter['Current Threshold'].values[0] == data_Rule_Pop_Parameter['Recommended Threshold'].values[0] or data_Rule_Pop_Parameter['Effectiveness'].values[0] == data_Rule_Pop_Parameter['Prop Effectiveness'].values[0]):
                line_eleven = f"At the recommended threshold, the overall effectiveness will remain at {data_Rule_Pop_Parameter['Effectiveness'].values[0]:.2%}."
            elif data_Rule_Pop_Parameter['Current Threshold'].values[0] == data_Rule_Pop_Parameter['Recommended Threshold'].values[0] or data_Rule_Pop_Parameter['Effectiveness'].values[0] == data_Rule_Pop_Parameter['Prop Effectiveness'].values[0]:
                line_eleven = f"At the recommended threshold, the overall effectiveness will remain at {data_Rule_Pop_Parameter['Effectiveness'].values[0]:.2%}"
            elif data_Rule_Pop_Parameter['Effectiveness'].values[0] < data_Rule_Pop_Parameter['Prop Effectiveness'].values[0] and data_Rule_Pop_Parameter['SAR Yield'].values[0] > data_Rule_Pop_Parameter['Prop SAR Yield'].values[0]:
                line_eleven = f"At the recommended threshold, the overall effectiveness will increase from {data_Rule_Pop_Parameter['Effectiveness'].values[0]:.2%} to {data_Rule_Pop_Parameter['Prop Effectiveness'].values[0]:.2%}."
            elif data_Rule_Pop_Parameter['Effectiveness'].values[0] < data_Rule_Pop_Parameter['Prop Effectiveness'].values[0]:
                line_eleven = f"At the recommended threshold, the overall effectiveness will increase from {data_Rule_Pop_Parameter['Effectiveness'].values[0]:.2%} to {data_Rule_Pop_Parameter['Prop Effectiveness'].values[0]:.2%}"

            # Change in SAR yield
            line_twelve = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0 or data_Rule_Pop_Parameter['SAR Yield'].values[0] > data_Rule_Pop_Parameter['Prop SAR Yield'].values[0]:
                pass
            elif line_eleven == '' and (data_Rule_Pop_Parameter['Current Threshold'].values[0] == data_Rule_Pop_Parameter['Recommended Threshold'].values[0] or data_Rule_Pop_Parameter['SAR Yield'].values[0] == data_Rule_Pop_Parameter['Prop SAR Yield'].values[0]):
                line_twelve = f"At the recommended threshold, the overall SAR yield will remain at {data_Rule_Pop_Parameter['SAR Yield'].values[0]:.2%}."
            elif data_Rule_Pop_Parameter['Current Threshold'].values[0] == data_Rule_Pop_Parameter['Recommended Threshold'].values[0] or data_Rule_Pop_Parameter['SAR Yield'].values[0] == data_Rule_Pop_Parameter['Prop SAR Yield'].values[0]:
                line_twelve = f"and the overall SAR yield will remain at {data_Rule_Pop_Parameter['SAR Yield'].values[0]:.2%}."
            elif data_Rule_Pop_Parameter['SAR Yield'].values[0] < data_Rule_Pop_Parameter['Prop SAR Yield'].values[0] and line_eleven == '':
                line_twelve = f"At the recommended threshold, the overall SAR yield will increase from {data_Rule_Pop_Parameter['SAR Yield'].values[0]:.2%} to {data_Rule_Pop_Parameter['Prop SAR Yield'].values[0]:.2%}."
            elif data_Rule_Pop_Parameter['SAR Yield'].values[0] < data_Rule_Pop_Parameter['Prop SAR Yield'].values[0]:
                line_twelve = f"and the overall SAR yield will increase from {data_Rule_Pop_Parameter['SAR Yield'].values[0]:.2%} to {data_Rule_Pop_Parameter['Prop SAR Yield'].values[0]:.2%}."

            # Not interesting alert reduction
            line_thirteen = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            elif data_Rule_Pop_Parameter['Not Interesting Alert Reduction'].values[0] == 0:
                pass
            elif data_Rule_Pop_Parameter['Effectiveness'].values[0] > data_Rule_Pop_Parameter['Prop Effectiveness'].values[0] and data_Rule_Pop_Parameter['Not Interesting Alert Reduction'].values[0] > 0:
                line_thirteen = f"The recommended threshold will reduce the number of not interesting rule breaks by approximately {data_Rule_Pop_Parameter['Not Interesting Alert Reduction'].values[0]:.2%}."
            elif data_Rule_Pop_Parameter['Not Interesting Alert Reduction'].values[0] > 0:
                line_thirteen = f"Additionally, the recommended threshold will reduce the number of not interesting rule breaks by approximately {data_Rule_Pop_Parameter['Not Interesting Alert Reduction'].values[0]:.2%}."

            # Conclusion Paragraph ----------------------------------------------------
            ### Generates the conclusion

            thresholds_changed = []
            thresholds_changed_values = []
            thresholds_kept = []
            thresholds_kept_values = []

            # Selects parameter and recommended threshold for each row
            for row in data_Rule_Pop.iterrows():
                temp_param = row[1]['Parameter']
                temp_val = row[1]['Recommended Threshold']
                temp_istunable = row[1]['Is Tunable']

                # Formats values as needed
                if temp_param in currencyF:
                    temp_val_formatted = f"${temp_val:,}"
                elif temp_param in numberF:
                    if temp_val < 10:
                        temp_val_formatted = numbers_df[numbers_df['numbers'] == str(temp_val)]['alpha_numbers'].values[0]
                    else:
                        temp_val_formatted = f"{temp_val:,}"
                elif temp_param in percentF:
                    temp_val_formatted = f"{temp_val:.2%}"
                elif temp_param in decimalF:
                    temp_val_formatted = f"{temp_val:,.2f}"
                else:
                    temp_val_formatted = temp_val

                # Determines whether the parameter was changed or maintained
                if row[1]['Current Threshold'] != row[1]['Recommended Threshold']:
                    thresholds_changed.append(temp_param)
                    thresholds_changed_values.append(temp_val_formatted)
                else:
                    thresholds_kept.append(temp_param)
                    thresholds_kept_values.append(temp_val_formatted)

            # Generate a list of parameters changed with proper formatting
            if len(thresholds_changed) == 1:
                thresholds_changed_formatted = f"{thresholds_changed[0]} parameter"
            elif len(thresholds_changed) == 2:
                thresholds_changed_formatted = f"{thresholds_changed[0]} and {thresholds_changed[1]} parameters"
            else:
                thresholds_changed_formatted = ', '.join(thresholds_changed[:-1]) + ', and ' + thresholds_changed[-1] + ' parameters'

            # Generates a list of thresholds changed with proper formatting
            if len(thresholds_changed_values) == 1:
                thresholds_changed_values_formatted = thresholds_changed_values[0]
            elif len(thresholds_changed_values) == 2:
                thresholds_changed_values_formatted = f"{thresholds_changed_values[0]} and {thresholds_changed_values[1]} respectively"
            else:
                thresholds_changed_values_formatted = ', '.join(thresholds_changed_values[:-1]) + ', and ' + thresholds_changed_values[-1] + ' respectively'

            # Generate a list of parameters kept with proper formatting
            if len(thresholds_kept) == 1:
                thresholds_kept_formatted = f"{thresholds_kept[0]} parameter"
            elif len(thresholds_kept) == 2:
                thresholds_kept_formatted = f"{thresholds_kept[0]} and {thresholds_kept[1]} parameters"
            else:
                thresholds_kept_formatted = ', '.join(thresholds_kept[:-1]) + ', and ' + thresholds_kept[-1] + ' parameters'

            # Generates a list of thresholds kept with proper formatting
            if len(thresholds_kept_values) == 1:
                thresholds_kept_values_formatted = thresholds_kept_values[0]
            elif len(thresholds_kept_values) == 2:
                thresholds_kept_values_formatted = f"{thresholds_kept_values[0]} and {thresholds_kept_values[1]} respectively"
            else:
                thresholds_kept_values_formatted = ', '.join(thresholds_kept_values[:-1]) + ', and ' + thresholds_kept_values[-1] + ' respectively'

            # Populates the recommendations made
            line_fourteen = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            elif len(thresholds_changed) == 0:
                line_fourteen = f"Maintaining the {thresholds_kept_formatted} at {thresholds_kept_values_formatted}"
            elif len(thresholds_changed) > 0 and len(thresholds_kept) == 0:
                line_fourteen = f"Adjusting the {thresholds_changed_formatted} to {thresholds_changed_values_formatted}"
            elif len(thresholds_changed) > 0 and len(thresholds_kept) > 0:
                line_fourteen = f"Adjusting the {thresholds_changed_formatted} to {thresholds_changed_values_formatted} while maintaining the {thresholds_kept_formatted} at {thresholds_kept_values_formatted}"

            # Impact on Effectiveness
            temp_eff = data_Rule_Pop['Effectiveness'].values[0]
            temp_eff_formatted = f"{temp_eff:.2%}"

            temp_prop_eff = data_Rule_Pop['Net Effectiveness'].values[0]
            temp_prop_eff_formatted = f"{temp_prop_eff:.2%}"

            temp_sar = data_Rule_Pop['SAR Yield'].values[0]
            temp_sar_formatted = f"{temp_sar:.2%}"

            temp_prop_sar = data_Rule_Pop['Prop SAR Yield'].values[0]
            temp_prop_sar_formatted = f"{temp_prop_sar:.2%}"

            temp_not_int = data_Rule_Pop['Net Not Interesting Alert Reduction'].values[0]
            temp_not_int_formatted = f"{temp_not_int:.2%}"

            line_fifteen = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            elif len(thresholds_changed) == 0:
                line_fifteen = f"is expected to maintain the effectiveness and SAR yield at {temp_eff_formatted} and {temp_sar_formatted} respectively."
            elif len(thresholds_changed) >= 1 and temp_eff < temp_prop_eff and temp_sar < temp_prop_sar and temp_not_int > 0:
                line_fifteen = f"is expected to increase the overall effectiveness from {temp_eff_formatted} to {temp_prop_eff_formatted} and increase the overall SAR yield from {temp_sar_formatted} to {temp_prop_sar_formatted} while reducing the number of not interesting rule breaks by {temp_not_int_formatted}."
            elif len(thresholds_changed) >= 1 and temp_eff < temp_prop_eff and temp_sar < temp_prop_sar and temp_not_int == 0:
                line_fifteen = f"is expected to increase the overall effectiveness from {temp_eff_formatted} to {temp_prop_eff_formatted} and increase the overall SAR yield from {temp_sar_formatted} to {temp_prop_sar_formatted}."
            elif len(thresholds_changed) >= 1 and temp_eff < temp_prop_eff and temp_sar >= temp_prop_sar and temp_not_int > 0:
                line_fifteen = f"is expected to increase the overall effectiveness from {temp_eff_formatted} to {temp_prop_eff_formatted} while reducing the number of not interesting rule breaks by {temp_not_int_formatted}."
            elif len(thresholds_changed) >= 1 and temp_eff < temp_prop_eff and temp_sar >= temp_prop_sar and temp_not_int == 0:
                line_fifteen = f"is expected to increase the overall effectiveness from {temp_eff_formatted} to {temp_prop_eff_formatted}."
            elif len(thresholds_changed) >= 1 and temp_eff >= temp_prop_eff and temp_sar < temp_prop_sar and temp_not_int > 0:
                line_fifteen = f"is expected to increase the overall SAR yield from {temp_sar_formatted} to {temp_prop_sar_formatted} while reducing the number of not interesting rule breaks by {temp_not_int_formatted}."
            elif len(thresholds_changed) >= 1 and temp_eff >= temp_prop_eff and temp_sar < temp_prop_sar and temp_not_int == 0:
                line_fifteen = f"is expected to increase the overall SAR yield from {temp_sar_formatted} to {temp_prop_sar_formatted}."
            elif len(thresholds_changed) >= 1 and temp_eff >= temp_prop_eff and temp_sar >= temp_prop_sar and temp_not_int > 0:
                line_fifteen = f"is expected to reduce the number of not interesting rule breaks by {temp_not_int_formatted}."
            elif len(thresholds_changed) >= 1 and temp_eff >= temp_prop_eff and temp_sar >= temp_prop_sar and temp_not_int == 0:
                line_fifteen = "###THIS RECOMMENDATION DOES NOT CHANGE THE EFFECTIVENESS, SAR YIELD, OR REDUCE FALSE POSITIVES. PLEASE REVIEW###"

            # Expected rule breaks per month
            line_sixteen = ''
            if data_Rule_Pop_Parameter['Num Alerts Extracted'].values[0] == 0:
                pass
            else:
                line_sixteen = "The segment is expected to generate approximately ## rule breaks per month."

            # Consolidate Narratives --------------------------------------------------
            # Populates narratives table made earlier for export
            narratives = narratives.append({'Rule ID': data_Rule_Pop_Parameter['Rule ID'].values[0],
                                            'Rule Name': data_Rule_Pop_Parameter['Rule Name'].values[0],
                                            'Population Group': data_Rule_Pop_Parameter['Population Group'].values[0],
                                            'Parameter': data_Rule_Pop_Parameter['Parameter'].values[0],
                                            'Summary': line_one + ' ' + line_two + ' ' + line_three + ' ' + line_four + ' ' + line_five,
                                            'Analysis': line_six + ' ' + line_seven + ' ' + line_eight + ' ' + line_nine + ' ' + line_ten + ' ' + line_eleven + ' ' + line_twelve + ' ' + line_thirteen,
                                            'Conclusion': line_fourteen + ' ' + line_fifteen + ' ' + line_sixteen}, ignore_index=True)

# Export Narratives to Word -----------------------------------------------
# Initialize output file
output = os.path.join(fp, report_name)
doc = Document()

# Create section named 'Analysis'
section = doc.add_section(WD_SECT.NEW_PAGE)
section.title = 'Analysis'

# Filter narratives to rule level
for x in ruleIDs:
    narratives_Rule = narratives[narratives['Rule ID'] == x]

    # Add Rule Name/ID header to doc
    section.add_heading(f"{narratives_Rule['Rule Name'].values[0]} | {x}", level=1)

    # Add threshold decisions table
    section.add_heading('Summary of Threshold Decisions', level=2)

    threshold_decisions = data[data['Rule ID'] == x][['Population Group', 'Parameter', 'Current Threshold', 'Recommended Threshold']]
    threshold_decisions.columns = ['Population', 'Parameter', 'Current Threshold', 'Recommended Threshold']
    threshold_decisions = threshold_decisions.sort_values(['Population', 'Parameter'])

    # Format thresholds in table to display with proper units
    for row in threshold_decisions.iterrows():
        temp_param = row[1]['Parameter']
        temp_values = [row[1]['Current Threshold'], row[1]['Recommended Threshold']]

        if temp_param in currencyF:
            temp_values_formatted = [f"${temp:,.2f}" for temp in temp_values]
        elif temp_param in numberF:
            temp_values_formatted = [f"{temp:,}" for temp in temp_values]
        elif temp_param in percentF:
            temp_values_formatted = [f"{temp:.2%}" for temp in temp_values]
        elif temp_param in decimalF:
            temp_values_formatted = [f"{temp:,.2f}" for temp in temp_values]
        else:
            temp_values_formatted = temp_values

        row[1]['Current Threshold'] = temp_values_formatted[0]
        row[1]['Recommended Threshold'] = temp_values_formatted[1]

    table = section.add_table(threshold_decisions.shape[0]+1, threshold_decisions.shape[1])
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Population'
    hdr_cells[1].text = 'Parameter'
    hdr_cells[2].text = 'Current Threshold'
    hdr_cells[3].text = 'Recommended Threshold'

    for idx, row in threshold_decisions.iterrows():
        cells = table.rows[idx+1].cells
        cells[0].text = row['Population']
        cells[1].text = row['Parameter']
        cells[2].text = row['Current Threshold']
        cells[3].text = row['Recommended Threshold']

    section.add_paragraph()

    # Filter to Pop Group
    popGroups = narratives_Rule['Population Group'].unique()

    for pop in popGroups:
        narratives_Rule_Pop = narratives_Rule[narratives_Rule['Population Group'] == pop]

        # Add header for population group
        section.add_heading(pop, level=2)

        # Add 'Threshold Recommendation' header to doc
        section.add_heading('Threshold Recommendation', level=3)

        # Print summary paragraph
        section.add_paragraph(narratives_Rule_Pop['Summary'].values[0])

        # Skip remaining headers/paragraphs if no alerts generated
        if narratives_Rule_Pop['Summary'].values[0].startswith('No alerts'):
            continue

        # Filter to each parameter within pop group
        params = narratives_Rule_Pop['Parameter'].unique()

        for param in params:
            narratives_Rule_Pop_Parameter = narratives_Rule_Pop[narratives_Rule_Pop['Parameter'] == param]

            # Add parameter header
            section.add_heading(param, level=4)

            # Print analysis paragraphs for each parameter
            section.add_paragraph(narratives_Rule_Pop_Parameter['Analysis'].values[0])

        # Add header for conclusion
        section.add_heading('Conclusion', level=3)


        # changelogv1: margin in document
        #Open the document
        #changing the page margins
        sections = document.sections
        margin = 0.5
        for section in sections:
            section.top_margin = Cm(margin)
            section.bottom_margin = Cm(margin)
            section.left_margin = Cm(margin)
            section.right_margin = Cm(margin)

        # Print conclusion paragraph
        section.add_paragraph(narratives_Rule_Pop['Conclusion'].values[0])

        # Add placeholder for scoring
        section.add_heading('Scoring Recommendation', level=2)

doc.save(output)