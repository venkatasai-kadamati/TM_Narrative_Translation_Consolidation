# This is a script to generate report narratives for Actimize production tuning engagements
# The following columns need to be present in the tuning tracker to use this script: 
#   Rule Name, Rule ID, Population Group, Parameter, Current Threshold, Recommended Threshold, Parameter Type, and Operator

# The script assumes that the total population of alerts = SARs Filed + Interesting Alerts + Not Interesting Alerts
# For analysis, data quality alerts are excluded from the total

# Effectiveness is (Interesting + SAR)/(Interesting + SAR + Not Interesting)
# References to interesting alerts in the narratives generated refer to Interesting + SAR Filed
# The script treats all parameters included in the tracker as tunable 
# Ratios are assumed to be multiplied by 100 (a ratio of 80% would be entered at 80)

# FOR RULES THAT USE A RATIO RANGE, THE ANALYSIS WILL NEED TO BE EDITED

# FOR ALL CONCLUSIONS, THE EXPECTED MONTHLY RULE BREAK COUNT WILL NEED TO BE PROVIDED


# Data Load ---------------------------------------------------------------
rm(list = ls()) # Clears environment before running

library('dplyr') # used for data manipulation
library('readxl') # used to load data from Excel
library('scales') # used to convert numbers to %
library('rtf') # used to create a rich text format (RTF) doc that can be opened in Word

options(scipen = 999)

# Specify location of the tuning tracker
fp = "C:/Users/KadamatiV/Desktop/WLF TUNING CODE CONSOLIDATION/NARRATIVE SCRIPT MERGE/"
file_name = 'Tuning Tracker - ATL Calculationsv11.xlsx'

report_name <- 'CNB - Production Report Narratives.docx' # Enter the desired name for the report that will be generated (must end in '.doc')

# Specify how to format the different parameters seen (change threshold names as needed)
currencyF <- c('Minimal Sum','Minimum Value','Minimal Transaction Amount','Sum Lower Bound', 'Min Value', 'Minimal Current Month Sum', 'Minimal Transaction Value', 'Transaction Amount Lower Bound') # Values should be formatted as money
numberF <- c('No. of Occurrences','Minimum Volume', 'Min Value') # Values should be formatted as integers 
percentF <- c('Ratio Lower Bound','Ratio Upper Bound') # Values should be formatted as percentage
decimalF <- c('STDEV exceeds Historical Average Sum', 'STDEV exceeds Historical Average Count') # Values should be formatted to 2 decimals

# Data is loaded
data <- read_excel(paste(fp, file_name, sep=''), sheet = 1)

# Enter the column name of the Rule IDs, if different
ruleIDs <- table(data$`Rule ID`)

# Create lookup for values below 10 that need to be written out (no change)
numbers <- c('0','1','2','3','4','5','6','7','8','9')
alpha_numbers <- c('zero (0)', 'one (1)', 'two (2)', 'three (3)', 'four (4)', 'five (5)', 'six (6)', 'seven (7)', 'eight (8)', 'nine (9)')
alpha_numbers_cap <- c('Zero (0)', 'One (1)', 'Two (2)', 'Three (3)', 'Four (4)', 'Five (5)', 'Six (6)', 'Seven (7)', 'Eight (8)', 'Nine (9)')
numbers.df <- data.frame(numbers, alpha_numbers, alpha_numbers_cap)

# Create an empty data frame to hold narratives
narratives <- data.frame(matrix(vector(),0,7))

# Begin Iterations --------------------------------------------------------

# Iterate over Rule IDs, Population Groups, and Parameters
for (x in names(ruleIDs)) { 
  
  data.Rule <- data %>% filter(`Rule ID` == x)
  popGroups <- table(data.Rule$`Population Group`)
  
  for (pop in names(popGroups)) {
    
    data.Rule.Pop <- data.Rule %>% filter(data.Rule$`Population Group` == pop)
    params <- table(data.Rule.Pop$Parameter)
    
    for (threshold in names(params)) {
      
      data.Rule.Pop.Parameter <- data.Rule.Pop %>% filter(data.Rule.Pop$Parameter == threshold)
      
      # Summary Paragraph -------------------------------------------------------
      ### Create paragraph to be populated before parameter analysis
      
      # Time frame of alerts that generated; if no alerts generated -> no analysis performed
      temp <- data.Rule.Pop.Parameter$`Num Alerts Extracted`
      date_range = unlist(strsplit(data.Rule.Pop.Parameter$`Date Range`,'-'))
      
      temp2 <- names(table(data.Rule.Pop$Parameter))
      if (length(temp2) == 1){params.used <- paste(temp2, ' parameter', sep = '')
      } else if (length(temp2) == 2){params.used <- paste(temp2[1],' and ',temp2[2], ' parameters',sep='')
      } else {params.used <- paste(paste(head(temp2,-1),sep = ', ', collapse = ', '), paste(', and ', tail(temp2,1), ' parameters', sep = ''), sep = '')}
      
      line_one <- case_when (
        temp == 0 ~ paste('No alerts generated in the Actimize environment for the ',data.Rule.Pop.Parameter$`Population Group`,' population group between ',date_range[1],' and ',date_range[2],'; therefore, no analysis was performed. The current thresholds are recommended to be maintained.',sep=''),
        temp > 0 ~ paste('Production tuning analysis was performed on the ', params.used, ' for the ', data.Rule.Pop$`Population Group`, ' population group.', sep = '')
      )
      
      # No. of rule breaks generated
      temp <- data.Rule.Pop.Parameter$`Num Alerts Extracted`
      
      line_two <- case_when (
        temp == 0 ~ '',
        temp == 1 ~ paste(numbers.df[temp+1,3], ' rule break was generated between ', date_range[1],' and ',date_range[2],'.', sep=''),
        temp < 10 ~ paste(numbers.df[temp+1,3], ' rule breaks were generated between ', date_range[1],' and ',date_range[2],'.', sep=''),
        temp >= 10 ~ paste(format(temp, big.mark = ','), ' rule breaks were generated between ', date_range[1],' and ',date_range[2],'.', sep=''))
      
      # Data quality alerts identified, if any
      temp <- data.Rule.Pop.Parameter$`Data Quality Alerts`
      
      line_three <- case_when (
        temp == 0 ~ '',
        temp == 1 ~ paste('The Bank identified ',numbers.df[temp+1,2],' Data Quality rule break that generated on duplicated or incorrectly mapped transactional activity. As a result, this rule break was marked as Data Quality and excluded from analysis.', sep=''),
        temp < 10 ~ paste('The Bank identified ',numbers.df[temp+1,2],' Data Quality rule breaks that generated on duplicated or incorrectly mapped transactional activity. As a result, these rule breaks were marked as Data Quality and excluded from analysis.', sep=''),
        temp >= 10 ~ paste('The Bank identified ',temp,' Data Quality rule breaks that generated on duplicated or incorrectly mapped transactional activity. As a result, these rule breaks were marked as Data Quality and excluded from analysis.', sep=''))
      
      # Interesting alerts / effectiveness
      temp.total <- data.Rule.Pop.Parameter$`Num Alerts Extracted` - data.Rule.Pop.Parameter$`Data Quality Alerts`
      temp.sar <- data.Rule.Pop.Parameter$`SARs Filed`
      temp.int <- data.Rule.Pop.Parameter$`Interesting Alerts` + data.Rule.Pop.Parameter$`SARs Filed`
      temp.eff <- format(round(data.Rule.Pop.Parameter$Effectiveness, 2), nsmall = 2)
      temp.sar_yield <- format(round(data.Rule.Pop.Parameter$`SAR Yield`, 2), nsmall = 2)
      
      line_four <- case_when (
        temp.total == 0 ~ '',
        temp.total == 1 & temp.int == 1 & temp.sar == 1 ~ paste('The one (1) reviwed rule break was determined to be Interesting and led to a SAR filing, resulting in a production effectiveness and SAR yield of 100.00%.', sep=''),
        temp.total == 1 & temp.int == 1 & temp.sar == 0 ~ paste('The one (1) reviwed rule break was determined to be Interesting but did not end in a SAR filing, resulting in a production effectiveness of 100.00% and a SAR yield of 0.00%.', sep=''),
        temp.total == 1 & temp.int != 1 ~ paste('The one (1) reviwed rule break was not determined to be Interesting, resulting in a production effectiveness of 0.00%.', sep=''),
        temp.total < 10 & temp.int == 1 ~ paste('Of the total ', numbers.df[temp.total+1,2], ' reviewed rule breaks, ', numbers.df[temp.int+1,2], 'rule break was determined to be Interesting, resulting in a production effectiveness of ', temp.eff, '%.', sep=''),
        temp.total < 10 & temp.int < 10 ~ paste('Of the total ', numbers.df[temp.total+1,2], ' reviewed rule breaks, ', numbers.df[temp.int+1,2], 'rule breaks were determined to be Interesting, resulting in a production effectiveness of ', temp.eff, '%.', sep=''),
        temp.total >= 10 & temp.int == 1 ~ paste('Of the total ', format(temp.total,big.mark = ','), ' reviewed rule breaks, ', numbers.df[temp.int+1,2], ' rule break was determined to be Interesting, resulting in a production effectiveness of ', temp.eff, '%.', sep=''),
        temp.total >= 10 & temp.int < 10 ~ paste('Of the total ', format(temp.total,big.mark = ','), ' reviewed rule breaks, ', numbers.df[temp.int+1,2], ' rule breaks were determined to be Interesting, resulting in a production effectiveness of ', temp.eff, '%.', sep=''),
        temp.total >= 10 & temp.int >= 10 ~ paste('Of the total ', format(temp.total,big.mark = ','), ' reviewed rule breaks, ', format(temp.int,big.mark = ','), ' rule breaks were determined to be Interesting, resulting in a production effectiveness of ', temp.eff, '%.', sep=''))
      
      # SARs filed / SAR yield (using temp variables from line_four)
      line_five <- case_when (
        temp.total == 0 ~ '',
        temp.int == 0 ~ '',
        temp.sar < 10 & temp.int < 10 ~ paste(numbers.df[temp.sar+1,3], ' of the ', numbers.df[temp.int+1,2], ' Interesting rule breaks led to a SAR filing, resulting in a SAR yield of ', temp.sar_yield, '%. Additional detail on the analysis results per population group is provided below.', sep=''),
        temp.sar < 10 & temp.int >= 10 ~ paste(numbers.df[temp.sar+1,3], ' of the ', temp.int, ' Interesting rule breaks led to a SAR filing, resulting in a SAR yield of ', temp.sar_yield, '%. Additional detail on the analysis results per population group is provided below.', sep=''),
        temp.sar >= 10 ~ paste(format(temp.sar,big.mark = ','), ' of the ', temp.int, 'Interesting rule breaks led to a SAR filing, resulting in a SAR yield of ', temp.sar_yield, '%. Additional detail on the analysis results per population group is provided below.', sep='')
      )
      
      # Analysis Paragraph ------------------------------------------------------
      ### Create paragraph to be populated under parameter analysis

      # Threshold tuned and value
      temp <- data.Rule.Pop.Parameter$`Current Threshold`
      if (threshold %in% currencyF){temp.formatted <- paste('$', format(temp,big.mark = ',', nsmall = 2), sep = '')
      } else if (threshold %in% numberF){
        if (temp < 10) {temp.formatted <- numbers.df[temp+1,2]
        } else temp.formatted <- format(temp,big.mark = ',')
      } else if (threshold %in% percentF){temp.formatted <- percent(temp/100)
      } else if (threshold %in% decimalF){temp.formatted <- format(round(temp, digits = 2), big.mark = ',', nsmall = 2)
      } else temp.formatted <- temp
      
      line_six <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        data.Rule.Pop.Parameter$`Num Alerts Extracted` > 0 ~ paste('Above-the-line tuning was conducted on the ', data.Rule.Pop.Parameter$Parameter, ' threshold, which was set at a value of ', temp.formatted, '.', sep=''))

      # Values that rule breaks generated at
      temp <- data.Rule.Pop.Parameter$`Max Val` - data.Rule.Pop.Parameter$`Min Val`

      if (threshold %in% currencyF){val.formatted <- c(paste('$', format(data.Rule.Pop.Parameter$`Min Val`, big.mark = ',', nsmall = 2),sep = ''), paste('$', format(data.Rule.Pop.Parameter$`Max Val`, big.mark = ',', nsmall = 2),sep = ''))
      } else if (threshold %in% numberF){
        if (is.na(data.Rule.Pop.Parameter$`Max Val`)) {val.formatted <- c('0','0')
        } else if (data.Rule.Pop.Parameter$`Max Val` < 10) {val.formatted <- c(numbers.df[data.Rule.Pop.Parameter$`Min Val`+1,2],numbers.df[data.Rule.Pop.Parameter$`Max Val`+1,2])
        } else if (data.Rule.Pop.Parameter$`Max Val` >= 10 & data.Rule.Pop.Parameter$`Min Val` < 10) {val.formatted <- c(numbers.df[data.Rule.Pop.Parameter$`Min Val`+1,2],format(data.Rule.Pop.Parameter$`Max Val`,big.mark = ','))
        } else val.formatted <- c(format(data.Rule.Pop.Parameter$`Min Val`,big.mark = ','),format(data.Rule.Pop.Parameter$`Max Val`,big.mark = ','))
      } else if (threshold %in% percentF){val.formatted <- c(percent(data.Rule.Pop.Parameter$`Min Val`/100,big.mark = ','),percent(data.Rule.Pop.Parameter$`Max Val`/100,big.mark = ','))
      } else if (threshold %in% decimalF){val.formatted <- c(format(round(data.Rule.Pop.Parameter$`Min Val`, digits = 2), big.mark = ',', nsmall = 2),format(round(data.Rule.Pop.Parameter$`Max Val`, digits = 2), big.mark = ',', nsmall = 2))
      } else val.formatted <- c(data.Rule.Pop.Parameter$`Min Val`, data.Rule.Pop.Parameter$`Max Val`)
      
      line_seven <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        temp == 0 ~ paste('Rule breaks were generated solely at a value of ', val.formatted[2], ' within the ', data.Rule.Pop.Parameter$`Population Group`, ' population segment.', sep=''),
        temp != 0 ~ paste('Rule breaks were generated for values ranging between ', val.formatted[1], ' and ', val.formatted[2], ' within the ', data.Rule.Pop.Parameter$`Population Group`, ' population segment.', sep=''))
      
      # No. of interesting rule breaks in rule population
      temp.sar <- data.Rule.Pop.Parameter$`SARs Filed`
      temp.int <- data.Rule.Pop.Parameter$`Interesting Alerts` + data.Rule.Pop.Parameter$`SARs Filed`
      temp.eff <- data.Rule.Pop.Parameter$Effectiveness
      temp.sar_yield <- data.Rule.Pop.Parameter$`SAR Yield`
      
      line_eight <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        temp.int == 1 & temp.sar == 1 ~ paste('One (1) Interesting rule break was noted in the production population which also resulted in a SAR filing.', sep=''),
        temp.int == 1 & temp.sar != 1 ~ paste('One (1) Interesting rule break was noted in the production population which did not result in a SAR filing.', sep=''),
        temp.int < 10 & temp.sar == 1 ~ paste(numbers.df[temp.int+1,3], ' Interesting rule breaks were noted in the production population, of which one (1) rule break resulted in a SAR filing.', sep=''),
        temp.int < 10 ~ paste(numbers.df[temp.int+1,3], ' Interesting rule breaks were noted in the production population, of which ', numbers.df[temp.sar+1,2], ' rule breaks resulted in SAR filings.', sep=''),
        temp.int >= 10 & temp.sar == 1 ~ paste(format(temp.int, big.mark = ','), ' Interesting rule breaks were noted in the production population, of which one (1) rule break resulted in a SAR filing.', sep=''),
        temp.int >= 10 & temp.sar < 10 ~ paste(format(temp.int, big.mark = ','), ' Interesting rule breaks were noted in the production population, of which ', numbers.df[temp.sar+1,2], ' rule breaks resulted in SAR filings.', sep=''),
        temp.int >= 10 & temp.sar >= 10 ~ paste(format(temp.int, big.mark = ','), ' Interesting rule breaks were noted in the production population, of which ', format(temp.sar, big.mark = ','), ' rule breaks resulted in SAR filings.', sep=''))
      
      # Placeholder for recommendation rationale 
      line_nine <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        data.Rule.Pop.Parameter$`Num Alerts Extracted` != 0 ~ '###INSERT TUNING DECISION###')
      
      # Tuning recommendation
      temp <- c(data.Rule.Pop.Parameter$`Current Threshold`, data.Rule.Pop.Parameter$`Recommended Threshold`)
      if (threshold %in% currencyF){temp.formatted <- c(paste('$', format(temp[1],big.mark = ',', nsmall = 2), sep = ''), paste('$', format(temp[2],big.mark = ',', nsmall = 2), sep = ''))
      } else if (threshold %in% numberF){
        if (temp[1] < 10 & temp[2] < 10) {temp.formatted <- c(numbers.df[temp[1]+1,2], numbers.df[temp[2]+1,2])
        } else if (temp[1] < 10 & temp[2] >= 10) {temp.formatted <- c(numbers.df[temp[1]+1,2],format(temp[2],big.mark = ','))
        } else temp.formatted <- c(format(temp[1],big.mark = ','), format(temp[2],big.mark = ','))
      } else if (threshold %in% percentF){temp.formatted <- c(percent(temp[1]/100), percent(temp[2]/100))
      } else if (threshold %in% decimalF){temp.formatted <- c(format(round(temp[1], digits=2), big.mark=',', nsmall=2),format(round(temp[2], digits=2), big.mark=',', nsmall=2))
      } else temp.formatted <- temp
      
      line_ten <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        data.Rule.Pop.Parameter$`Current Threshold` == data.Rule.Pop.Parameter$`Recommended Threshold` ~ paste('Therefore, it is recommended to maintain the ',data.Rule.Pop.Parameter$Parameter,' threshold at ',temp.formatted[1],'.',sep=''),
        data.Rule.Pop.Parameter$`Current Threshold` != data.Rule.Pop.Parameter$`Recommended Threshold` ~ paste('Therefore, it is recommended to increase the ',data.Rule.Pop.Parameter$Parameter,' threshold from ', temp.formatted[1], ' to ', temp.formatted[2], '.',sep=''))
      
      # Change in effectiveness
      line_eleven <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 | data.Rule.Pop.Parameter$`Effectiveness` > data.Rule.Pop.Parameter$`Prop Effectiveness` ~ '',
        data.Rule.Pop.Parameter$`SAR Yield` > data.Rule.Pop.Parameter$`Prop SAR Yield` & (data.Rule.Pop.Parameter$`Current Threshold` == data.Rule.Pop.Parameter$`Recommended Threshold` | data.Rule.Pop.Parameter$`Effectiveness` == data.Rule.Pop.Parameter$`Prop Effectiveness`) ~ paste('At the recommended threshold, the overall effectiveness will remain at ',format(data.Rule.Pop.Parameter$Effectiveness, nsmall = 2),'%.',sep=''),
        data.Rule.Pop.Parameter$`Current Threshold` == data.Rule.Pop.Parameter$`Recommended Threshold` | data.Rule.Pop.Parameter$`Effectiveness` == data.Rule.Pop.Parameter$`Prop Effectiveness` ~ paste('At the recommended threshold, the overall effectiveness will remain at ',format(data.Rule.Pop.Parameter$Effectiveness, nsmall = 2),'%',sep=''),
        data.Rule.Pop.Parameter$`Effectiveness` < data.Rule.Pop.Parameter$`Prop Effectiveness` & data.Rule.Pop.Parameter$`SAR Yield` > data.Rule.Pop.Parameter$`Prop SAR Yield` ~ paste('At the recommended threshold, the overall effectiveness will increase from ',format(data.Rule.Pop.Parameter$Effectiveness, nsmall = 2), '% to ', format(data.Rule.Pop.Parameter$`Prop Effectiveness`, nsmall = 2),'%.',sep=''),
        data.Rule.Pop.Parameter$`Effectiveness` < data.Rule.Pop.Parameter$`Prop Effectiveness` ~ paste('At the recommended threshold, the overall effectiveness will increase from ',format(data.Rule.Pop.Parameter$Effectiveness, nsmall = 2), '% to ', format(data.Rule.Pop.Parameter$`Prop Effectiveness`, nsmall = 2),'%',sep=''))
      
      # Change in SAR yield
      line_twelve <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 | data.Rule.Pop.Parameter$`SAR Yield` > data.Rule.Pop.Parameter$`Prop SAR Yield` ~ '',
        line_eleven == '' & (data.Rule.Pop.Parameter$`Current Threshold` == data.Rule.Pop.Parameter$`Recommended Threshold` | data.Rule.Pop.Parameter$`SAR Yield` == data.Rule.Pop.Parameter$`Prop SAR Yield`) ~ paste('At the recommended threshold, the overall SAR yield will remain at ',format(data.Rule.Pop.Parameter$`SAR Yield`, nsmall = 2),'%.',sep=''),
        data.Rule.Pop.Parameter$`Current Threshold` == data.Rule.Pop.Parameter$`Recommended Threshold` | data.Rule.Pop.Parameter$`SAR Yield` == data.Rule.Pop.Parameter$`Prop SAR Yield` ~ paste('and the overall SAR yield will remain at ',format(data.Rule.Pop.Parameter$`SAR Yield`, nsmall = 2),'%.',sep=''),
        data.Rule.Pop.Parameter$`SAR Yield` < data.Rule.Pop.Parameter$`Prop SAR Yield` & line_eleven == '' ~ paste('At the recommended threshold, the overall SAR yield will increase from ',format(data.Rule.Pop.Parameter$`SAR Yield`, nsmall = 2), '% to ', format(data.Rule.Pop.Parameter$`Prop SAR Yield`, nsmall = 2),'%.',sep=''),
        data.Rule.Pop.Parameter$`SAR Yield` < data.Rule.Pop.Parameter$`Prop SAR Yield` ~ paste('and the overall SAR yield will increase from ',format(data.Rule.Pop.Parameter$`SAR Yield`, nsmall = 2), '% to ', format(data.Rule.Pop.Parameter$`Prop SAR Yield`, nsmall = 2),'%.',sep=''))
      
      # Not interesting alert reduction
      line_thirteen <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        data.Rule.Pop.Parameter$`Not Interesting Alert Reduction` == 0 ~ '',
        data.Rule.Pop.Parameter$`Effectiveness` > data.Rule.Pop.Parameter$`Prop Effectiveness` & data.Rule.Pop.Parameter$`Not Interesting Alert Reduction` > 0 ~ paste('The recommended threshold will reduce the number of not interesting rule breaks by approximately ',format(data.Rule.Pop.Parameter$`Not Interesting Alert Reduction`, nsmall = 2),'%.',sep=''),
        data.Rule.Pop.Parameter$`Not Interesting Alert Reduction` > 0 ~ paste('Additionally, the recommended threshold will reduce the number of not interesting rule breaks by approximately ',format(data.Rule.Pop.Parameter$`Not Interesting Alert Reduction`, nsmall = 2),'%.',sep=''))
      
      # Conclusion Paragraph ----------------------------------------------------
      ### Generates the conclusion
      
      thresholds.changed <- character()
      thresholds.changed.values <- character()
      thresholds.kept <- character()
      thresholds.kept.values <- character()
      
      # Selects parameter and recommended threshold for each row
      for (row in 1:nrow(data.Rule.Pop)){
        temp.param <- data.Rule.Pop[row,'Parameter'][[1]]
        temp.val <- data.Rule.Pop[row,'Recommended Threshold'][[1]]
        temp.istunable <- data.Rule.Pop[row,'Is Tunable'][[1]]
        
        # Formats values as needed
        if (temp.param %in% currencyF){temp.val.formatted <- paste('$', format(temp.val,big.mark = ',', nsmall = 2), sep = '')
        } else if (temp.param %in% numberF){
          if(temp.val < 10){temp.val.formatted <- numbers.df[temp.val+1,2]
          } else temp.val.formatted <- format(temp.val,big.mark = ',')
        } else if (temp.param %in% percentF){temp.val.formatted <- percent(temp.val/100)
        } else if (temp.param %in% decimalF){temp.val.formatted <- format(round(temp.val, digits=2), big.mark=',', nsmall=2)
        } else temp.val.formatted <- temp.val
        
        # Determines whether the parameter was changed or maintained
        if (data.Rule.Pop[row,'Current Threshold'] != data.Rule.Pop[row,'Recommended Threshold']){
          thresholds.changed <- c(thresholds.changed, temp.param)
          thresholds.changed.values <- c(thresholds.changed.values, temp.val.formatted)
        } else {
          thresholds.kept <- c(thresholds.kept, temp.param)
          thresholds.kept.values <- c(thresholds.kept.values, temp.val.formatted)}}
      
      # Generate a list of parameters changed with proper formatting
      if (length(thresholds.changed) == 1){thresholds.changed.formatted <- paste(thresholds.changed, ' parameter', sep = '')
      } else if (length(thresholds.changed) == 2){thresholds.changed.formatted <- paste(thresholds.changed[1],' and ',thresholds.changed[2], ' parameters',sep='')
      } else {thresholds.changed.formatted <- paste(paste(head(thresholds.changed,-1),sep = ', ', collapse = ', '), paste(', and ', tail(thresholds.changed,1), ' parameters', sep = ''), sep = '')}
      
      # Generates a list of thresholds changed with proper formatting
      if (length(thresholds.changed.values) == 1){thresholds.changed.values.formatted <- thresholds.changed.values
      } else if (length(thresholds.changed.values) == 2){thresholds.changed.values.formatted <- paste(thresholds.changed.values[1],' and ',thresholds.changed.values[2], ' respectively',sep='')
      } else {thresholds.changed.values.formatted <- paste(paste(head(thresholds.changed.values,-1),sep = ', ', collapse = ', '), paste(', and ', tail(thresholds.changed.values,1), ' respectively', sep = ''), sep = '')}
      
      # Generate a list of parameters kept with proper formatting
      if (length(thresholds.kept) == 1){thresholds.kept.formatted <- paste(thresholds.kept, ' parameter', sep = '')
      } else if (length(thresholds.kept) == 2){thresholds.kept.formatted <- paste(thresholds.kept[1],' and ',thresholds.kept[2], ' parameters',sep='')
      } else {thresholds.kept.formatted <- paste(paste(head(thresholds.kept,-1),sep = ', ', collapse = ', '), paste(', and ', tail(thresholds.kept,1), ' parameters', sep = ''), sep = '')}
      
      # Generates a list of thresholds kept with proper formatting
      if (length(thresholds.kept.values) == 1){thresholds.kept.values.formatted <- thresholds.kept.values
      } else if (length(thresholds.kept.values) == 2){thresholds.kept.values.formatted <- paste(thresholds.kept.values[1],' and ',thresholds.kept.values[2], ' respectively',sep='')
      } else {thresholds.kept.values.formatted <- paste(paste(head(thresholds.kept.values,-1),sep = ', ', collapse = ', '), paste(', and ', tail(thresholds.kept.values,1), ' respectively', sep = ''), sep = '')}
      
      # Populates the recommendations made
      line_fourteen <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        length(thresholds.changed) == 0 ~ paste('Maintaining the ', thresholds.kept.formatted, ' at ', thresholds.kept.values.formatted, sep = ''),
        length(thresholds.changed) > 0 & length(thresholds.kept) == 0 ~ paste('Adjusting the ', thresholds.changed.formatted, ' to ', thresholds.changed.values.formatted, sep = ''),
        length(thresholds.changed) > 0 & length(thresholds.kept) > 0 ~ paste('Adjusting the ', thresholds.changed.formatted, ' to ', thresholds.changed.values.formatted, ' while maintaining the ', thresholds.kept.formatted, ' at ', thresholds.kept.values.formatted, sep = ''))
      
      # Impact on Effectiveness
      temp.eff <- data.Rule.Pop$Effectiveness[1]
      temp.eff.formatted <- format(temp.eff, nsmall = 2)
      
      temp.prop.eff <- data.Rule.Pop$`Net Effectiveness`[1]
      temp.prop.eff.formatted <- format(temp.prop.eff, nsmall = 2)
      
      temp.sar <- data.Rule.Pop$`SAR Yield`[1]
      temp.sar.formatted <- format(temp.sar, nsmall = 2)
      
      temp.prop.sar <- data.Rule.Pop$`Prop SAR Yield`[1]
      temp.prop.sar.formatted <- format(temp.prop.sar, nsmall = 2)
      
      temp.not.int <- data.Rule.Pop$`Net Not Interesting Alert Reduction`[1]
      temp.not.int.formatted <- format(temp.not.int, nsmall = 2)
      
      line_fifteen <- case_when (
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        length(thresholds.changed) == 0 ~ paste('is expected to maintain the effectiveness and SAR yield at ', temp.eff.formatted, '% and ', temp.sar.formatted, '% respectively.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff < temp.prop.eff & temp.sar < temp.prop.sar & temp.not.int > 0 ~ paste('is expected to increase the overall effectiveness from ', temp.eff.formatted, '% to ', temp.prop.eff.formatted, '% and increase the overall SAR yield from ', temp.sar.formatted, '% to ', temp.prop.sar.formatted, '% while reducing the number of not interesting rule breaks by ', temp.not.int.formatted, '%.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff < temp.prop.eff & temp.sar < temp.prop.sar & temp.not.int == 0 ~ paste('is expected to increase the overall effectiveness from ', temp.eff.formatted, '% to ', temp.prop.eff.formatted, '% and increase the overall SAR yield from ', temp.sar.formatted, '% to ', temp.prop.sar.formatted, '%.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff < temp.prop.eff & temp.sar >= temp.prop.sar & temp.not.int > 0 ~ paste('is expected to increase the overall effectiveness from ', temp.eff.formatted, '% to ', temp.prop.eff.formatted, '% while reducing the number of not interesting rule breaks by ', temp.not.int.formatted, '%.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff < temp.prop.eff & temp.sar >= temp.prop.sar & temp.not.int == 0 ~ paste('is expected to increase the overall effectiveness from ', temp.eff.formatted, '% to ', temp.prop.eff.formatted, '%.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff >= temp.prop.eff & temp.sar < temp.prop.sar & temp.not.int > 0 ~ paste('is expected to increase the overall SAR yield from ', temp.sar.formatted, '% to ', temp.prop.sar.formatted, '% while reducing the number of not interesting rule breaks by ', temp.not.int.formatted, '%.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff >= temp.prop.eff & temp.sar < temp.prop.sar & temp.not.int == 0 ~ paste('is expected to increase the overall SAR yield from ', temp.sar.formatted, '% to ', temp.prop.sar.formatted, '%.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff >= temp.prop.eff & temp.sar >= temp.prop.sar & temp.not.int > 0 ~ paste('is expected to reduce the number of not interesting rule breaks by ', temp.not.int.formatted, '%.', sep = ''),
        length(thresholds.changed) >= 1 & temp.eff >= temp.prop.eff & temp.sar >= temp.prop.sar & temp.not.int == 0 ~ '###THIS RECOMMENDATION DOES NOT CHANGE THE EFFECTIVENESS, SAR YIELD, OR REDUCE FALSE POSITIVES. PLEASE REVIEW###',
      )
      
      # Expected rule breaks per month
      line_sixteen <- case_when(
        data.Rule.Pop.Parameter$`Num Alerts Extracted` == 0 ~ '',
        TRUE ~ 'The segment is expected to generate approximately ## rule breaks per month.')
      
      # Consolidate Narratives --------------------------------------------------
      # Populates narratives table made earlier for export
      narratives <- rbind(narratives, cbind(data.Rule.Pop.Parameter$`Rule ID`, # Rule ID
                                            data.Rule.Pop.Parameter$`Rule Name`, # Rule Name
                                            data.Rule.Pop.Parameter$`Population Group`, # Population group
                                            data.Rule.Pop.Parameter$Parameter, # Threshold tuned
                                            paste(line_one, line_two, line_three, line_four, line_five, sep = ' '), # Tuning summary 
                                            paste(line_six, line_seven, line_eight, line_nine, line_ten, line_eleven, line_twelve, line_thirteen, sep = ' '), # Analysis
                                            paste(line_fourteen, line_fifteen, line_sixteen, sep = ' '))) # Conclusion
    }}}

narratives <- narratives %>% rename(`Rule ID` = V1, `Rule Name` = V2, `Population Group` = V3, Parameter = V4, Summary = V5, Analysis = V6, Conclusion = V7) # Update column headers for export
narratives$Summary <- gsub('\\s+', ' ', narratives$Summary) # Remove any instances of extra spaces
narratives$Analysis <- gsub('\\s+', ' ', narratives$Analysis)
narratives$Conclusion <- gsub('\\s+', ' ', narratives$Conclusion)
#write.csv(narratives, paste(fp, 'Actimize Production Report Narratives.csv', sep=''), row.names = FALSE) # Export CSV to file path specified above

# Export Narratives to Word -----------------------------------------------

# Initialize output file
output <- paste(fp, report_name, sep = '') # You will have to re-save the document as a .docx file to use latest version of word
rtf <- RTF(output,width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
done(rtf)

# Create section named 'Analysis'
addHeader(rtf, 'Analysis', TOC.level = 1)

# Filter narratives to rule level
for (x in names(ruleIDs)) {
  narratives.Rule <- narratives %>% filter(`Rule ID` == x)
  
  # Add Rule Name/ID header to doc
  addHeader(rtf, title=paste(narratives.Rule$`Rule Name`, x, sep = ' | '), TOC.level = 2)
  
  # Add threshold decisions table
  addParagraph(rtf, 'Summary of Threshold Decisions')
  
  threshold.decisions <- data %>% select(`Rule ID`, `Population Group`, Parameter, `Current Threshold`, `Recommended Threshold`) # Selects relevant columns
  threshold.decisions <- threshold.decisions %>% filter(threshold.decisions$`Rule ID`==x) # Filters to rule
  threshold.decisions <- threshold.decisions %>% select(2:5) # Removes rule id from table
  threshold.decisions <- threshold.decisions[order(threshold.decisions$`Population Group`,threshold.decisions$Parameter),] # Sort table alphabetically
  threshold.decisions <- threshold.decisions %>% rename(Population = `Population Group`) # Clean up column label
  
  # Format thresholds in table to display with proper units
  for (row in 1:nrow(threshold.decisions)) {
    temp.param <- threshold.decisions$Parameter[row]
    temp.values <- c(as.numeric(threshold.decisions$`Current Threshold`[row]), as.numeric(threshold.decisions$`Recommended Threshold`[row]))
    
    if (temp.param %in% currencyF){temp.values.formatted <- c(paste('$', format(temp.values[1],big.mark = ',', nsmall = 2), sep = ''), paste('$', format(temp.values[2],big.mark = ',', nsmall = 2), sep = ''))
    } else if (temp.param %in% numberF){temp.values.formatted <- c(format(temp.values[1],big.mark = ','), format(temp.values[2],big.mark = ','))
    } else if (temp.param %in% percentF){temp.values.formatted <- c(percent(temp.values[1]/100), percent(temp.values[2]/100))
    } else if (temp.param %in% decimalF){temp.values.formatted <- c(format(round(temp.values[1], digits=2), big.mark=',', nsmall=2),format(round(temp.values[2], digits=2), big.mark=',', nsmall=2))
    } else temp.values.formatted <- temp.values 
    
    threshold.decisions$`Current Threshold`[row] <- temp.values.formatted[1]
    threshold.decisions$`Recommended Threshold`[row] <- temp.values.formatted[2]
  }
  addTable(rtf, threshold.decisions, row.names = FALSE)
  
  addNewLine(rtf)
  
  # Filter to Pop Group
  popGroups <- table(narratives.Rule$`Population Group`)
  
  for (pop in names(popGroups)) {
    narratives.Rule.Pop <- narratives.Rule %>% filter(narratives.Rule$`Population Group` == pop)
    
    # Add header for population group
    addHeader(rtf, narratives.Rule.Pop$`Population Group`, TOC.level = 3)
    
    # Add 'Threshold Recommendation' header to doc
    addHeader(rtf, title='Threshold Recommendation', TOC.level = 4)
    
    # Print summary paragraph
    addParagraph(rtf, paste(narratives.Rule.Pop$Summary[1],'\n', sep = ''))
    
    # Skip remaining headers/paragraphs if no alerts generated
    if (substr(narratives.Rule.Pop$Summary[1],1,9)=='No alerts') {
      next
    }
    
    # Filter to each parameter within pop group
    params <- table(narratives.Rule.Pop$Parameter)
    
    for (param in names(params)) {
      narratives.Rule.Pop.Parameter <- narratives.Rule.Pop %>% filter(narratives.Rule.Pop$`Parameter` == param)
      
      addHeader(rtf, narratives.Rule.Pop.Parameter$Parameter, TOC.level = 5) # Add parameter header
      addParagraph(rtf, paste(narratives.Rule.Pop.Parameter$Analysis, '\n', sep = '')) # Add analysis paragraphs for each parameter
    }
    
    # Add header for conclusion
    addHeader(rtf, title='Conclusion', TOC.level = 4)
    
    # Print conclusion paragraph
    addParagraph(rtf, paste(narratives.Rule.Pop$Conclusion[1], '\n', sep = ''))
    
  }
  # Add placeholder for scoring
  addHeader(rtf, 'Scoring Recommendation', TOC.level = 3)
}
done(rtf)