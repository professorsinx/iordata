Steps:
(1)
Place the downloaded BAL & ENT xlsx files inside the input folder.

Mode                 LastWriteTime         Length Name                                                                                                                                                               
----                 -------------         ------ ----                                                                                                                                                                                                                                                                                                                
-a----        29-01-2025     18:19         459204 BAL_WISE_DATA.xlsx                                                                                                                                                                                                                                                                                                  
-a----        29-01-2025     18:19         830094 ENT_WISE_DATA.xlsx  

                                                                                                                                               
(2) 
run step_0_convert_xlsx_to_csv_ver_2.bat

This creates new .csv files inside the .\input folder for further processing

Mode                 LastWriteTime         Length Name                                                                                                                                                               
----                 -------------         ------ ----                                                                                                                                                               
-a----        31-01-2025     15:10        1006283 BAL_WISE_DATA.csv                                                                                                                                                  
-a----        29-01-2025     18:19         459204 BAL_WISE_DATA.xlsx                                                                                                                                                 
-a----        31-01-2025     15:10        1876551 ENT_WISE_DATA.csv                                                                                                                                                  
-a----        29-01-2025     18:19         830094 ENT_WISE_DATA.xlsx                                                                                                                                                 


(3) 
run step_1_generate_merged_files_v_100.ps1 
by opening the Powershell ISE and then opening the ps1 file and selecting the code and clicking "Run selection" if running scripts
are blocked in your PC.

This creates Merged_Output.csv inside .\output folder


Mode                 LastWriteTime         Length Name                                                                                                                                                               
----                 -------------         ------ ----                                                                                                                                                               
-a----        31-01-2025     16:14         360070 Merged_Output.csv                                                                                                                                                  


(4) 
run step_2_generate_filtered_files.ps1
by opening the Powershell ISE and then opening the ps1 file and selecting the code and clicking "Run selection" if running scripts
are blocked in your PC.

This creates 8 new files inside .\output folder

Mode                 LastWriteTime         Length Name                                                                                                                                                               
----                 -------------         ------ ----                                                                                                                                                               
-a----        31-01-2025     16:15           2211 Filtered_BGL_4597998.csv                                                                                                                                           
-a----        31-01-2025     16:15            439 Filtered_BGL_4599635.csv                                                                                                                                           
-a----        31-01-2025     16:15           2016 Filtered_BGL_4897932.csv                                                                                                                                           
-a----        31-01-2025     16:15         301569 Filtered_Merged_Output.csv                                                                                                                                         
-a----        31-01-2025     16:15          61221 Filtered_SL_NO_1_2_DaysPassed_GT_20.csv                                                                                                                            
-a----        31-01-2025     16:15           2469 Filtered_SL_NO_9.csv                                                                                                                                               
-a----        31-01-2025     16:15          23089 Filtered_SL_NO_NE_9_Overdue_GE_0.csv                                                                                                                               
-a----        31-01-2025     16:15          34029 Filtered_TAT_NE_45_Overdue_NEG_123.csv                                                                                                                             
-a----        31-01-2025     16:14         360070 Merged_Output.csv 

(5) 
run step_3_merge_csv_to_excel.bat
by double clicking

This will merge all Filtered*.csv files in the .\output folder and merges and creates an MergedExcel.xlsx 

Mode                 LastWriteTime         Length Name                                                                                                                                                               
----                 -------------         ------ ----                                                                                                                                                               
-a----        31-01-2025     16:18         239367 MergedExcel.xlsx                                                                                                                                                   
-a----        31-01-2025     16:14         360070 Merged_Output.csv 


Notes:
The .bat files that convert the csv to xlsx and xlsx to csv use .vbs as helper script. Don't delete them. 