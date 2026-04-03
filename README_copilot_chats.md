slots can be from 1-25. update the code to account for that

the 'wfr' order is important for both TEST and POR data and corresponds to a COATER. so create a COATER column for each data set. For POR data, each run is a unique LOT and LOT_DATA_COLLECT_DATE combo. There should be 3 wfrs for each run. the first wfr COAT=1, 2nd COAT=2, 3rd COAT=3. For TEST data, there is 8 COATER. So 1st wfr COAT=1301, 2nd wfr COAT=1401, 3rd wfr COAT=1303, 4th wfr COAT=1403, 5th COAT=1302, 6th COAT=1402, 7th COAT=1304, 8th COAT=1404. the 1301, 1401, 1303, 1403, 1302, 1402, 1304, 1404 are text categories not integers

for TEST it's not the slot but the order of the data in the xml that matters. so for each xml file 1st wfr should map COAT=1301, 2nd wfr COAT=1401, 3rd wfr COAT=1303, 4th wfr COAT=1403, 5th COAT=1302, 6th COAT=1402, 7th COAT=1304, 8th COAT=1404. with 1st, 2nd, 3rd etc being the order of the wafer

the TEST entity should be hard coded as 'TZJ501'. and make it so i can change it in the future easily

make them easily configurable

A run for POR is each unique LOT_DATA_COLLECT_DATE and COATER. A run for TEST is each unique File Name and COATER. I want a summary table for the statistics for the runs for each ENTITY and COATER (will be comparing TZJ501 to the other, TTG624, TTG625, etc). There should be 6 statistics for each ENTITY COATER combination. Mean thickness and std dev of Mean thickness across the runs, the average delta from run to run and the std dev of the deltas, the average std dev of each run and the std dev of the std dev of each run

update TEST Data Time Series to color by COATER not by Wafer. connect the means of the data. 

but each x axis point should be one of the File Name Date/Time. There are too large of gaps between date/time right now. It can be more of a category or discrete points than literal date/time

copy the TEST Date Time Series and create another chart that is grouped by DECK. so COATER 1301, 1302 go to DECK=13-L, 1303, 1304 go to DECK=13-R, 1401,1402 go to DECK=14-L, 1403,1404 go to DECK=14-R

for the TEST data add a spline where it plots the thickness by DECK. X axis is radius from 0-150mm. and add a drop down so i can select one or multiple 'File Name' to plot. 

add an option to export the TEST data to excel. it should have Statistical Summary by COATER. each row should be data from one of the File Name. So there should be a row for each File Name in TEST. There should be 24 columns. 8 COAT and 3 columns for each COAT. The columns should be 1301 Mean Thickness (Å), 1301 Avg delta to previous run (will be blank for first row), 1301 std dev. then next 3 columns are same columns but for 1401, then 1303, 1403, 1302, 1402, 1304, 1404. 