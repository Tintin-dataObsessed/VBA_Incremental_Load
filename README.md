# VBA_Incremental_Load
This was an experiment to create an incremental load on Excel which would allow me to update Exchange rates automatically at a certain time of the day.
 # Step_1
 Create a VBA Script file with the code input in this file. The VBA Script finds the last row in the 'historical' table then finds the last row updating table sourcing from an RSS feed. It copies the data on the feed to that of the historical table when the script is run. This is the FX_Rates_New-Copy.xlsm
 Usually Visual Basic modules are attached to file in this case the .xlsm but I extracted the .bat file incase.
 # Step_2
 Create a Batch file, which is a list of plain text commands you would like executed.This is the AutoFX file which will be opened by the task scheduler.
 # Step_3
 Use Task scheduler to Create a Task that can be run at certain times during the day.

 The whole process has been documented further on my blog: https://dataobsessed4.wordpress.com/
