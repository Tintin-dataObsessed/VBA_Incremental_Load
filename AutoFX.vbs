Option Explicit

' Define constants for log file path and Excel file path
Const logFilePath = "C:\Users\MichelleChekwooti\OneDrive - 637 Capital\Desktop\AutoFX.bat"' Update with your log file path
Const excelFilePath = "C:\Users\MichelleChekwooti\OneDrive - 637 Capital\Documents - LIMBUA_Confidential\sales dashboards beta\database\Fx_Rates_new - Copy.xlsm"

' Create a FileSystemObject for logging
Dim fso, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.OpenTextFile(logFilePath, 8, True)

' Log function
Sub LogMessage(msg)
    logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

' Start logging
LogMessage "Script started"

' Create the Excel application object
Dim xlsxApp
Set xlsxApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    LogMessage "Error creating Excel application: " & Err.Description
    logFile.Close
    WScript.Quit
End If
LogMessage "Excel application created successfully"

xlsxApp.Visible = False ' Set to False for background execution

' Open the workbook
Dim objWorkbook
Set objWorkbook = xlsxApp.Workbooks.Open(excelFilePath)
If Err.Number <> 0 Then
    LogMessage "Error opening workbook: " & Err.Description
    xlsxApp.Quit
    logFile.Close
    WScript.Quit
End If
LogMessage "Workbook opened successfully"

' Run the macro
xlsxApp.Run "RefreshAndAppendData"
If Err.Number <> 0 Then
    LogMessage "Error running macro: " & Err.Description
Else
    LogMessage "Macro executed successfully"
End If

' Save and close the workbook
objWorkbook.Close True
If Err.Number <> 0 Then
    LogMessage "Error closing workbook: " & Err.Description
Else
    LogMessage "Workbook closed successfully"
End If

' Quit the Excel application
xlsxApp.Quit
If Err.Number <> 0 Then
    LogMessage "Error quitting Excel application: " & Err.Description
Else
    LogMessage "Excel application quit successfully"
End If

' Release the objects
Set objWorkbook = Nothing
Set xlsxApp = Nothing

' Close the log file
logFile.Close

' End logging
LogMessage "Script ended"

On Error GoTo 0
