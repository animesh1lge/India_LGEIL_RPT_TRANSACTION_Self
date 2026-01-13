Sub MergeFilesInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim ws As Worksheet
    Dim masterWB As Workbook
    Dim masterWS As Worksheet
    Dim lastRow As Long
    
    folderPath = "Z:\RPA_Repository_SELFRPA\LGEIL\Animesh Singh\INDIA_LGEIL_receivng-Data_transaction amount_PUNE_SELF\temp\PU SNS" ' Adjust the folder path as needed
    Set masterWB = Workbooks.Add
    Set masterWS = masterWB.Sheets(1) ' Use the first sheet in the master workbook
    
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        
        For Each ws In wb.Sheets
            lastRow = masterWS.Cells(masterWS.Rows.Count, "A").End(xlUp).Row
            ws.UsedRange.Copy Destination:=masterWS.Range("A" & lastRow + 1)
        Next ws
        
        wb.Close False
        fileName = Dir
    Loop
    
    Application.DisplayAlerts = False
    masterWB.SaveAs "Z:\RPA_Repository_SELFRPA\LGEIL\Animesh Singh\INDIA_LGEIL_receivng-Data_transaction amount_PUNE_SELF\DATA\Assy Changed By Sales Price.xlsx"
    masterWB.Close
    Application.DisplayAlerts = True
End Sub