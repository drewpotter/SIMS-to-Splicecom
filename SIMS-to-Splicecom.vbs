Dim i As Integer
Dim z As Integer

Dim studentNamesColumn As Integer
Dim parentNamesColumn As Integer
Dim numbersColumn As Integer
Dim priorityColumn As Integer
Dim cellString As String

z = ActiveWorkbook.ActiveSheet.UsedRange.Rows.Count
parentsNamesColum = 1
studentNamesColumn = 2
numbersColumn = 4
priorityColumn = 3

' Column names
ActiveWorkbook.Worksheets("Sheet2").Cells(1, 1).Value = "Contact Name"
ActiveWorkbook.Worksheets("Sheet2").Cells(1, 2).Value = "Company Name"
ActiveWorkbook.Worksheets("Sheet2").Cells(1, 3).Value = "Description"
ActiveWorkbook.Worksheets("Sheet2").Cells(1, 4).Value = "Phone"
ActiveWorkbook.Worksheets("Sheet2").Cells(1, 5).Value = "Mobile"

' Concatenating the Surname and Forename
For i = 2 To z
     ActiveWorkbook.Worksheets("Sheet2").Cells(i, 1).Value = Cells(i, 2).Value + " " + Cells(i, 3).Value
Next i

' Fix the blank names
For i = 2 To z
    cellString = Cells(i, 1).Value
    
    ActiveWorkbook.Worksheets("Sheet2").Cells(i, studentNamesColumn).Value = Cells(i, 1).Value
    
    If cellString & "" = "" Then
        ActiveWorkbook.Worksheets("Sheet2").Cells(i, studentNamesColumn).Value = Cells(i - 1, 1).Value
    End If
    
Next i

' Fix the priority
For i = 2 To z
    cellString = Cells(i, 6).Value
    
    If cellString = "1" Or cellString = "2" Then
        ActiveWorkbook.Worksheets("Sheet2").Cells(i, priorityColumn).Value = "Priority " + Cells(i, 6).Value
    End If
    
Next i

' Fix the numbers
For i = 2 To z
    cellString = Cells(i, 4).Value

    ActiveWorkbook.Worksheets("Sheet2").Cells(i, numbersColumn).Value = "'" + CStr(Cells(i, 4).Value)
    ActiveWorkbook.Worksheets("Sheet2").Cells(i, numbersColumn + 1).Value = "'" + CStr(Cells(i, 5).Value)
    
    If cellString & "" = "" Then
        
        ActiveWorkbook.Worksheets("Sheet2").Cells(i, 4).Value = CStr(Cells(i, 5).Value)
        ActiveWorkbook.Worksheets("Sheet2").Cells(i, 5).Value = ""
       
    End If
    
Next i