Sub Macro1()
'
' Macro1 Macro
'
    Dim Wkbk As Workbook
    Dim CombineSheet As Worksheet
    Dim TableSheet As Worksheet
    Dim WS_Count As Integer
    
    Set Wkbk = ActiveWorkbook
    WS_Count = Wkbk.Worksheets.Count
'
    Sheets.Add.Name = "CombineData"
    Set CombineSheet = Sheets("CombineData")
    WS_Count = Wkbk.Worksheets.Count
    
    CombineSheet.Activate
    
    ActiveSheet.Next.Select
    Rows("1:1").Select
    Selection.Copy
    CombineSheet.Select
    Range("A1").Select
    ActiveSheet.Paste
    
    For I = 1 To WS_Count
    
    Wkbk.Worksheets(I).Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    CombineSheet.Select
    If Range("A2") > 0 Then
        Range("A2").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1).Select
        ActiveSheet.Paste
    Else
        Range("A2").Select
        ActiveSheet.Paste
        End If
        Next
        
    CombineSheet.Activate
    Cells.Select
    Selection.Columns.AutoFit
    
    
End Sub
