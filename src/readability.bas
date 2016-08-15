Sub Readability()
    Dim pData As Workbook
    Dim CCDG As Worksheet
    
    Set pData = ActiveWorkbook
    Set CDGD = Sheets("Circuit data gamma drawing")
	
	'Activate CDGD
	Worksheets(CDGD).Activate
    
    'This cleans up useless columns to increase readability
    Range("E:E, F:F, J:J, K:K, P:P, Q:Q, R:R, S:S, U:U, V:V, W:W, X:X, Y:Y, AL:AL, AM:AM, AN:AN").Delete
    
    'This creates a new column for concantenating string
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    
    'This creates concantenation of wire type
    Dim lngLastRow As Long
    
    'Uses Column B to set the 'lngLastRow' variable _
    (find the last row) - change if required.
    lngLastRow = Cells(Rows.Count, "C").End(xlUp).Row

    Range("K5").Formula = "=I5&J5"
    Range("K5").Copy Range("K6:K" & lngLastRow)
    
    'Hides Source concantenation column
    Columns("I").Hidden = True
    Columns("J").Hidden = True
    
    'Resize Columns to Autofit
    Range("K:K").Columns.AutoFit
    
    'Rename Column
    Range("K4").Value = "Wire_Type"
    
    'Changes wire length to meter
    Dim element As Range
    Dim MaxRows As Long

    With Worksheets("Circuit data gamma drawing")
        MaxRows = .Cells(.Rows.Count, "M").End(xlUp).Row
    End With

    
    For Each element In Worksheets("Circuit data gamma drawing").Range("M5:M" & MaxRows)
        If IsNumeric(element.Value) Then
            element.Value = element.Value / 1000
        End If
    Next

    'Changes Column N for Left(8) of its value
    With Worksheets("Circuit data gamma drawing")
        MaxRows = .Cells(.Rows.Count, "N").End(xlUp).Row
    End With

    
    For Each element In Worksheets("Circuit data gamma drawing").Range("N5:N" & MaxRows)
        If IsNumeric(element.Value) Then
            element.NumberFormat = "@"
            element.Value = Left(element.Value, 8)
        End If
    Next
    
    'Changes Column P for Left(8) of its value
    With Worksheets("Circuit data gamma drawing")
        MaxRows = .Cells(.Rows.Count, "P").End(xlUp).Row
    End With

    
    For Each element In Worksheets("Circuit data gamma drawing").Range("P5:P" & MaxRows)
        If IsNumeric(element.Value) Then
            element.NumberFormat = "@"
            element.Value = Left(element.Value, 8)
        End If
    Next
    
    'Changes Column R for Left(8) of its value
    With Worksheets("Circuit data gamma drawing")
        MaxRows = .Cells(.Rows.Count, "R").End(xlUp).Row
    End With

    
    For Each element In Worksheets("Circuit data gamma drawing").Range("R5:R" & MaxRows)
        If IsNumeric(element.Value) Then
            element.NumberFormat = "@"
            element.Value = Left(element.Value, 8)
        End If
    Next
    
    'Changes Column T for Left(8) of its value
    With Worksheets("Circuit data gamma drawing")
        MaxRows = .Cells(.Rows.Count, "T").End(xlUp).Row
    End With

    
    For Each element In Worksheets("Circuit data gamma drawing").Range("T5:T" & MaxRows)
        If IsNumeric(element.Value) Then
            element.NumberFormat = "@"
            element.Value = Left(element.Value, 8)
        End If
    Next
    
    'Changes Column V for Left(8) of its value
    With Worksheets("Circuit data gamma drawing")
        MaxRows = .Cells(.Rows.Count, "V").End(xlUp).Row
    End With

    
    For Each element In Worksheets("Circuit data gamma drawing").Range("V5:V" & MaxRows)
        If IsNumeric(element.Value) Then
            element.NumberFormat = "@"
            element.Value = Left(element.Value, 8)
        End If
    Next
    
    'Changes Column X for Left(8) of its value
    With Worksheets("Circuit data gamma drawing")
        MaxRows = .Cells(.Rows.Count, "X").End(xlUp).Row
    End With

    
    For Each element In Worksheets("Circuit data gamma drawing").Range("X5:X" & MaxRows)
        If IsNumeric(element.Value) Then
            element.NumberFormat = "@"
            element.Value = Left(element.Value, 8)
        End If
    Next
    
End Sub
