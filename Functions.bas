Attribute VB_Name = "Functions"
Sub Update()
Dim UnitTotal As Integer

'Average Total Units
Cells(10, "D").Value = Round(Application.WorksheetFunction.Average(Range("D3", Range("D3").End(xlToRight))), 0)

'Average Units/PO
Cells(11, "D").Value = Round(Application.WorksheetFunction.Average(Range("D4", Range("D4").End(xlToRight))), 0)

'Average Pick Sec/Unit
Cells(12, "D").Value = Round(Application.WorksheetFunction.Average(Range("D5", Range("D5").End(xlToRight))), 0)

'Average Job Sec/Unit
Cells(13, "D").Value = Round(Application.WorksheetFunction.Average(Range("D6", Range("D6").End(xlToRight))), 0)

'Average Sec/New Carton
Cells(14, "D").Value = Round(Application.WorksheetFunction.Average(Range("D7", Range("D7").End(xlToRight))), 0)
'Average Sec/Order End
Cells(15, "D").Value = Round(Application.WorksheetFunction.Average(Range("D8", Range("D8").End(xlToRight))), 0)

'Data Points
Cells(16, "D").Value = Application.WorksheetFunction.CountA(Range("D2", Range("D2").End(xlToRight)))

End Sub

Function LocationMap(location As String)
Dim lcount As Integer
Dim Col1, Col2, Col3 As String

If Len(location) = 5 Then
    Col1 = "A"
    Col2 = "B"
    Col3 = "C"
ElseIf Len(location) > 5 Then
    Col1 = "E"
    Col2 = "F"
    Col3 = "G"
End If

Worksheets("Location Maps").Activate
lcount = Application.WorksheetFunction.CountA(Worksheets("Location Maps").Range(Col1 & 3, Range(Col1 & 3).End(xlDown)))    'Number of location groups
Worksheets("Data").Activate

    For n = 3 To lcount
        If Worksheets("Location Maps").Cells(n, Col1).Value < location _
            Or Worksheets("Location Maps").Cells(n, Col1).Value = location Then
            
            If Worksheets("Location Maps").Cells(n, Col2).Value > location _
                Or Worksheets("Location Maps").Cells(n, Col2).Value = location Then
                
                pickrun = Worksheets("Location Maps").Cells(n, Col3).Value
                LocationMap = pickrun
                Exit For
            End If
        End If
    Next n
    If pickrun = Empty Then
        LocationMap = "Other"
    End If
    pickrun = Empty
End Function

    
' Save data to a new worksheet
'Range("A1", Cells(ActiveCell.End(xlDown).Row, ActiveCell.End(xlToRight).Column)).Copy

'Set wb = Workbooks.Add
'wb.Activate
'wb.Sheets(1).Range("A1:J1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks _
':=False, Transpose:=False
'Application.CutCopyMode = False
'Columns("A:J").EntireColumn.AutoFit

'Dim a As String, b As String
'a = Range("A" & Cells.Rows.Count).End(xlUp)

'wb.SaveAs ("X:\Optimization\Analysis\Pick Data " & Format(Now(), "DD-MMM-YYYY") & ".xlsm"), FileFormat:=52

'------------------------------------

' Sort by ascending chronology
 '   Range("A1", Cells(ActiveCell.End(xlDown).Row, ActiveCell.End(xlToRight).Column)).Select
 '   ActiveSheet.Sort.SortFields.Add Key:=Range( _
 '       "A:A"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
 '       xlSortNormal
  '  With ActiveSheet.Sort
  '      .SetRange Range("A1", Cells(ActiveCell.End(xlDown).Row, ActiveCell.End(xlToRight).Column))
  '      .Header = xlYes
   '     .MatchCase = False
   '     .Orientation = xlTopToBottom
   '     .SortMethod = xlPinYin
   '     .Apply
   ' End With

