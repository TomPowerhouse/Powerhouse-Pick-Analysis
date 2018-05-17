Attribute VB_Name = "Convert"
Sub Convert_Data()

'---------------------------- Macro Header ----------------------------

' This program modifies data pulled from the Pick Analysis Set Viewer
' to be more readily manipulated for analysis

' Created by Tom Morris
' Updated May 7, 2018


'-----------------Establishing Variables---------------------------------

    Dim i As Integer        'Row index for Pick Run Mapping
    i = 2

' ----------------------------- Program --------------------------------

'Columns AS IMPORTED
'|Trans Date| Operator| Order | Folder | Division| Item| Description| Location| Time| Qty| Full Carton| New Carton| End Order|

' Turn off dialogue boxes to automatically overwrite column data
    Application.DisplayAlerts = False
    
  'Sort Data by Folder and Order
    Range("A1", Cells(ActiveCell.End(xlDown).Row, ActiveCell.End(xlToRight).Column)).Select
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "D:D"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "C:C"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "A:A"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "B:B"), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
        
    With ActiveSheet.Sort
        .SetRange Range("A1", Cells(ActiveCell.End(xlDown).Row, ActiveCell.End(xlToRight).Column))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
 ' Adding Rate-of-Pick Column
    Columns("K:K").Select           'Make space for Time/Pick
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Dim c As Integer
    c = 2
    Do Until IsEmpty(Cells(c, 1).Value) = True
        If IsEmpty(Cells(c, "J")) = False Then
            Cells(c, "K").Value = (Cells(c, "I").Value) / (Cells(c, "J").Value)
        End If
    c = c + 1
    Loop
        
    Columns("K:K").Select
    Selection.NumberFormat = "0"
    
' Splitting transaction time into distinct date and time columns
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.NumberFormat = "hh:mm:ss"
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1), Array(22, 1)), TrailingMinusNumbers _
        :=True
    Selection.NumberFormat = "m/d/yyyy"
    
' Removing "Powerhouse\" from operator name for legibility
    Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="\", FieldInfo:=Array(Array(1, 9), Array(2, 1)), TrailingMinusNumbers:=True
        
Application.ScreenUpdating = False
'Add Pick Run Column
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Do Until IsEmpty(Cells(i, 1).Value) = True
        If IsEmpty(Cells(i, "I").Value) = False Then
            Cells(i, "J").Value = LocationMap(Cells(i, "I").Value)
        End If
    i = i + 1
    Loop
    i = 2
    Do Until IsEmpty(Cells(i, 1).Value) = True
        If IsEmpty(Cells(i, "J").Value) = True Then
            Cells(i, "J").Value = Cells(i - 1, "J").Value
        End If
    i = i + 1
    Loop
    Application.ScreenUpdating = True
 ' Modifying Data Column Headers
    Cells(1, "A").Value = "Date"
    Cells(1, "B").Value = "Time"
    Cells(1, "C").Value = "Operator"
    Cells(1, "E").Value = "Folder"
    Cells(1, "M").Value = "Time/Pick"
    Cells(1, "J").Value = "Pick Run"
    Columns("A:P").EntireColumn.AutoFit
    
    'Clean Up New Carton and Order End tags
        Columns("N:P").Select
        Selection.Replace What:="False", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
End Sub
