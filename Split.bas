Attribute VB_Name = "Split"
Sub Split_Data()
Dim checkval As String

Dim i As Integer, n As Integer, index As Integer, lcount As Integer, b As Integer
Dim destsheet As Worksheet, sourcesheet As Worksheet
Dim sheetexists As Boolean
Dim datarow As Range, destrow As Range

    Dim cell As Range
    
Application.ScreenUpdating = False
Set sourcesheet = Worksheets(1)  'Set the source data worksheet
i = 2
index = 2

Do Until IsEmpty(sourcesheet.Cells(i, 1).Value) = True

checkval = sourcesheet.Cells(i, "J").Value
    sheetexists = False

     For Each Sheet In Worksheets
            If Sheet.Name = checkval Then
                sheetexists = True
                Set destsheet = Sheet
                index = Application.WorksheetFunction.CountA(destsheet.Columns(1)) + 1
                Exit For
            End If
     Next Sheet
       
 'Create New Worksheet if Needed
        If sheetexists = False Then
            Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = checkval
            Set destsheet = ActiveSheet
            sourcesheet.Activate
            Worksheets(checkval).Range("A1:P1").Value = sourcesheet.Range("A1:P1").Value
            index = 2
        End If

sourcesheet.Activate
    For b = 1 To 16
        destsheet.Cells(index, b).Value = sourcesheet.Cells(i, b).Value
    Next b
    
i = i + 1
Loop

For Each Sheet In Worksheets
    Sheet.Activate
    Columns("A:P").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.NumberFormat = "hh:mm:ss"
Next Sheet
Application.ScreenUpdating = True
End Sub
