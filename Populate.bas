Attribute VB_Name = "Populate"
Public flist() As New Folder  'Array of all Folders
Public selectlist() As New Folder

Sub Populate_Data()

Dim i As Integer, n As Integer, size As Integer, lcount As Integer
Dim sourcecount As Integer, destcount As Integer
Dim destbook As Workbook, sourcebook As Workbook
Dim destsheet As Worksheet, sourcesheet As Worksheet
Dim folderdata() As Folder
Dim sheetexists As Boolean
Dim DataSet As Range
Set sourcebook = ActiveWorkbook
Set destbook = Workbooks.Open("X:\Optimization\Analysis\Pick Run Analysis.xlsm")    'Open worksheet to place data
sourcecount = sourcebook.Worksheets.Count                         'Count Worksheets                                                 'Initialize Table Indices
destcount = destbook.Worksheets.Count
i = 2

Do While i < sourcecount                                         'For each worksheet in the workbook
    Set sourcesheet = sourcebook.Worksheets(i + 1)                     'Set the source data worksheet
    sourcesheet.Activate
    Call Analyze_Folders
    sheetexists = False
    destbook.Activate
     For Each Sheet In destbook.Worksheets
            If Sheet.Name = sourcesheet.Name Then
                sheetexists = True
                Set destsheet = Sheet
                Exit For
            End If
     Next Sheet
       
 'Create New Worksheet if Needed
        If sheetexists = False Then
            Worksheets("Template").Copy After:=Worksheets(destcount)
            ActiveSheet.Name = sourcesheet.Name
            ActiveSheet.Cells(2, "B").Value = sourcesheet.Cells(2, "F").Value
            Set destsheet = ActiveSheet
            Cells(1, 1) = ActiveSheet.Name
        End If
        
  'Populate Sheet with Selected Data
          size = UBound(selectlist) - LBound(selectlist)
        destsheet.Activate
        ActiveSheet.Cells(3, "B").Value = Date
        lcount = destsheet.Cells(2, Columns.Count).End(xlToLeft).Column + 1
        For n = 0 To size
            destsheet.Cells(1, n + lcount).Value = selectlist(n).PDate
            destsheet.Cells(2, n + lcount).Value = selectlist(n).Number
            destsheet.Cells(3, n + lcount).Value = selectlist(n).Qty
            destsheet.Cells(4, n + lcount).Value = selectlist(n).UPP
            destsheet.Cells(5, n + lcount).Value = selectlist(n).URate
            destsheet.Cells(6, n + lcount).Value = selectlist(n).JRate
            destsheet.Cells(7, n + lcount).Value = selectlist(n).NCTime
            destsheet.Cells(8, n + lcount).Value = selectlist(n).EOTime
        Next n

    Set DataSet = Range(Range("D1", Range("D1").End(xlToRight)), Range("D1", Range("D8")))
    DataSet.Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("D1", Range("D1").End(xlToRight)), _
        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange DataSet
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With

   Call Update
   
 i = i + 1 'Increment indices
Loop
End Sub
