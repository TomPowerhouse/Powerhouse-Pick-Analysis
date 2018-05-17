Attribute VB_Name = "Folders"

'---------------------------- Macro Header ----------------------------
'This program analyses pick data to:
'    Create an array containing every folder with it's associated total qty, orders, pick stats
'    Determine the criteria for folders to be analyzed (e.g. Largest 25% of folders)
'    Creates an array of the selected folders
' Created by Tom Morris
' Updated May 7, 2018

' ----------------------------- Program -----------------------------

' Define Global Variables to establish location of folder #, order #, and Unit #
Public ordercol As String, foldercol As String, qtycol As String, EOcol As String, NCCol As String, TimeCol As String
Public orderrange As String, folderrange As String, qtyrange As String
 
Sub Analyze_Folders()

'------------------Variable Declaration --------------------
Dim lcount As Long
Dim upper_range As Integer, lower_range As Integer, c As Integer, n As Integer, i As Integer
'lcount -> # entries in data set
'upper/lower_range -> track the bounds of a single folder's data points
'c,n,i -> index variables

Dim checkval As Variant             'Folder number to be analyzed
Dim Q As Double                     'Lowest number to select final data
Dim Qlist() As Variant              'Array of folder selection criteria
ReDim flist(0)                      'Array of a pick run's Folder Objects
ReDim selectlist(0)                 'Array of a pick run's SELECTED folder objects

ordercol = "D"                          'Column containing order #s
foldercol = "E"                         'Folder #s
TimeCol = "K"
qtycol = "L"                            'Unit #s
EOcol = "P"                             'End of Order Tag
NCCol = "O"                             'New Carton Tag
orderrange = "D:D"                      'Column containing order #s
folderrange = "E:E"                     'Folder #s
qtyrange = "L:L"                        'Unit #s

n = 0   'folder array index
i = 2   'worksheet row index
c = 1   'Display Index (Debug only)

lcount = Application.WorksheetFunction.CountA(Columns(1))    'Determine number of entries in the dataset
upper_range = 2                                              'Denote the range of the current folder number
lower_range = 2

'--------------------------------------

Do While i < lcount + 1                'For every entry in the dataset

If (Cells(i, EOcol).Value) <> "True" Then  'If Not an End of Order

checkval = Cells(i, foldercol).Value        'Select folder cell to analyze
        
    If n = 0 Then                                   'If the folder array is empty
        flist(n).Number = checkval                  'Populate first entry
        flist(n).PDate = Cells(i, 1).Value
        c = c + 1
        n = n + 1
    ElseIf flist(n - 1).Number <> checkval Then     'Or If a new folder number is found
       ReDim Preserve flist(0 To n)                 'Increase array size and record folder #
        flist(n).Number = checkval
        flist(n).PDate = Cells(i, 1).Value
        Call Check_Order(n, lower_range, upper_range)   'Call Function to calculate pick stats for PREVIOUS folder
        n = n + 1
        c = c + 1
        upper_range = i                                 'Reset Folder 'Bounds' to new folder position
        lower_range = i
    Else
        lower_range = i                                 'If folder number is the same, increase lower bound
    End If
   
End If
i = i + 1
Loop
                Call Check_Order(n, lower_range, upper_range)   'Perform order check on last folder group
                
'Display Folder Summary Data
For i = 0 To n - 1
        Cells(i + 2, "Q").Value = flist(i).PDate
        Cells(i + 2, "R").Value = flist(i).Number
        Cells(i + 2, "S").Value = flist(i).Orders
        Cells(i + 2, "T").Value = flist(i).Qty
Next i
Call Find_Q(n)              'Call function to find minimum selection criteria
Erase flist              'Erase folder list to conserve memory

End Sub

Sub Check_Order(n As Integer, lr As Integer, ur As Integer)
'Determines the pick stats for a folder
Dim lindex As Integer, ordercount As Integer, cartoncount As Integer
Dim newctime As Long, oetime As Long, timesum As Long, picktime As Long
Dim pickqty As Long
Dim order As Long

uindex = ur
ordercount = 0
cartoncount = 0
timesum = 0
picktime = 0
newctime = 0
pickqty = 0

    Do While uindex < lr + 2                                'Scan all elements in Folder group
    If Cells(uindex, EOcol).Value <> "True" Then            'If Not End of Order tag
    
        If Cells(uindex, NCCol).Value = "True" Then         'If pick is a New Carton
            cartoncount = cartoncount + 1                   'Count carton
            newctime = newctime + Cells(uindex, TimeCol).Value  'Add to new carton time
            ordercount = ordercount + 1
        'End If  'End of new carton if
        
        ElseIf ordercount = 0 Or Cells(uindex, ordercol) <> order Then                              'If first order, count order
          '  ordercount = ordercount + 1
          '  order = Cells(uindex, ordercol)                 'record order#
          '  newctime = newctime + Cells(uindex, TimeCol).Value    'Add to new carton time (First carton in an order is new, but not tagged)
       ' ElseIf Cells(uindex, ordercol) <> order Then        'If a new Order# is found, count order
            ordercount = ordercount + 1
            order = Cells(uindex, ordercol)
            cartoncount = cartoncount + 1                   'Count carton
            newctime = newctime + Cells(uindex, TimeCol).Value  'Add to new carton time
        Else 'If Cells(uindex, NCCol).Value = "True" Then
            picktime = picktime + Cells(uindex, TimeCol).Value
            pickqty = pickqty + Cells(uindex, qtycol).Value
        End If  'End of orders if
        
        flist(n - 1).Qty = flist(n - 1).Qty + Cells(uindex, qtycol) 'For any Order found, sum units
        
    Else        'If pick IS the end of order
        oetime = oetime + Cells(uindex, TimeCol).Value              'Sum end of order time
    End If            'End of Order-End IF
            timesum = timesum + Cells(uindex, TimeCol).Value        'Sum entire time taken
            uindex = uindex + 1                                     'increment index to scan next row
    Loop
    
    With flist(n - 1)
            .UPP = Round(flist(n - 1).Qty / ordercount, 0)    'Units per PO
            .Orders = ordercount                              'Record Folder object's total orders
                If pickqty <> 0 Then
                    .URate = picktime / pickqty                        'Average time/unit
                End If
            .JRate = timesum / flist(n - 1).Qty               'Total time/unit
            .NCTime = newctime / cartoncount                  'Average time taken for a new carton
            .EOTime = oetime / ordercount                     'Average time taken at end of an order
    End With
    
End Sub

Sub Find_Q(n As Integer)
'Finds mimimum folder selection Criteria
Dim a As Integer, size As Integer
size = 0
ReDim Qlist(0 To n)

For a = 0 To n - 1
    Qlist(a) = flist(a).Qty                     'Transfer folder total quantities to a separate array to be analyzed
Next a

Q = WorksheetFunction.Quartile_Inc(Qlist(), 3)  'Compare total units in each folder, find third quartile
Q = WorksheetFunction.RoundUp(Q, 0)             'Round to whole number of units

For a = 0 To n - 1          'For all folders in the data series
    
    If Qlist(a) > Q Or Qlist(a) = Q Then        'If the total units for a folder is equal or above criteria
    ReDim Preserve selectlist(0 To size)
        With selectlist(size)
            .Number = flist(a).Number
            .Orders = flist(a).Orders
            .Qty = flist(a).Qty
            .PDate = flist(a).PDate
            .JRate = flist(a).JRate
            .UPP = flist(a).UPP
            .URate = flist(a).URate
            .EOTime = flist(a).EOTime
            .NCTime = flist(a).NCTime
            
        End With
        size = size + 1
        'Display Data
        Cells(size + 1, "U").Value = selectlist(size - 1).PDate
        Cells(size + 1, "V").Value = selectlist(size - 1).Number
        Cells(size + 1, "W").Value = selectlist(size - 1).Orders
        Cells(size + 1, "X").Value = selectlist(size - 1).Qty
    End If
Next a
Cells(1, "Q").Value = Q
Erase Qlist
End Sub


