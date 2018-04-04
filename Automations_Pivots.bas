Attribute VB_Name = "Module1"
Sub FormatReport()
'
' Janina Umpa / janina.umpa@hp.com
' ES ITO Global Service Management
''APJ Service Level and Information Management - Operational Reporting
'
Application.ScreenUpdating = False

'(0) Check if Report tab exists and if it is already formatted.
Dim sh As Worksheet, flg As Boolean
For Each sh In Worksheets
If sh.Name Like "Report*" Then flg = True: Exit For
Next
If flg = True Then
Worksheets("Report").Select
Else
Worksheets("Interface").Select
MsgBox "Report tab does not exist. Copy tab from source file first!"
End
End If

Range("A1").Select
If ActiveCell.Value = vbNullString Then
    Worksheets("Interface").Select
    MsgBox "Report tab is already formatted or in a non-standard format!"
    End
Else
    Range("A1").Select
End If

'(1) This part will format Report tab
''Declare variables
    Dim intRow As Integer, i As Integer, j  As Integer

''(1.1) Delete rows 1 to 19
    Rows("1:19").Select
    Range("A19").Activate
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
''(1.2) Unmerge all cells
    Cells.Select
    With Selection
        .MergeCells = False
    End With

''(1.3) Fill empty cells from unmerged cells with information (Row 6 through last row, Columns A through M)
'''(1.3.1) Identify last row of data
        intRow = Cells(Cells.Rows.Count, "A").End(xlUp).Row - 2
'''(1.3.2) Double Loop: Movement --> Up to Down, Move Right, Up to Down again
                     ': Logic --> If cell is not blank, keep cell data. If cell is blank, copy data from the cell one row above it.
                     ': Legend --> i is row number, j is column number. i.e., Cells(i,j) = Range("B1") if i = 1 and j = 2
        For j = 1 To 13
            For i = 6 To intRow
                If Cells(i, j) <> "" Then
                    Cells(i, j) = Cells(i, j)
                ElseIf Cells(i, j) = "" Then
                    Cells(i, j) = Cells(i - 1, j)
                End If
            Next i
        Next j

Worksheets("interface").Select

MsgBox "Your Report tab is now pivot-ready!"
'Your raw data in Report tab should now be "pivot-ready".

Application.ScreenUpdating = True

End Sub


 
 Sub PivotHeaders()
 
Application.ScreenUpdating = False

Dim Month1 As String, Month2 As String, Month3 As String, Month4 As String, Month5 As String, Month6 As String
Dim Month7 As String, Month8 As String, Month9 As String, Month10 As String, Month11 As String, Month12 As String

Worksheets("Report").Activate
'Counting Columns
    Dim intCol As Integer
    intCol = Worksheets("Report").Cells(4, Columns.Count).End(xlToLeft).Column
    
Month1 = Worksheets("Report").Range("R3")
Month2 = Worksheets("Report").Range("U3")
Month3 = Worksheets("Report").Range("X3")
Month4 = Worksheets("Report").Range("AA3")
Month5 = Worksheets("Report").Range("AD3")
Month6 = Worksheets("Report").Range("AG3")
Month7 = Worksheets("Report").Range("AJ3")
Month8 = Worksheets("Report").Range("AM3")
Month9 = Worksheets("Report").Range("AP3")
Month10 = Worksheets("Report").Range("AS3")
Month11 = Worksheets("Report").Range("AV3")
Month12 = Worksheets("Report").Range("AY3")

Worksheets("Pivot").Activate
''1
Worksheets("Pivot").Range("C4:E4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .MergeCells = True
        .Interior.Color = RGB(184, 204, 228)
        .NumberFormat = "yy/mmm"
    End With
    Range("C4").Value = Month1
    
''2
If intCol > 20 Then
    Worksheets("Pivot").Range("F4:H4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("F4").Value = Month2
End If

''3
If intCol > 23 Then
    Worksheets("Pivot").Range("I4:K4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("I4").Value = Month3
End If

''4
If intCol > 26 Then
    Worksheets("Pivot").Range("L4:N4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("L4").Value = Month4
End If

''5
If intCol > 29 Then
    Worksheets("Pivot").Range("O4:Q4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("O4").Value = Month5
End If

''6
If intCol > 32 Then
    Worksheets("Pivot").Range("R4:T4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("R4").Value = Month6
End If

''7
If intCol > 35 Then
    Worksheets("Pivot").Range("U4:W4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("U4").Value = Month7
End If

''8
If intCol > 38 Then
    Worksheets("Pivot").Range("X4:Z4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("X4").Value = Month8
End If
    
''9
If intCol > 41 Then
    Worksheets("Pivot").Range("AA4:AC4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("AA4").Value = Month9
End If

''10
If intCol > 44 Then
    Worksheets("Pivot").Range("AD4:AF4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("AD4").Value = Month10
End If

''11
If intCol > 47 Then
    Worksheets("Pivot").Range("AG4:AI4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("AG4").Value = Month11
End If

''12
If intCol > 50 Then
    Worksheets("Pivot").Range("AJ4:AL4").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .MergeCells = True
            .Interior.Color = RGB(184, 204, 228)
            .NumberFormat = "yy/mmm"
        End With
    Range("AJ4").Value = Month12
End If

Application.ScreenUpdating = True
 End Sub

Sub CreatePivot()

Application.ScreenUpdating = False

'Check if Report tab is already formatted.
Dim sh As Worksheet, flg As Boolean
For Each sh In Worksheets
If sh.Name Like "Report*" Then flg = True: Exit For
Next
If flg = True Then
Worksheets("Report").Select
Else
Worksheets("Interface").Select
MsgBox "Report tab does not exist. Copy it and format first!"
End
End If

Range("A1").Select
If ActiveCell.Value = vbNullString Then
    Range("A1").Select
Else
    MsgBox "Report tab is not pivot-ready. We will attempt to format Report tab before creating pivot."
    Call FormatReport
End If


Worksheets("Report").Activate
'Counting Columns
    Dim intCol As Integer
    intCol = Worksheets("Report").Cells(4, Columns.Count).End(xlToLeft).Column
'Counting Rows
    Dim intRow As Integer
    intRow = Worksheets("Report").Cells(Cells.Rows.Count, "A").End(xlUp).Row - 2
'Set PivotRange
    Range(Cells(4, 1), Cells(intRow, intCol)).Name = "PivotRange"
    
    

    
'Check if Pivot tab already exists.
Dim ws As Worksheet
Dim FoundTemp As Boolean
For Each ws In Worksheets
    If ws.Name = "Pivot" Then
        If MsgBox("Overwrite existing Pivot sheet?", vbYesNo, "Pivot Sheet Already Exists") = vbYes Then
            Application.DisplayAlerts = False
            Sheets("Pivot").Delete
            Application.DisplayAlerts = True
        Else
            Worksheets("Interface").Select
            End
        End If
    End If
Next

   

Sheets.Add.Name = "Pivot"
    Sheets("Pivot").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Report!PivotRange", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Pivot!R5C1", TableName:="UDPivot", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets("Pivot").Select
    Cells(5, 1).Select

    With ActiveSheet.PivotTables("UDPivot").PivotFields("Demand Status")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("UDPivot").PivotFields("Demand Status"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("UDPivot").PivotFields("Portfolio")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("UDPivot").PivotFields("Portfolio").ClearAllFilters
    ActiveSheet.PivotTables("UDPivot").PivotFields("Portfolio").CurrentPage = _
        "ITO"
    With ActiveSheet.PivotTables("UDPivot").PivotFields("Demand Entity Type")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("UDPivot").PivotFields("Demand Entity Type"). _
        ClearAllFilters
    ActiveSheet.PivotTables("UDPivot").PivotFields("Demand Entity Type"). _
        CurrentPage = "Delivery Project"
    With ActiveSheet.PivotTables("UDPivot").PivotFields("Demand Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("UDPivot").PivotFields("Position Label")
        .Orientation = xlRowField
        .Position = 2
    End With

''1
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE"), _
        "Sum of Position Forecast FTE", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE").Caption = "Sum of Demand"
        '=======
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE"), _
        "Sum of Allocated Resource Committed FTE", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE").Caption = "Sum of Supply"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE"), "Sum of Unmet Demand FTE", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE"). _
        Caption = "Sum of Unmet Demand"
        '=======
    
''2
If intCol > 20 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE2"), _
        "Sum of Position Forecast FTE2", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE2").Caption = "Sum of Demand2"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE2"), _
        "Sum of Allocated Resource Committed FTE2", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE2").Caption = "Sum of Supply2"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE2"), "Sum of Unmet Demand FTE2", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE2"). _
        Caption = "Sum of Unmet Demand2"
        '=======
End If

    
    
''3
If intCol > 23 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE3"), _
        "Sum of Position Forecast FTE3", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE3").Caption = "Sum of Demand3"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE3"), _
        "Sum of Allocated Resource Committed FTE3", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE3").Caption = "Sum of Supply3"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE3"), "Sum of Unmet Demand FTE3", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE3"). _
        Caption = "Sum of Unmet Demand3"
        '=======
End If

    
''4
If intCol > 26 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE4"), _
        "Sum of Position Forecast FTE4", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE4").Caption = "Sum of Demand4"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE4"), _
        "Sum of Allocated Resource Committed FTE4", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE4").Caption = "Sum of Supply4"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE4"), "Sum of Unmet Demand FTE4", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE4"). _
        Caption = "Sum of Unmet Demand4"
        '=======
End If


''5
If intCol > 29 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE5"), _
        "Sum of Position Forecast FTE5", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE5").Caption = "Sum of Demand5"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE5"), _
        "Sum of Allocated Resource Committed FTE5", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE5").Caption = "Sum of Supply5"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE5"), "Sum of Unmet Demand FTE5", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE5"). _
        Caption = "Sum of Unmet Demand5"
        '=======
End If


''6
If intCol > 32 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE6"), _
        "Sum of Position Forecast FTE6", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE6").Caption = "Sum of Demand6"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE6"), _
        "Sum of Allocated Resource Committed FTE6", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE6").Caption = "Sum of Supply6"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE6"), "Sum of Unmet Demand FTE6", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE6"). _
        Caption = "Sum of Unmet Demand6"
        '=======
End If


''7
If intCol > 35 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE7"), _
        "Sum of Position Forecast FTE7", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE7").Caption = "Sum of Demand7"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE7"), _
        "Sum of Allocated Resource Committed FTE7", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE7").Caption = "Sum of Supply7"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE7"), "Sum of Unmet Demand FTE7", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE7"). _
        Caption = "Sum of Unmet Demand7"
        '=======
End If

        
''8
If intCol > 38 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE8"), _
        "Sum of Position Forecast FTE8", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE8").Caption = "Sum of Demand8"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE8"), _
        "Sum of Allocated Resource Committed FTE8", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE8").Caption = "Sum of Supply8"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE8"), "Sum of Unmet Demand FTE8", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE8"). _
        Caption = "Sum of Unmet Demand8"
        '=======
 End If
 
 
 ''9
If intCol > 41 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE9"), _
        "Sum of Position Forecast FTE9", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE9").Caption = "Sum of Demand9"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE9"), _
        "Sum of Allocated Resource Committed FTE9", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE9").Caption = "Sum of Supply9"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE9"), "Sum of Unmet Demand FTE9", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE9"). _
        Caption = "Sum of Unmet Demand9"
        '=======
 End If
 
 
 ''10
If intCol > 44 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE10"), _
        "Sum of Position Forecast FTE10", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE10").Caption = "Sum of Demand10"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE10"), _
        "Sum of Allocated Resource Committed FTE10", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE10").Caption = "Sum of Supply10"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE10"), "Sum of Unmet Demand FTE10", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE10"). _
        Caption = "Sum of Unmet Demand10"
        '=======
 End If
 
 
  ''11
If intCol > 47 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE11"), _
        "Sum of Position Forecast FTE11", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE11").Caption = "Sum of Demand11"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE11"), _
        "Sum of Allocated Resource Committed FTE11", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE11").Caption = "Sum of Supply11"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE11"), "Sum of Unmet Demand FTE11", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE11"). _
        Caption = "Sum of Unmet Demand11"
        '=======
 End If
 
 
   ''12
If intCol > 50 Then
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Position Forecast FTE12"), _
        "Sum of Position Forecast FTE12", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Position Forecast FTE12").Caption = "Sum of Demand12"
        '=======
     ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Allocated Resource Committed FTE12"), _
        "Sum of Allocated Resource Committed FTE12", xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields( _
        "Sum of Allocated Resource Committed FTE12").Caption = "Sum of Supply12"
        '========
    ActiveSheet.PivotTables("UDPivot").AddDataField ActiveSheet.PivotTables( _
        "UDPivot").PivotFields("Unmet Demand FTE12"), "Sum of Unmet Demand FTE12", _
        xlSum
    ActiveSheet.PivotTables("UDPivot").PivotFields("Sum of Unmet Demand FTE12"). _
        Caption = "Sum of Unmet Demand12"
        '=======
 End If
        
        
    
    Range("A9").Select
    
     With ActiveSheet.PivotTables("UDPivot")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Range("A8").Select
    ActiveSheet.PivotTables("UDPivot").PivotFields("Demand Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Columns("A:A").ColumnWidth = 73
    Columns("B:B").ColumnWidth = 64
    
 'Zoom window to 80%
    ActiveWindow.Zoom = 80
     
 'Set Formats for columns C to AL
    Columns("C:AL").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .ColumnWidth = 10
        .NumberFormat = "0.00"
    End With
    
'Call Sub for Conditional Formats
    Call ColorPivot
    
'Disable Autoformat for the Pivot
    With ActiveSheet.PivotTables("UDPivot")
        .HasAutoFormat = False
    End With
    
    Call PivotHeaders
    
 'Enable wrap-text for row 6 and column A
    Rows("6:6").Select
    Selection.WrapText = True
    Columns("A:A").Select
    Selection.WrapText = True

 'Hide Row 5
    Rows("5:5").Select
    Selection.EntireRow.Hidden = True
    
    Range("A1").Select
    
Application.ScreenUpdating = True

End Sub


Sub ColorPivot()

Application.ScreenUpdating = False

Worksheets("Report").Activate
'Counting Columns
    Dim intCol As Integer
    intCol = Worksheets("Report").Cells(4, Columns.Count).End(xlToLeft).Column

Worksheets("Pivot").Activate
Set Myrange = Range("C7:AL1000")
For Each Cell In Myrange
If Cell.Value = 0 Then
Cell.Font.ColorIndex = 3
End If
Next

''1
Set Myrange = Range("E7:E1000")
    For Each Cell In Myrange
        If Cell.Value > 0 Then
            Cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next

''2
If intCol > 20 Then
    Set Myrange = Range("H7:H1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''3
If intCol > 23 Then
    Set Myrange = Range("K7:K1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''4
If intCol > 26 Then
    Set Myrange = Range("N7:N1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''5
If intCol > 29 Then
    Set Myrange = Range("Q7:Q1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''6
If intCol > 32 Then
    Set Myrange = Range("T7:T1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''7
If intCol > 35 Then
    Set Myrange = Range("W7:W1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''8
If intCol > 38 Then
    Set Myrange = Range("Z7:Z1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''9
If intCol > 41 Then
    Set Myrange = Range("AC7:AC1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''10
If intCol > 44 Then
    Set Myrange = Range("AF7:AF1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''11
If intCol > 47 Then
    Set Myrange = Range("AI7:AI1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

''12
If intCol > 50 Then
    Set Myrange = Range("AL7:AL1000")
        For Each Cell In Myrange
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(255, 0, 0)
            End If
        Next
End If

Application.ScreenUpdating = True

End Sub









