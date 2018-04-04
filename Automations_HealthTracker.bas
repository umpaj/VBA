Attribute VB_Name = "Module1"
Sub Show_Months()
Attribute Show_Months.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Show_Months Macro
'

    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    'Worksheets("Top Quadrant All Cap").Range("FA1").Value = "CONFIGURATION MANAGEMENT"
    'Worksheets("Top Quadrant All Cap").Range("FO1").Value = "AVAILABILITY MANAGEMENT"
    'Worksheets("Top Quadrant All Cap").Range("FV1").Value = "CAPACITY MANAGEMENT"
    
End Sub
Sub Hide_Months()
Attribute Hide_Months.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Hide_Months Macro
'

    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    'Worksheets("Top Quadrant All Cap").Range("FA1").Value = "CFGM"
    'Worksheets("Top Quadrant All Cap").Range("FO1").Value = "AVLM"
    'Worksheets("Top Quadrant All Cap").Range("FV1").Value = "CAPM"
    
End Sub


Sub Main_Page()
'
'
'

    Worksheets("Top 20 All Cap").Select
End Sub

Sub VOC_Status()

'Blue is 11
'Green is 43
'Yellow is 44
'Red is 3
Dim i As Integer

Worksheets("ASM Summary").Activate
i = Range("U6")

'VoC - Loyalty
If Worksheets("Top 20 All Cap").Range("K" & i).Interior.ColorIndex = 43 Then
    Range("K9").Value = 0
ElseIf Worksheets("Top 20 All Cap").Range("K" & i).Interior.ColorIndex = 44 Then
    Range("K9").Value = 1
ElseIf Worksheets("Top 20 All Cap").Range("K" & i).Interior.ColorIndex = 3 Then
     Range("K9").Value = 2
Else: Range("K9").Value = 0
End If

'VoC - Quality
If Worksheets("Top 20 All Cap").Range("L" & i).Interior.ColorIndex = 43 Then
    Range("M9").Value = 0
ElseIf Worksheets("Top 20 All Cap").Range("L" & i).Interior.ColorIndex = 44 Then
    Range("M9").Value = 1
ElseIf Worksheets("Top 20 All Cap").Range("L" & i).Interior.ColorIndex = 3 Then
     Range("M9").Value = 2
Else: Range("M9").Value = 0
End If

'Internal Rating
If Worksheets("Top 20 All Cap").Range("N" & i).Interior.ColorIndex = 43 Then
    Range("Q9").Value = 0
ElseIf Worksheets("Top 20 All Cap").Range("N" & i).Interior.ColorIndex = 44 Then
    Range("Q9").Value = 1
ElseIf Worksheets("Top 20 All Cap").Range("N" & i).Interior.ColorIndex = 3 Then
     Range("Q9").Value = 2
Else: Range("Q9").Value = 0
End If
End Sub




Sub ASM_Static()
'
' Janina Umpa - janina.umpa@hp.com - 25 April 2012

    'Call VOC_Status
    Sheets("ASM Summary").Copy After:=Sheets(Sheets.Count)
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L2:O2").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
      
    Dim Account As String, Month As String
    Account = Range("L2")
    Month = Range("U3")
    
    ActiveSheet.Name = Account & "_" & Month
    
    Range("R6").Select
    
    Call EmailWithOutlook
    
End Sub

Sub EmailWithOutlook()
     'Variable declaration
    Dim oApp As Object, _
    oMail As Object, _
    WB As Workbook, _
    FileName As String, PreFileName As String, EmailSubject As String
   
   EmailSubject = Range("L2") & " ASM Summary: " & Range("U3")
   PreFileName = Range("L2") & "_" & Range("U3") & ".xls"
   
     'Turn off screen updating
    Application.ScreenUpdating = False
     
     'Make a copy of the active sheet and save it to
     'a temporary file
    ActiveSheet.Copy
    Set WB = ActiveWorkbook
    FileName = PreFileName
    On Error Resume Next
    Kill "C:\" & FileName
    On Error GoTo 0
    WB.SaveAs FileName:="C:\" & FileName
     
     'Create and show the outlook mail item
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    With oMail
         'Uncomment the line below to hard code a recipient
         .To = "manilaslim@hp.com"
         'Uncomment the line below to hard code a subject
         .Subject = EmailSubject
        .Attachments.Add WB.FullName
        .Display
    End With
     
     'Delete the temporary file
    WB.ChangeFileAccess Mode:=xlReadOnly
    Kill WB.FullName
    WB.Close SaveChanges:=False
     
     'Restore screen updating and release Outlook
    Application.ScreenUpdating = True
    Set oMail = Nothing
    Set oApp = Nothing
End Sub

Sub Subregion_Static()
'
' Janina Umpa - janina.umpa@hp.com - 04 July 2012

    'Call VOC_Status
    Application.DisplayAlerts = False
    Sheets("Sub Region").Copy After:=Sheets(Sheets.Count)
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L2:O2").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
      
    Dim Account As String, Month As String
    Account = Range("L2")
    Month = Range("U3")
    
    ActiveSheet.Name = Account & "_" & Month
    
    Range("R6").Select
    
    Call EmailWithOutlook_SB
    
End Sub

Sub EmailWithOutlook_SB()
     'Variable declaration
    Dim oApp As Object, _
    oMail As Object, _
    WB As Workbook, _
    FileName As String, PreFileName As String, EmailSubject As String
   
   EmailSubject = Range("L2") & " Sub-Region Summary: " & Range("U3")
   PreFileName = Range("L2") & "_" & Range("U3") & ".xls"
   
     'Turn off screen updating
    Application.ScreenUpdating = False
     
     'Make a copy of the active sheet and save it to
     'a temporary file
    ActiveSheet.Copy
    Set WB = ActiveWorkbook
    FileName = PreFileName
    On Error Resume Next
    Kill "C:\" & FileName
    On Error GoTo 0
    WB.SaveAs FileName:="C:\" & FileName
     
     'Create and show the outlook mail item
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    With oMail
         'Uncomment the line below to hard code a recipient
         .To = "manilaslim@hp.com"
         'Uncomment the line below to hard code a subject
         .Subject = EmailSubject
        .Attachments.Add WB.FullName
        .Display
    End With
     
     'Delete the temporary file
    WB.ChangeFileAccess Mode:=xlReadOnly
    Kill WB.FullName
    WB.Close SaveChanges:=False
     
     'Restore screen updating and release Outlook
    Application.ScreenUpdating = True
    Set oMail = Nothing
    Set oApp = Nothing
Application.DisplayAlerts = True
End Sub

