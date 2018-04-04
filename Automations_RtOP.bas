Attribute VB_Name = "Module1"
Sub ProcessRawData()

'Janina Umpa - ES ITO GSM APJ SLIM
''09 October 2014

Dim row_ws As Integer

Application.ScreenUpdating = False

Call FilterData

'Count of rows and clear additional fields from previous run
    Worksheets("RtOP Raw Data").Activate
    row_ws = ActiveSheet.Range("A1").CurrentRegion.Rows.Count 'last row in worksheet
    
    Columns("BB:BL").Select
    Selection.ClearContents
    
 Application.Calculation = xlManual

'Recalculate for additional fields

    Range("BB1").Value = "Month"
    Range("BC1").Value = "Subregion"
    Range("BD1").Value = "Country"
    Range("BE1").Value = "Capability"
    Range("BF1").Value = "RCA Days"
    Range("BG1").Value = "RCA Count"
    Range("BH1").Value = "RCA % Met"
    Range("BI1").Value = "Avoidable"
    Range("BJ1").Value = "Avoidable Caused by Change"
    Range("BK1").Value = "TTR Target"
    Range("BL1").Value = "BU"
    Range("BM1").Value = "Fiscal Year"
    Range("BN1").Value = "MTTR Target"
    Range("BO1").Value = "Outlier?"
    Range("BP1").Value = "P1 RtOP"
    Range("BQ1").Value = "Cause Capability"
    Range("BR1").Value = "Lead Capability2"
    Range("BS1").Value = "RCA Count_Complete Date"
    

    Range("BB2").Value = "=TEXT(V2,""MMM"")"
    Range("BC2").Value = "=F2"
    'Range("BC2").Value = "=VLOOKUP($I2,Lists!$B:$D,2,0)"
    Range("BD2").Value = "=IFERROR(VLOOKUP($I2,Lists!$B:$D,3,0),"""")"
    Range("BF2").Value = "=IF(AN2="""","""",IF(BD2=""Australia"",NETWORKDAYS(AM2,AN2,Lists!$O$4:$O$100)-1,NETWORKDAYS(AM2,AN2)-1))"
    Range("BG2").Value = "=IF(AL2=""Closed"",1,0)"
    Range("BH2").Value = "=IF(BG2=0,0,IF(BF2<=5,1,0))"
    Range("BI2").Value = "=IF(ISERROR(FIND(""YES"",AW2,1)),0,1)"
    Range("BJ2").Value = "=IF(ISERROR(FIND(""YES - CHANGE"",AW2,1)),0,1)"
    Range("BK2").Value = 240
    Range("BL2").Value = "=IF(S2=""APPS Org."",""APPS"",""ITO"")"
    Range("BM2").Value = "=""FY ""&IF(V2<41579,13,IF(V2<41944,14,if(V2<42309,15,16)))"
    Range("BN2").Value = "=IF(BM2=""FY 14"",240,210)"
    Range("BO2").Value = "=IF(AO2>BN2,1,0)"
    'Range("BP2").Value = "=IF(AND(C2=""P1"",G2=""RtOP""),1,0)"
     Range("BP2").Value = "=IF(AND(C2=""P1"",OR(G2=""RtOP"",G2=""ALPHA"")),1,0)"
    Range("BQ2").Value = "=VLOOKUP(AR2,Lists!H:I,2,0)"
    Range("BR2").Value = "=VLOOKUP(S2,Lists!H:I,2,0)"
    Range("BS2").Value = "=IF(AN2="""",0,1)"
    
    Range("BB2:BS2").Copy
    Range("BB3:BS" & row_ws).Select
    ActiveSheet.Paste
    ActiveSheet.Calculate
    Range("BB2:BS" & row_ws).Copy
    Range("BB2:BS" & row_ws).PasteSpecial xlPasteValues

Worksheets("Dashboard").Select
Range("C5").Select

ThisWorkbook.RefreshAll

'Worksheets("Pivot").Select
'Range("A1").Select

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True

End Sub

Sub FilterData()
'This script will delete OOS RtOP data

Worksheets("RtOP Raw Data").Activate
Worksheets("RtOP Raw Data").AutoFilterMode = False

    
'This will delete iRtOP and vRtOP
    ActiveSheet.Range("A1:BA10000").AutoFilter Field:=7, _
        Criteria1:="=iRtOP", Operator:=xlOr, Criteria2:="=vRtOP"
    Worksheets("RtOP Raw Data").Range("2:10000").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("A1:BA10000").AutoFilter Field:=7
    
'This will delete non-P1 RtOPs
    ActiveSheet.Range("A1:BA10000").AutoFilter Field:=3, _
        Criteria1:="<>P1", Operator:=xlAnd
    Worksheets("RtOP Raw Data").Range("2:10000").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("A1:BA10000").AutoFilter Field:=3
    
'This will delete Unclear Subregion
    ActiveSheet.Range("A1:BA10000").AutoFilter Field:=5, _
        Criteria1:="=Unclear", Operator:=xlAnd
    Worksheets("RtOP Raw Data").Range("2:10000").Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("A1:BA10000").AutoFilter Field:=5
    
End Sub

