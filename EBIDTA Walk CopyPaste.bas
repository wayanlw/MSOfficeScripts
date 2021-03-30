Option Explicit

Sub TPH_PnLValuePaste()
    Dim phwb1  As Workbook, phwb2 As Workbook
    Dim Ret1, Ret2

    '    Application.ScreenUpdating = False
    ActiveSheet.Range("R1").Copy
    Set phwb1 = ActiveWorkbook
    Application.DisplayAlerts = False


    '~~> Get the File
    Ret1 = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", _
    , "Please select file")
    If Ret1 = False Then Exit Sub

    'validate the file name
    If InStr(1, Ret1, "Loss.xls") = 0 Then
        MsgBox ("Wrong Selection. Please select the TPH P&L file")
        Exit Sub
    End If

    Set phwb2 = Workbooks.Open(Ret1)

'    phwb1.Worksheets("TPH EBITDA Rec").Activate

    'Clear the contents
    phwb1.Worksheets("TPH EBITDA Rec").Range("B2:N38").ClearContents
    phwb1.Worksheets("TPH vs LY Mth Rec").Range("B2:N38").ClearContents
    phwb1.Worksheets("TPH YTD vs BUD Rec").Range("B2:N38").ClearContents
    phwb1.Worksheets("PH vs LYTD Rec").Range("B2:N38").ClearContents



    'TPH Month. -- Change source sheet and range in line1. Destination sheet and range in line2
    'Forecast Range "F230:R266"| Budget range "B78:N114"
    phwb2.Worksheets("PH EBITDA Rec").Range("B78:N114").Copy
    phwb1.Worksheets("TPH EBITDA Rec").Range("B2").PasteSpecial Paste:=xlPasteValues


    'TPH vs LY Mth Rec. -- Change source sheet and range in line1. Destination sheet and range in line2
    phwb2.Worksheets("PH EBITDA Rec").Range("B154:N190").Copy
    phwb1.Worksheets("TPH vs LY Mth Rec").Range("B2").PasteSpecial Paste:=xlPasteValues

    'TPH YTD V Bud . -- Change source sheet and range in line1. Destination sheet and range in line2
    'Forecast Range "F192:R228"| Budget range "B2:N38"
    phwb2.Worksheets("PH EBITDA Rec").Range("B2:N38").Copy
    phwb1.Worksheets("TPH YTD vs BUD Rec").Range("B2").PasteSpecial Paste:=xlPasteValues


    'TPH LYTD
    phwb2.Worksheets("PH EBITDA Rec").Range("B40:N76").Copy
    phwb1.Worksheets("PH vs LYTD Rec").Range("B2").PasteSpecial Paste:=xlPasteValues

    phwb2.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    '
    '
End Sub

Sub HC_PnLValuePaste()
    Dim hcwb1    As Workbook, hcwb2 As Workbook
    Dim Ret1, Ret2


    Application.ScreenUpdating = False

    ActiveSheet.Range("R1").Copy
    Set hcwb1 = ActiveWorkbook
    Application.DisplayAlerts = False


    '~~> Get the File
    Ret1 = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", _
    , "Please select file")
    If Ret1 = False Then Exit Sub

    'validate the file name
    If InStr(Ret1, "HC") = 0 Then
        MsgBox ("Wrong Selection. Please select the HC file")
        Exit Sub
    End If


    Set hcwb2 = Workbooks.Open(Ret1)

'    hcwb1.Worksheets("HC EBITDA Rec").Activate

    'Clear the contents
    hcwb1.Worksheets("HC EBITDA Rec").Range("B2:N38").ClearContents
    hcwb1.Worksheets("HC vs BUD YTD rec").Range("B2:N38").ClearContents
    hcwb1.Worksheets("HC vs LY MTH Rec").Range("B2:N38").ClearContents
    hcwb1.Worksheets("HC vs LYTD rec").Range("B2:N38").ClearContents

    'HC Month. -- Change source sheet and range in line1. Destination sheet and range in line2
    'Forecast Range "F202:R238"| Budget range"F122:R158"
    hcwb2.Worksheets("HC EBITDA Rec").Range("F122:R158").Copy
    hcwb1.Worksheets("HC EBITDA Rec").Range("B2").PasteSpecial Paste:=xlPasteValues

    'hcwb1.Worksheets("HC EBITDA Rec").ActivateHC YTD V Bud . -- Change source sheet and range in line1. Destination sheet and range in line2
    'Forecast range "F242:R278"| Budget range "F4:R40"
    hcwb2.Worksheets("HC EBITDA Rec").Range("F4:R40").Copy
    hcwb1.Worksheets("HC vs BUD YTD rec").Range("B2").PasteSpecial Paste:=xlPasteValues

    'HC vs LY Mth Rec. -- Change source sheet and range in line1. Destination sheet and range in line2
    hcwb2.Worksheets("HC EBITDA Rec").Range("F162:R198").Copy
    hcwb1.Worksheets("HC vs LY MTH Rec").Range("B2").PasteSpecial Paste:=xlPasteValues

    'HC LYTD
    hcwb2.Worksheets("HC EBITDA Rec").Range("F82:R118").Copy
    hcwb1.Worksheets("HC vs LYTD rec").Range("B2").PasteSpecial Paste:=xlPasteValues


    hcwb2.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Set hcwb2 = Nothing
    Set hcwb1 = Nothing

End Sub


