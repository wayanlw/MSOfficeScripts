'################################################################################
'########### Master function used to calculate ms time           ################
'########### Used by many other methods                          ################
'################################################################################


Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
        "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
        "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Function MicroTimer() As Double
    '
    ' Returns seconds.
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0
    
    ' Get frequency.
    If cyFrequency = 0 Then getFrequency cyFrequency
    
    ' Get ticks.
    getTickCount cyTicks1
    
    ' Seconds
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function


'################################################################################
'######## Calculates the time to calculate Range,sheet,workbook  ################
'######## Uses Microtimer function                               ################
'################################################################################

Sub DoCalcTimer()
    Dim dtime  As Double
    Dim dOvhd  As Double
    Dim oRng   As Range
    Dim oCell  As Range
    Dim oArrRange As Range
    Dim sCalcType As String
    Dim lCalcSave As Long
    Dim bIterSave As Boolean
    Dim jmethod As Integer

    On Error GoTo Errhandl
    
    jmethod = Application.InputBox("What do you want to time?" & _
              "(Must be 1-6)" & vbCr & vbCr & _
              "   [1] Range " & vbCr & _
              "   [2] Sheet Recalculate " & vbCr & _
              "   [3] Workbook Recalculate " & vbCr & _
              "   [4] Open Workbooks Recalculate" & vbCr & _
              "   [5] All open workbooks Calculate Full " & vbCr & _
              " ", Type:=1, Title:="Align Shapes")
    
    ' Initialize
    dtime = MicroTimer
    
    ' Save calculation settings.
    lCalcSave = Application.Calculation
    bIterSave = Application.Iteration
    If Application.Calculation <> xlCalculationManual Then
        Application.Calculation = xlCalculationManual
    End If
    Select Case jmethod
        Case 1
            
            ' Switch off iteration.
            
            If Application.Iteration <> False Then
                Application.Iteration = False
            End If
            
            ' Max is used range.
            
            If Selection.count > 1000 Then
                Set oRng = Intersect(Selection, Selection.Parent.UsedRange)
            Else
                Set oRng = Selection
            End If
            
            ' Include array cells outside selection.
            
            For Each oCell In oRng
                If oCell.HasArray Then
                    If oArrRange Is Nothing Then
                        Set oArrRange = oCell.CurrentArray
                    End If
                    If Intersect(oCell, oArrRange) Is Nothing Then
                        Set oArrRange = oCell.CurrentArray
                        Set oRng = Union(oRng, oArrRange)
                    End If
                End If
            Next oCell
            
            sCalcType = "Calculate " & CStr(oRng.count) & _
                        " Cell(s) in Selected Range: "
        Case 2
            sCalcType = "Recalculate Sheet " & ActiveSheet.Name & ": "
        Case 3
            sCalcType = "Recalculate Active Workbook " & ActiveWorkbook.Name & ": "
        Case 4
            sCalcType = "Recalculate open workbooks: "
        Case 5
            sCalcType = "Full Calculate open workbooks: "
    End Select
    
    ' Get start time.
    dtime = MicroTimer
    Select Case jmethod
        Case 1
            If Val(Application.Version) >= 12 Then
                oRng.CalculateRowMajorOrder
            Else
                oRng.Calculate
            End If
        Case 2
            ActiveSheet.Calculate
        Case 3
            ActiveWorkbook.Sheets.Select
            ActiveSheet.Calculate
        Case 4
            Application.Calculate
        Case 5
            Application.CalculateFull
    End Select
    
    ' Calculate duration.
    dtime = MicroTimer - dtime
    On Error GoTo 0
    
    dtime = Round(dtime, 5)
    MsgBox sCalcType & " " & CStr(dtime) & " Seconds", _
           vbOKOnly + vbInformation, "CalcTimer"
    
Finish:
    
    ' Restore calculation settings.
    If Application.Calculation <> lCalcSave Then
        Application.Calculation = lCalcSave
    End If
    If Application.Iteration <> bIterSave Then
        Application.Iteration = bIterSave
    End If
    Exit Sub
Errhandl:
    On Error GoTo 0
    MsgBox "Unable to Calculate " & sCalcType, _
           vbOKOnly + vbCritical, "CalcTimer"
    GoTo Finish
End Sub


'################################################################################
'########### Calculate the time to run any code within the block ################
'########### Depends on the Micro Timer                          ################
'################################################################################

Sub MacroTimeCalculator()
    Dim dtime  As Double

    'Get start time.
    dtime = MicroTimer
    
    
    '#########################################
    ' Code in here
    ActiveWorkbook.Calculate
    
    '##############################################
    
    ' Calculate duration.
    dtime = MicroTimer - dtime
    On Error GoTo 0
    dtime = Round(dtime, 5)
    MsgBox " Time to run the macro " & CStr(dtime) & " Seconds", _
           vbOKOnly + vbInformation, "CalcTimer"
End Sub

'################################################################################
'########### Calculate the time to run each sheet and reports    ################
'########### Depends on the Micro Timer & sheetExist             ################
'################################################################################


Sub calculateAllsheetsnReport()
    Dim ws     As Worksheet
    Dim x      As Integer
    Dim wsName As String
    Dim ResultSheetName As String
    Dim lCalcSave As Long
    Dim bIterSave As Boolean

    ResultSheetName = "SheetCalcSummary"
    
    
    x = 3
    
    If Sht_Fnc_SheetExists(ResultSheetName) = True Then
        Worksheets(ResultSheetName).Range("B:D").Clear
    Else
        Sheets.Add(before:=Worksheets(1)).Name = ResultSheetName
    End If
    
    Worksheets(ResultSheetName).Range("B2").Value = "List of Worksheets"
    Worksheets(ResultSheetName).Range("C2").Value = "Calculation Time"
    Worksheets(ResultSheetName).Range("D2").Value = "UsedRange"
    
    With Worksheets(ResultSheetName).Range("B2:D2")
        .Font.Bold = True
        .Font.Underline = True
    End With
    
    lCalcSave = Application.Calculation
    bIterSave = Application.Iteration
    
    If Application.Calculation <> xlCalculationManual Then
        Application.Calculation = xlCalculationManual
    End If
    
    If Application.Iteration <> False Then
        Application.Iteration = False
    End If
    
    For Each ws In Worksheets
        If ws.Name <> ResultSheetName Then
            
            If ws.Visible = True Then
                wsName = ws.Name
            Else
                wsName = ws.Name & " (Hidden)"
            End If
            
            dtime = MicroTimer
            ws.Calculate
            ' Calculate duration.
            dtime = MicroTimer - dtime
            On Error GoTo 0
            
            dtime = Round(dtime, 5)
            Sheets(ResultSheetName).Hyperlinks.Add _
                                                   Anchor:=Sheets(ResultSheetName).Cells(x, 2), Address:="", SubAddress:= _
                                                   "'" & ws.Name & "'!A1", TextToDisplay:=wsName
            Sheets(ResultSheetName).Cells(x, 3).Value = dtime
            Sheets(ResultSheetName).Cells(x, 4).Value = ws.UsedRange.Address
            x = x + 1
            
        End If
    Next ws
    
    Worksheets(ResultSheetName).Activate
    Worksheets(ResultSheetName).UsedRange.EntireColumn.AutoFit
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    Application.GoTo Reference:=Range("a1"), Scroll:=True
    Application.Calculation = xlCalculationAutomatic
    Application.Iteration = True
    
End Sub


'################################################################################
'######## Checks whether a sheet already exists in the workbook  ################
'######## Used by other sheet methods                            ################
'################################################################################

Function Sht_Fnc_SheetExists(sheetname As String) As Boolean
    'PURPOSE: Determine is a worksheet exists in the ActiveWorkbook
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim sht    As Worksheet
    'Test if sheet can be found
    On Error Resume Next
    Set sht = ActiveWorkbook.Worksheets(sheetname)
    On Error GoTo 0
    'Determine function result
    If Not sht Is Nothing Then Sht_Fnc_SheetExists = True
    'Clear Memory
    Set sht = Nothing
End Function
