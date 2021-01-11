

' ---------------------------------- Full clean all sheets
' Crreutnly there is sull clean all sheets with options where user selects whether all sheets or current sheet.

Sub sht_FullCleanAllsheets()

    Dim ws     As Worksheet
    Dim cursheet As Worksheet
    Dim curcell As Range

    Set cursheet = ActiveSheet
    Set curcell = ActiveCell
    
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        'Remove gridlines, set zoom to 70%, go to A1
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 70
        Application.GoTo Reference:=Range("a1"), Scroll:=True
        
        With ws
            
            'unhide all sheets and expand outlines
            .Visible = xlSheetVisible
            .Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
            With .UsedRange
                'autofit rows and columns
                .EntireColumn.AutoFit
                .EntireRow.AutoFit
                'unhide rows and columns
                .Columns.EntireColumn.Hidden = False
                .Rows.EntireRow.Hidden = False
                'change fonts
                With .Font
                    .Name = "Calibri"
                    .FontStyle = "Regular"
                    .Size = 11
                End With
            End With
        End With
        
    Next ws
    cursheet.Activate
    curcell.Activate
    
End Sub



'################################################################################################################################'
' Finding the celll fill color and text color 
'################################################################################################################################'


Sub GetRGBColor_Font()
'PURPOSE: Output the RGB color code for the ActiveCell's Font Color
'SOURCE: www.TheSpreadsheetGuru.com

Dim HEXcolor As String
Dim RGBcolor As String

HEXcolor = Right("000000" & Hex(ActiveCell.Font.Color), 6)

RGBcolor = "RGB (" & CInt("&H" & Right(HEXcolor, 2)) & _
", " & CInt("&H" & Mid(HEXcolor, 3, 2)) & _
", " & CInt("&H" & Left(HEXcolor, 2)) & ")"

MsgBox RGBcolor, vbInformation, "Cell " & ActiveCell.Address(False, False) & ":  Font Color"

End Sub

Sub GetRGBColor_Fill()
'PURPOSE: Output the RGB color code for the ActiveCell's Fill Color
'SOURCE: www.TheSpreadsheetGuru.com

Dim HEXcolor As String
Dim RGBcolor As String

HEXcolor = Right("000000" & Hex(ActiveCell.Interior.Color), 6)

RGBcolor = "RGB (" & CInt("&H" & Right(HEXcolor, 2)) & _
", " & CInt("&H" & Mid(HEXcolor, 3, 2)) & _
", " & CInt("&H" & Left(HEXcolor, 2)) & ")"

MsgBox RGBcolor, vbInformation, "Cell " & ActiveCell.Address(False, False) & ":  Fill Color"

End Sub


'################################################################################################################################'
' Make all charts in workbook plot non-visible cells
'################################################################################################################################'



Sub PlotNonVisibleCells()
'PURPOSE: Make all charts in workbook plot non-visible cells
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    
Dim sht As Worksheet
Dim cht As ChartObject
Dim CurrentSheet As Worksheet

Application.ScreenUpdating = False
Application.EnableEvents = False

Set CurrentSheet = ActiveSheet

'Loop Through All Worksheets in Workbook
  For Each sht In ActiveWorkbook.Worksheets
    'Loop Through all Charts in Worksheet
      For Each cht In sht.ChartObjects
        cht.Activate
        ActiveChart.PlotVisibleOnly = False
      Next cht
  Next sht

CurrentSheet.Activate
Application.EnableEvents = True

'Completed
  MsgBox "All charts will now plot non-visible cells!", , "Macro Complete!"
  
End Sub




'################################################################################################################################
' Timer for worksheet, workbook etc etc. 
'################################################################################################################################'


#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
        "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
         "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias _                                            "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias _
        "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If
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

Sub RangeTimer()
    DoCalcTimer 1
End Sub
Sub SheetTimer()
    DoCalcTimer 2
End Sub
Sub RecalcTimer()
    DoCalcTimer 3
End Sub
Sub FullcalcTimer()
    DoCalcTimer 4
End Sub

Sub DoCalcTimer(jMethod As Long)
    Dim dTime As Double
    Dim dOvhd As Double
    Dim oRng As Range
    Dim oCell As Range
    Dim oArrRange As Range
    Dim sCalcType As String
    Dim lCalcSave As Long
    Dim bIterSave As Boolean
    '
    On Error GoTo Errhandl

' Initialize
    dTime = MicroTimer

    ' Save calculation settings.
    lCalcSave = Application.Calculation
    bIterSave = Application.Iteration
    If Application.Calculation <> xlCalculationManual Then
        Application.Calculation = xlCalculationManual
    End If
    Select Case jMethod
    Case 1

        ' Switch off iteration.

        If Application.Iteration <> False Then
            Application.Iteration = False
        End If
        
        ' Max is used range.

        If Selection.Count > 1000 Then
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

        sCalcType = "Calculate " & CStr(oRng.Count) & _
            " Cell(s) in Selected Range: "
    Case 2
        sCalcType = "Recalculate Sheet " & ActiveSheet.Name & ": "
    Case 3
        sCalcType = "Recalculate open workbooks: "
    Case 4
        sCalcType = "Full Calculate open workbooks: "
    End Select

' Get start time.
    dTime = MicroTimer
    Select Case jMethod
    Case 1
        If Val(Application.Version) >= 12 Then
            oRng.CalculateRowMajorOrder
        Else
            oRng.Calculate
        End If
    Case 2
        ActiveSheet.Calculate
    Case 3
        Application.Calculate
    Case 4
        Application.CalculateFull
    End Select

' Calculate duration.
    dTime = MicroTimer - dTime
    On Error GoTo 0

    dTime = Round(dTime, 5)
    MsgBox sCalcType & " " & CStr(dTime) & " Seconds", _
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







'################################################################################################################################
' Time taken to run the macro 
'################################################################################################################################'



Sub CalculateRunTime_Seconds()
'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer

'*****************************
'Insert Your Code Here...
'*****************************

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub







'################################################################################################################################
' Align Shapes 
'################################################################################################################################'

Sub Shape_AlignMultipleShapes()
    'PURPOSE: Align each shape in user's selection (first shape selected stays put)
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim Shp1   As Shape
    Dim Shp2   As Shape
    Dim x      As Integer
    Dim y      As Integer
    Dim align  As Integer

    align = Application.InputBox("how do you want to align" & _
            "(Must be 1-6)" & vbCr & vbCr & _
            "   [1] Left " & vbCr & _
            "   [2] Right " & vbCr & _
            "   [3] Top " & vbCr & _
            "   [4] Bottom " & vbCr & _
            "   [5] Middle " & vbCr & _
            "   [6] Centre" & vbCr & _
            " ", Type:=1, Title:="Align Shapes")
    
    If align > 6 Or align < 1 Then
        MsgBox "Wrong Input.. "
    End If
    
    
    
    'Count How Many Shapes Are Selected
    x = Windows(1).Selection.ShapeRange.count
    
    'Loop Through each selected Shape (align with first selected)
    For y = 1 To x
        If Shp1 Is Nothing Then
            Set Shp1 = Windows(1).Selection.ShapeRange(y)
        Else
            Set Shp2 = Windows(1).Selection.ShapeRange(y)
            
            Select Case align
                Case 1
                    'align Left
                    Shp2.Left = Shp1.Left
                Case 2
                    'Align Right
                    Shp2.Left = Shp1.Left + (Shp1.Width - Shp2.Width)
                Case 3
                    'Align Top
                    Shp2.Top = Shp1.Top
                Case 4
                    'Align Bottom
                    Shp2.Top = Shp1.Top + (Shp1.Height - Shp2.Height)
                Case 5
                    'Align Middle (Horizontal Center)
                    Shp2.Top = Shp1.Top + ((Shp1.Height - Shp2.Height) / 2)
                Case 6
                    'Align Center (Vertical Center)
                    Shp2.Left = Shp1.Left + ((Shp1.Width - Shp2.Width) / 2)
            End Select
            
        End If
    Next y
    
End Sub




'################################################################################################################################
' Looping within an inputbox until the correct input is received
'################################################################################################################################'





Sub LoopInputBox()
'PURPOSE: An example of how to keep looping an input box until a valid answer is entered
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim InputQuestion As String
Dim myAnswer As Variant

'Question to pose to the user
  InputQuestion = "Please enter your age." & vbNewLine _
    & "(Don't even think about lying!)"

'Keeping looping until you get a valid answer to your question
  Do
    'Retrieve an answer from the user
      myAnswer = Application.InputBox(InputQuestion, "Your Age?", Type:=1)
    
    'Check if user selected cancel button
      If TypeName(myAnswer) = "Boolean" Then Exit Sub
      
  Loop While myAnswer <= 0 Or myAnswer > 120

'We've got a valid answer from the user, let's continue on...

End Sub