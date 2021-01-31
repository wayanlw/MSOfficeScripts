Attribute VB_Name = "Formatting_Module"
Option Explicit


Sub AutoColorSelection_A()
Attribute AutoColorSelection_A.VB_ProcData.VB_Invoke_Func = "A\n14"
    '
    ' SuMacro Macro
    ' Keyboard Shortcut: Ctrl+Shift+A
    '

    Dim cell   As Range, constantCell As Range, formulaCells As Range
    Dim cellFormula As String

    With Selection
        On Error Resume Next
        Set constantCell = .SpecialCells(xlCellTypeConstants, xlNumbers)
        Set formulaCells = .SpecialCells(xlCellTypeFormulas, 23)
        On Error GoTo 0
    End With
    
    
    If Not constantCell Is Nothing Then
        constantCell.Font.Color = vbBlue
    End If
    
    If Not formulaCells Is Nothing Then
        For Each cell In formulaCells
            cellFormula = cell.Formula
            
            If cellFormula Like "*.xls*]*!*" Then
                cell.Font.Color = RGB(0, 176, 80)
            ElseIf cellFormula Like "*!*" Then
                
                '    And Not cellFormula Like "*\**" _
                '    And Not cellFormula Like "*+*" _
                '    And Not cellFormula Like "*-*" _
                '    And Not cellFormula Like "*/*" _
                '    And Not cellFormula Like "*^*" _
                '    And Not cellFormula Like "*%*" _
                '    And Not cellFormula Like "*>*" _
                '    And Not cellFormula Like "*<*" _
                '    And Not cellFormula Like "*=<*" _
                '    And Not cellFormula Like "*=>*" _
                '    And Not cellFormula Like "*<>*" _
                '    And Not cellFormula Like "*&*" Then
                cell.Font.Color = RGB(100, 0, 100)
            Else
                cell.Font.Color = vbBlack
            End If
        Next cell
    End If
    
End Sub


Sub NumberFormat_N()
Attribute NumberFormat_N.VB_ProcData.VB_Invoke_Func = "N\n14"
    ' Ctrl+Shift+N
    ' number Macro
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
End Sub




Sub DecimalIncrease_J()
Attribute DecimalIncrease_J.VB_ProcData.VB_Invoke_Func = "J\n14"
    ' Ctrol + Shift + J
    Application.CommandBars.FindControl(ID:=398).Execute
    
End Sub


Sub DecimalDecrease_K()
Attribute DecimalDecrease_K.VB_ProcData.VB_Invoke_Func = "K\n14"
    ' Ctrol + Shift + K
    Application.CommandBars.FindControl(ID:=399).Execute
End Sub


Sub CenterAcrossSelection_M()
Attribute CenterAcrossSelection_M.VB_ProcData.VB_Invoke_Func = "M\n14"
    '
    ' Macro3 Macro
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .WrapText = True
        .MergeCells = False
    End With
End Sub

Sub AllSheetsNoGridZoom()
    '
    ' WorksheetGridZoom Macro

    Dim ws     As Worksheet
    Dim curcell As Range
    Dim cursheet As Worksheet

    Set cursheet = ActiveSheet
    
    For Each ws In Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Select
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 70
            Application.GoTo Reference:=Range("a1"), Scroll:=True
            
        End If
        
    Next ws
    
    cursheet.Activate
    
End Sub

Sub CurrentSheetNoGridZoom70_G()
Attribute CurrentSheetNoGridZoom70_G.VB_ProcData.VB_Invoke_Func = "G\n14"
    '
    ' WorksheetGridZoom Macro

    Dim ws     As Worksheet
    Dim curcell As Range


    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 70
    Application.GoTo Reference:=Range("a1"), Scroll:=True
    
End Sub

Sub AllSheetsFontCalibri11()
    Dim ws     As Worksheet
    For Each ws In Worksheets
        With ws
            .Cells.Font.Name = "calibri"
            .Cells.Font.Size = 11
        End With
    Next
End Sub


Sub CycleFill_X()
Attribute CycleFill_X.VB_ProcData.VB_Invoke_Func = "X\n14"
    Dim yellow As Long, grey As Long
    Dim Color3 As Long, Color4 As Long

    yellow = 10092543
    grey = 15395562
    
    'Selection.Interior.Pattern = xlNone
    
    If Selection.Interior.Pattern = xlNone Then
        Selection.Interior.Color = grey
    ElseIf Selection.Interior.Color = grey Then
        Selection.Interior.Color = yellow
    Else
        Selection.Interior.Pattern = xlNone
    End If
    
End Sub

Sub CycleFontColor_C()
Attribute CycleFontColor_C.VB_ProcData.VB_Invoke_Func = " \n14"

    '    Cycle through different font colors
    '

    If Selection.Font.Color = vbBlack Then
        Selection.Font.Color = vbBlue
    ElseIf Selection.Font.Color = vbBlue Then
        Selection.Font.Color = -4144960
    Else
        Selection.Font.Color = vbBlack
        'Selection.Font.Bold = False
    End If
    
End Sub


Sub CycleCellStyle_T()
Attribute CycleCellStyle_T.VB_ProcData.VB_Invoke_Func = "T\n14"
    '
    ' Cycle through different cell styles
    '
    Dim YellowBack As Long
    Dim GreyBack As Long
    Dim BlueBack As Long
    Dim Bluefont As Long

    GreyBack = 15395562
    YellowBack = 10092543
    BlueBack = 8011008
    Bluefont = vbBlue
    
    
    If Selection.Interior.Pattern = xlNone And Selection.Font.Bold = False Then
        '        Activate if necessary. Add another step for grey background and bold font. Removed because can use CycleFill_V()
        '        Selection.Interior.Color = GreyBack
        '        Selection.Font.Bold = True
        '    ElseIf Selection.Interior.Color = GreyBack And Selection.Font.Bold = True Then
        Selection.Interior.Color = YellowBack
        Selection.Font.Color = Bluefont
        Selection.Font.Bold = False
    ElseIf Selection.Interior.Color = YellowBack And Selection.Font.Color = Bluefont And Selection.Font.Bold = False Then
        Selection.Interior.Color = BlueBack
        Selection.Font.Color = vbWhite
        Selection.Font.Bold = True
    Else
        Selection.Interior.Pattern = xlNone
        Selection.Font.Bold = False
        Selection.Font.Color = vbBlack
    End If
    
End Sub

Sub CycleBorders_B()
Attribute CycleBorders_B.VB_ProcData.VB_Invoke_Func = "B\n14"

    If Selection.Borders(xlEdgeBottom).LineStyle = xlNone And Selection.Borders(xlEdgeTop).LineStyle = xlNone Then
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = vbBlack
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
    ElseIf Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous And Selection.Borders(xlEdgeBottom).Weight = xlThin Then
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    ElseIf Selection.Borders(xlEdgeTop).LineStyle = xlContinuous And Selection.Borders(xlEdgeBottom).Weight = xlThin Then
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -8766208
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        '    ElseIf Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous And Selection.Borders(xlEdgeBottom).Weight = xlMedium Then
        '        With Selection.Borders
        '            .LineStyle = xlContinuous
        '            .Color = RGB(220, 220, 220)
        '
        '        End With
        
    Else
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
    End If
    
End Sub

Sub PasteValues_V()
Attribute PasteValues_V.VB_ProcData.VB_Invoke_Func = "V\n14"
    '
    ' PasteValues Macro
    ' Paste only Values from copied cell data.
    '
    ' Keyboard Shortcut: Ctrl+Shift+V
    '
    On Error Resume Next
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:= _
                           False
End Sub


Sub AllGreyBorders_E()
Attribute AllGreyBorders_E.VB_ProcData.VB_Invoke_Func = "E\n14"
    With Selection.Borders
        .LineStyle = xlContinuous
        .Color = RGB(220, 220, 220)
    End With
    
End Sub
