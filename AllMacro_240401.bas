Attribute VB_Name = "AllMacro"

' List of Functions
' SHORTCUTS
'   Alignment - CS R
'   Bo
'
'
'
'
'
'
'
'
'
'
'
''
'
'
'
'
'













' /* --------------------------- Formatting Section --------------------------- */

' Cycle Number Format
' Cycle Cell Fill
' Cycle Font Color
' Cycle Cell Style
' Cycle Border Design
' Decimal Increase  Decrease
' Alignment
' Sheets No Grid and Zoom
' Current Sheet no Grid Zoon
' All sheets No Grid Zoom
' All Sheets Font Type and Font Size
' Auto Color the Selection
' Auto Color All worksheets

' /* --------------------------- Worksheets Section --------------------------- */

' Clean Sheet
' Reset Used Rnage
' Insert sheet with names
' Create Sheets from Selected Name Range
' Funcion to check whether a sheet already exist
' List All sheets
' Hide and Unhide Worksheets

' /* ------------------------------ Tools Section ----------------------------- */

' Find All Select them highlight them
' Highlight all Duplicates in the selected range
' Remove duplicates with error handling
' Clean and trim the selection
' Ask to save workbook before running the macro
' Load brand colors to the recent colors
' Convert the selection to numbers
' Convert the selection to numbers Old Version
' Wrap if error
' List unique values in first two columns
' Remove same sheet reference
' Functions to extract numbers and text
' This VBA code will create a function to get the numeric part from a string
' This VBA code will create a function to get the text part from a string
' Turn Off Pivot Table Autofit Column Width On Update Setting
' Change Pivot table fields to sum and set formating
' Swap two selected ranges
' Backup a copy
' Documenting and commenting
' Remove white fills
' Insert Picture as Comment
' Finding the celll fill color and text color
' Save         as a Text File in the current location
' Shape Section
' Change whehter the shapescharts resize with rowcolumn widths change
' cycle shapescharts resize with rowcolumn widths change
' Align Shapes
' Make all charts in workbook plot nonvisible cells

' References
' https://www.thespreadsheetguru.com/the-code-vault?offset=1432340368034
' https://www.excelcampus.com/vba/remove-duplicates-list-unique-values
' http://www.contextures.on.ca/xlfaqApp.html#Unused

' define brand colors
'Const branddark1 As String = "000,032,091" 'dark Purple essity
'Const branddark2 As String = "055,095,146" 'Dark blue mckenzie
Const branddark1 As String = "000,000,090" 'dark Purple essity
Const branddark2 As String = "245,000,130" 'Dark blue mckenzie
Const brandlight1 As String = "200,220,255" ' light blue essity
Const brandlight2 As String = "255,217,237" ' light pink

'Define the zoom level
Const sheetZoomLevel As Integer = 80



' /* -------------------------------------------------------------------------- */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /*                             Formatting Section                             */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /* -------------------------------------------------------------------------- */


' /* --------------------------- Cycle Number Format -------------------------- */

Sub Fmt_NumberFormat_N()
Attribute Fmt_NumberFormat_N.VB_ProcData.VB_Invoke_Func = "N\n14"
    ' Ctrl+Shift+N
    ' number Macro

    Const numberFormat1 As String = "_- #,##0_-;- #,##0_-;_-* ""-""??_-;_-@_-"
    Const numberFormat2 As String = "_-$ #,##0_-;-$ #,##0_-;_- ""-""??_-;_-@_-"
    Const numberFormat3 As String = "_(* #,##0_);_(* [red](#,##0);_(* """"-""""??_);_(@_)"

    Select Case Selection.NumberFormat
        Case "General"
            Selection.NumberFormat = numberFormat1
            
        Case numberFormat1
            Selection.NumberFormat = numberFormat2
            
        Case numberFormat2
            Selection.NumberFormat = numberFormat3
            
        Case Else
            Selection.NumberFormat = "General"
    End Select
End Sub



' /* ----------------------------- Cycle Cell Fill ---------------------------- */

Sub Fmt_CycleFill_X()
Attribute Fmt_CycleFill_X.VB_ProcData.VB_Invoke_Func = "X\n14"
    Dim color_1, color_2 As Long
    Dim isBrandColor As Boolean
    'Dim Color3 As Long, Color4 As Long

    isBrandColor = True
    
    If isBrandColor Then
        color_1 = 15395562
        color_2 = RGB(Left(brandlight2, 3), Mid(brandlight2, 5, 3), Right(brandlight2, 3))
        color_3 = RGB(Left(brandlight1, 3), Mid(brandlight1, 5, 3), Right(brandlight1, 3))
    Else
        color_1 = 15395562                        ' Light Grey
        color_2 = 15587839                        ' Light pink
        
    End If
    
    If Selection.Interior.Pattern = xlNone Then
        Selection.Interior.Color = color_1
    ElseIf Selection.Interior.Color = color_1 Then
        Selection.Interior.Color = color_2
        'Selection.Font.Bold = True
        Selection.Font.Color = RGB(156, 0, 6)
    ElseIf Selection.Interior.Color = color_2 Then
        Selection.Interior.Color = color_3
        Selection.Font.Color = RGB(32, 55, 100)
        'Selection.Font.Bold = True
    Else
        Selection.Interior.Pattern = xlNone
        Selection.Font.Bold = False
        Selection.Font.Color = vbBlack
    End If
    
    
End Sub

' /* ---------------------------- Cycle Font Color ---------------------------- */

Sub Fmt_CycleFontColor_F()
Attribute Fmt_CycleFontColor_F.VB_ProcData.VB_Invoke_Func = "F\n14"
    '    Cycle through different font colors
    ' Black >> Grey >> White

    If Selection.Font.Color = vbBlack Then
        Selection.Font.Color = 12632256
        Selection.Font.Italic = True
    ElseIf Selection.Font.Color = 12632256 Then
        Selection.Font.Color = vbBlue
        Selection.Font.Italic = False
    ElseIf Selection.Font.Color = vbBlue Then
        Selection.Font.Color = vbRed
        Selection.Font.Italic = False
    ElseIf Selection.Font.Color = vbRed Then
        Selection.Font.Color = vbWhite
        Selection.Font.Italic = False
    Else
        Selection.Font.Color = vbBlack
        Selection.Font.Italic = False
    End If
    
End Sub

' /* ---------------------------- Cycle Cell Style ---------------------------- */

Sub Fmt_CycleCellStyle_T()
Attribute Fmt_CycleCellStyle_T.VB_ProcData.VB_Invoke_Func = "T\n14"
    ' Cycle through different cell styles

    Dim cellColor1, cellColor2, fontcolor1, fontColor2, borderColor As Long
    Dim tmp    As String

    tmp = branddark1
    tmp2 = branddark2 ' <<<< Change this variable to change the dark cell color (ie. cellColor2)


    cellColor1 = RGB(255, 255, 150)
    cellColor2 = RGB(Left(tmp, 3), Mid(tmp, 5, 3), Right(tmp, 3))
    cellColor3 = RGB(Left(tmp2, 3), Mid(tmp2, 5, 3), Right(tmp2, 3))
    fontcolor1 = RGB(0, 0, 255)
    fontColor2 = RGB(255, 255, 255)
    borderColor = RGB(220, 220, 220)
    
    
    If Selection.Interior.Pattern = xlNone And Selection.Font.Bold = False Then
        Selection.Interior.Color = cellColor1
        Selection.Font.Color = vbBlue
        Selection.Borders.Color = borderColor
        Selection.Borders.LineStyle = xlContinuous
        Selection.Font.Bold = False
    ElseIf Selection.Interior.Color = cellColor1 And Selection.Font.Color = fontcolor1 And Selection.Font.Bold = False Then
        Selection.Interior.Color = cellColor2
        Selection.Font.Color = vbWhite
        Selection.Borders.LineStyle = xlNone
        Selection.Font.Bold = False
    ElseIf Selection.Interior.Color = cellColor2 And Selection.Font.Color = fontColor2 And Selection.Font.Bold = False Then
        Selection.Interior.Color = cellColor3
        Selection.Font.Color = vbWhite
        Selection.Borders.LineStyle = xlNone
        Selection.Font.Bold = False
    Else
        Selection.Interior.Pattern = xlNone
        Selection.Font.Bold = False
        Selection.Font.Color = vbBlack
    End If
    '
    '
    '
End Sub


' /* --------------------------- Cycle Border Design -------------------------- */

Sub Fmt_CycleBorders_B()
Attribute Fmt_CycleBorders_B.VB_ProcData.VB_Invoke_Func = "B\n14"

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
    Else
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
    End If
    
End Sub



' /* ----------------------- Decimal Increase / Decrease ---------------------- */

Sub Fmt_DecimalIncrease_J()
Attribute Fmt_DecimalIncrease_J.VB_ProcData.VB_Invoke_Func = "J\n14"
    ' Ctrol + Shift + J
    Application.CommandBars.FindControl(ID:=398).Execute
    
End Sub

Sub Fmt_DecimalDecrease_K()
Attribute Fmt_DecimalDecrease_K.VB_ProcData.VB_Invoke_Func = "K\n14"
    ' Ctrol + Shift + K
    Application.CommandBars.FindControl(ID:=399).Execute
End Sub

' /* -------------------------------- Alignment ------------------------------- */

Sub Fmt_CenterAcrossSelection_M()
Attribute Fmt_CenterAcrossSelection_M.VB_ProcData.VB_Invoke_Func = "M\n14"
    '
    ' Macro3 Macro
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .WrapText = True
        .MergeCells = False
    End With
End Sub

Sub Fmt_Alignment_E()
Attribute Fmt_Alignment_E.VB_ProcData.VB_Invoke_Func = "E\n14"

    If Selection.HorizontalAlignment = xlGeneral Then
        Selection.HorizontalAlignment = xlRight
    ElseIf Selection.HorizontalAlignment = xlRight Then
        Selection.HorizontalAlignment = xlCenter
    Else
        Selection.HorizontalAlignment = xlGeneral
    End If
    
End Sub

Sub Fmt_WrapText_W()
Attribute Fmt_WrapText_W.VB_ProcData.VB_Invoke_Func = "W\n14"

    If Selection.WrapText = True Then
        Selection.WrapText = False
    Else
        Selection.WrapText = True
    End If
    
End Sub

'------------------------ Adds a custom suffix ------------------------ */
Sub Fmt_CustomSuffix_H()
Attribute Fmt_CustomSuffix_H.VB_ProcData.VB_Invoke_Func = "H\n14"
    Dim suffix As String
    Dim newFmt As String
    suffix = Application.InputBox("Insert the custom suffix" & _
                vbCr & " ", Type:=2, Title:="Insert New Sheet")

    newFmt = "_-#,##0_-""" & suffix & """"
    Selection.NumberFormat = newFmt

End Sub
    


'------------------------ Sheets No Grid and Zoom ------------------------ */

Function Fnc_NoGridZoom(ws_name As String)
    Sheets(ws_name).Select
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = sheetZoomLevel
    Application.GoTo Reference:=Range("a1"), Scroll:=True
End Function

' /* ----------------------- Current Sheet no Grid Zoon ----------------------- */

Sub Fmt_CurrentSheetNoGridZoom70_G()
Attribute Fmt_CurrentSheetNoGridZoom70_G.VB_ProcData.VB_Invoke_Func = "G\n14"
    ' Shortcut Ctrl+shift+G
    ' WorksheetGridZoom Macro
    Fnc_NoGridZoom (ActiveSheet.Name)
End Sub


'# -------------------------- All sheets No Grid Zoom ------------------------- #

Sub Fmt_AllSheetsNoGridZoom()
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
            ActiveWindow.Zoom = sheetZoomLevel
            Application.GoTo Reference:=Range("a1"), Scroll:=True
        End If
    Next ws
    
    cursheet.Activate
    
    Set ws = Nothing
    Set curcell = Nothing
    Set cursheet = Nothing
    
End Sub



' /* ------------------- All Sheets Font Type and Font Size ------------------- */

Sub Fmt_AllSheets_FontTypdSizeChange()
    Dim ws     As Worksheet
    For Each ws In Worksheets
        With ws
            .Cells.Font.Name = "calibri"
            .Cells.Font.Size = 11
        End With
    Next
    
End Sub


'------------------------ Auto Color the Selection ------------------------

Sub Fmt_AutoColorSelection_A()
Attribute Fmt_AutoColorSelection_A.VB_ProcData.VB_Invoke_Func = "A\n14"
    ' Keyboard Shortcut: Ctrl+Shift+A
    Dim cell   As Range, constantCell As Range, formulaCells As Range
    Dim cellFormula As String

    ' define colors for Model Coloring
    Const constColorBlue As String = "000,000,200" 'Blue Color
    Const linksOtherSheetsColor As String = "000,125,000" ' Green Color
    Const externalLinksColor As String = "125,000,125" ' Purple Colo

    With Selection
        On Error Resume Next
        Set constantCell = .SpecialCells(xlCellTypeConstants, xlNumbers)
        Set formulaCells = .SpecialCells(xlCellTypeFormulas, 23)
        On Error GoTo 0
    End With
    
    ' Color the constant numbers with Blue
    If Not constantCell Is Nothing Then
        constantCell.Font.Color = RGB(Left(constColorBlue, 3), Mid(constColorBlue, 5, 3), Right(constColorBlue, 3))
    End If
    
    ' check all cells with formulas
    If Not formulaCells Is Nothing Then
        For Each cell In formulaCells
            cellFormula = cell.Formula
            
            ' identify the cells with links to other workbooks
            If cellFormula Like "*.xls*]*!*" Then
                cell.Font.Color = RGB(Left(externalLinksColor, 3), Mid(externalLinksColor, 5, 3), Right(externalLinksColor, 3))
                
                
                ' identify the cells with links to other sheets
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
                cell.Font.Color = RGB(Left(linksOtherSheetsColor, 3), Mid(linksOtherSheetsColor, 5, 3), Right(linksOtherSheetsColor, 3))
            Else
                cell.Font.Color = vbBlack
            End If
        Next cell
    End If
End Sub



'------------------------ Auto Color All worksheets -----------------------

Sub Fmt_AllSheetsAutoColor()

    ' colors cells in all worksheets based on the content of the cells Constants, formula, links etc

    Dim ws     As Worksheet
    Dim cell   As Range, constantCell As Range, formulaCells As Range
    Dim cellFormula As String


    ' define colors for Model Coloring
    Const constColorBlue As String = "000,000,200" 'Blue Color
    Const linksOtherSheetsColor As String = "000,125,000" ' Green Color
    Const externalLinksColor As String = "125,000,125" ' Purple Colo

    For Each ws In Worksheets
        On Error Resume Next
        Set constantCell = ws.Cells.SpecialCells(xlCellTypeConstants, xlNumbers)
        Set formulaCells = ws.Cells.SpecialCells(xlCellTypeFormulas, 23)
        On Error GoTo 0
        
        If Not constantCell Is Nothing Then
            constantCell.Font.Color = RGB(Left(constColorBlue, 3), Mid(constColorBlue, 5, 3), Right(constColorBlue, 3))
        End If
        
        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells
                cellFormula = cell.Formula
                
                If cellFormula Like "*.xls*]*!*" Then
                    cell.Font.Color = RGB(Left(externalLinksColor, 3), Mid(externalLinksColor, 5, 3), Right(externalLinksColor, 3))
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
                    cell.Font.Color = RGB(Left(linksOtherSheetsColor, 3), Mid(linksOtherSheetsColor, 5, 3), Right(linksOtherSheetsColor, 3))
                Else
                    cell.Font.Color = vbBlack
                End If
            Next cell
        End If
        
    Next ws
    
End Sub



' /* -------------------------------------------------------------------------- */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /*                             Worksheets Section                             */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /* -------------------------------------------------------------------------- */


' -------------------------------- Clean Sheet ------------------------------- '

Sub sht_FullCleansheets_Option()

    'For current sheeet
    'Remove gridlines, set zoom to sheetZoomLevel%, go to A1
    'expand outlines
    'For the used range Unhide and autofit rows and columns and change font
    Dim ws     As Worksheet
    Dim PropertyOption As Integer
    Dim curcell As Range
    Dim cursheet As Worksheet
    Set cursheet = ActiveSheet
    Set curcell = ActiveCell
    
    
    PropertyOption = Application.InputBox("Clean current sheet or all sheets" & _
                     "(Must be 1, 2)" & vbCr & vbCr & "   [1] Current Sheet" & vbCr & _
                     "   [2] All worksheets" & vbCr & " ", Type:=1, Title:="Scope of Clean up")
    
    ' Change default style
    
    
    'Handle If User Cancels
    If PropertyOption = 0 Then Exit Sub
    
    If PropertyOption = 1 Then
        'if the user selcted current sheet only
        If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
        
        With ActiveSheet
            'Remove gridlines, set zoom to sheetZoomLevel%, go to A1
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = sheetZoomLevel
            Application.GoTo Reference:=Range("a1"), Scroll:=True
            'expand outlines
            .Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
            
            With .UsedRange
                'Unhide and autofit rows and columns
                .EntireColumn.AutoFit
                .EntireRow.AutoFit
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
        
    ElseIf PropertyOption = 2 Then
        'if the user selected all sheets
        With ActiveWorkbook
            .Styles("Normal").Font.Name = "Calibri"
            .Styles("Normal").Font.Size = 11
        End With
        
        For Each ws In ActiveWorkbook.Worksheets
            ws.Activate
            'Remove gridlines, set zoom to sheetZoomLevel%, go to A1
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = sheetZoomLevel
            Application.GoTo Reference:=Range("a1"), Scroll:=True
            
            If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
                ActiveSheet.ShowAllData
            End If
            
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
    Else
        Exit Sub
    End If
    
End Sub



' ----------------------------- Reset Used Rnage ----------------------------- '

Sub Sht_ResetUsedRange()
    'Helps to reset the usedrange by deleting rows and columns AFTER your true used range

    Dim myLastRow As Long
    Dim myLastCol As Long
    Dim dummyRng As Range
    Dim AnyMerged As Variant
    Dim curUR  As String

    curUR = ActiveSheet.UsedRange.Address
    
    'Check for merged cells
    AnyMerged = ActiveSheet.UsedRange.MergeCells
    If AnyMerged = True Or IsNull(AnyMerged) Then
        MsgBox "There are merged cells on this sheet." & vbCrLf & _
               "The macro will not work with merged cells.", vbOKOnly + vbCritical, "Macro will be Stopped"
        Exit Sub
    End If
    
    With ActiveSheet
        myLastRow = 0
        myLastCol = 0
        Set dummyRng = .UsedRange
        On Error Resume Next
        myLastRow = _
                    .Cells.Find("*", After:=.Cells(1), _
                    LookIn:=xlFormulas, LookAt:=xlWhole, _
                    searchdirection:=xlPrevious, _
                    SearchOrder:=xlByRows).Row
        myLastCol = _
                    .Cells.Find("*", After:=.Cells(1), _
                    LookIn:=xlFormulas, LookAt:=xlWhole, _
                    searchdirection:=xlPrevious, _
                    SearchOrder:=xlByColumns).Column
        On Error GoTo 0
        
        If myLastRow * myLastCol = 0 Then
            .Columns.Delete
        Else
            .Range(.Cells(myLastRow + 1, 1), _
                                    .Cells(.Rows.count, 1)).EntireRow.Delete
            .Range(.Cells(1, myLastCol + 1), _
                                    .Cells(1, .Columns.count)).EntireColumn.Delete
        End If
    End With
    
    MsgBox (curUR & " changed to " & ActiveSheet.UsedRange.Address)
    
End Sub



' -------------------------- Insert sheet with names ------------------------- '

Sub Sht_insertwithName_I()
Attribute Sht_insertwithName_I.VB_ProcData.VB_Invoke_Func = "I\n14"

    Dim NewSheet As Worksheet
    Dim sheetname As String


    sheetname = Application.InputBox("Insert the name of the new Sheet" & _
                vbCr & " ", Type:=2, Title:="Insert New Sheet")
    
    'The Inputbox types and their input types
    '0   A Formula
    '1   A Number
    '2   Text (Default)
    '4   A logical value (True or False)
    '8   A cell reference, as a Range object
    '16  An error value, such as #N/A
    '64  An array of values
    
    'Handle If User Cancels
    If sheetname = "" Then                        ' if user doesnt enter anything exit sub
        Exit Sub
    ElseIf sheetname = "False" Then               ' if user cancels exit sub
        Exit Sub
    End If
    
    ' check if a sheet exists by that name and if not create sheet
    If Fnc_Sheet_Exists(sheetname) = False Then
        Set NewSheet = Sheets.Add(After:=ActiveSheet)
        NewSheet.Name = sheetname
        Fnc_NoGridZoom (NewSheet.Name)
    Else
        MsgBox ("Sheetname already exists")
    End If
    
End Sub

' ------------------ Create Sheets from Selected Name Range ------------------ '

Sub Sht_CreateSheetsFromSelectedRange()
    'PURPOSE: Create new worksheets from a list of names within a table
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim NewSheet As Worksheet
    Dim cell   As Range
    Dim cursheet As Worksheet

    Set cursheet = ActiveSheet
    
    'Opitimize Code
    Application.ScreenUpdating = False
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbOKOnly, "List Unique Values Macro"
        Exit Sub
    End If
    
    
    'Create a new worksheet for every name inside the table
    For Each cell In Selection
        If Sht_Fnc_SheetExists(cell.Value) = False And cell.Value <> "" Then
            Set NewSheet = Sheets.Add(After:=Sheets(Sheets.count))
            NewSheet.Name = cell.Value
        End If
    Next cell
    
    cursheet.Activate
    
    'Opitimize Code
    Application.ScreenUpdating = True
    
End Sub

' /* -------------- Funcion to check whether a sheet already exist -------------- */

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



' /* ----------------------------- List All sheets ---------------------------- */

Sub Sht_SheetsList()

    Dim ws     As Worksheet
    Dim firstCell      As Integer
    Dim wsName As String

    firstCell = 3
    
    If Sht_Fnc_SheetExists("SheetsList") = True Then
        Worksheets("SheetsList").Range("B:B").Clear
    Else
        Sheets.Add(before:=Worksheets(1)).Name = "SheetsList"
    End If
    
    With Worksheets("sheetslist").Range("B2")
        .Value = "List of Worksheets"
        .Font.Bold = True
        .Font.Underline = True
    End With
    
    For Each ws In Worksheets
        If ws.Name <> "SheetsList" Then
            
            If ws.Visible = True Then
                wsName = ws.Name
            Else
                wsName = ws.Name & " (Hidden)"
            End If
            
            Sheets("SheetsList").Hyperlinks.Add _
                                                Anchor:=Sheets("SheetsList").Cells(firstCell, 2), Address:="", SubAddress:= _
                                                "'" & ws.Name & "'!A1", TextToDisplay:=wsName
            firstCell = firstCell + 1
            
        End If
        
    Next ws
    
    Worksheets("SheetsList").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    Application.GoTo Reference:=Range("a1"), Scroll:=True
    
End Sub




' /* ----------------------- Hide and Unhide Worksheets ----------------------- */

Sub Sht_SheetsUnhide()
    'Asks the user whether he needs to rehide
    'If he wants, copies the names of the hiddens seets to a tempsheet
    'then unhide all the sheets
    'Then he can rehide them. At the rehiding the tempsheet is deleted
    'If he doesnt want to re-hide it just unhides everything

    Dim ws     As Worksheet
    Dim count  As Integer
    Dim hws    As Worksheet
    Dim r      As Integer
    Dim checksheet As Boolean
    Dim Answer
    
    Answer = MsgBox("Do you want to re-hide?", vbYesNoCancel)
    If Answer = vbCancel Then End
    
    r = 1
    
    If Answer = vbYes Then
        checksheet = Sht_Fnc_SheetExists("temphidden")
        
        If checksheet = True Then
            Set hws = Worksheets("temphidden")
        Else
            Set hws = Worksheets.Add
            hws.Name = "temphidden"
        End If
        
        hws.Cells.Clear
        hws.Visible = False
        
        For Each ws In Worksheets
            
            If ws.Visible = False And ws.Name <> "temphidden" Then
                hws.Cells(r, 1).Value = ws.Name
                r = r + 1
                ws.Visible = xlSheetVisible
            End If
            
        Next ws
        
    Else
        For Each ws In Worksheets
            If ws.Visible = False And ws.Name <> "temphidden" Then
                ws.Visible = xlSheetVisible
                r = r + 1
            End If
        Next ws
    End If
    
    If r = 1 Then
        MsgBox ("you didnt have any hidden sheets")
    Else
        MsgBox ("you had " & r - 1 & " hidden sheets")
    End If
    
End Sub


Sub Sht_SheetsHideBack()
    Dim cell   As Range
    Dim sheetname As String

    If Sht_Fnc_SheetExists("temphidden") = True And Worksheets("temphidden").Range("A1").Value <> "" Then
        For Each cell In Worksheets("temphidden").Range("A1").CurrentRegion
            sheetname = cell.Value
            
            Worksheets(sheetname).Visible = False
            
        Next cell
        Application.DisplayAlerts = False
        Worksheets("temphidden").Delete
        Application.DisplayAlerts = True
    Else
        MsgBox ("You have not stored the hidden sheets")
    End If
    
End Sub



' /* -------------------------------------------------------------------------- */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /*                             Tools Section                                  */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /*                                                                            */
' /* -------------------------------------------------------------------------- */
/

Sub Tool_copyAcross_R()
Attribute Tool_copyAcross_R.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' copyAcross Macro: Copies the current cell or range to the same contiguos range and the column to the left.
' This is the ususal copy down that a professional will always use.
'


    Application.ScreenUpdating = False
    
    Dim currCell As Range
    Dim topCell As Range
    Dim rightCell As Range
    Dim bottomCell As Range
    Dim targetRange As Range
    Dim rowOffset As Long
    
    Set currCell = Selection
           
    currCell.Copy
    
    Set topCell = Selection.End(xlUp)
    rowOffset = currCell.Row - topCell.Row
    If IsEmpty(topCell.Offset(0, 1).Value) Then Exit Sub
    
    Set rightCell = topCell.End(xlToRight)
    Set bottomCell = rightCell.Offset(rowOffset, 0)
    
    If bottomCell.Address = currCell.Address Then
        Set targetRange = bottomCell ' If bottomCell and currCell are the same
    Else
        Set targetRange = currCell.Worksheet.Range(bottomCell, currCell)
    End If
    
    targetRange.PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False ' Clear clipboard after pasting and removes marching ants
    Application.ScreenUpdating = True
    
End Sub

Sub Tool_copyDown_D()
Attribute Tool_copyDown_D.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' copyDown Macro: Copies the current cell or range to the same contiguos range and the column to the left.
' This is the ususal copy down that a professional will always use.
'

'
    Application.ScreenUpdating = False
    Dim currCell As Range
    Dim leftCell As Range
    Dim bottomCell As Range
    Dim rightCell As Range
    Dim targetRange As Range
    Dim rowOffset As Long
    
    Set currCell = Selection
    
    Set leftCell = currCell.End(xlToLeft)
    columnOffset = currCell.Column - leftCell.Column
    
    If IsEmpty(currCell.Value) Then Exit Sub
    If IsEmpty(leftCell.Offset(1, 0).Value) Then Exit Sub
    
    Set bottomCell = leftCell.End(xlDown)
    Set rightCell = bottomCell.Offset(0, columnOffset)
    
    If leftCell.Address = currCell.Address Then
        Set targetRange = leftCell ' If bottomCell and currCell are the same
    Else
        Set targetRange = currCell.Worksheet.Range(rightCell, currCell)
    End If
        
    currCell.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False ' Clear clipboard after pasting and removes marching ants
    Application.ScreenUpdating = True

End Sub



' ------------------- Find All, Select them highlight them ------------------- '

Sub Tool_FindSelectAndHighlightAllCells()
    'Find Select and highlight all cells
    'Shortcut Ctrl+Shift+f

    Dim c      As Range, FoundCells As Range
    Dim firstaddress As String
    Dim fnd    As String
    Dim search_within As Range
    Dim LastCell As Range

    Application.ScreenUpdating = False
    
    'What value do you want to find?
    fnd = InputBox("I want to hightlight cells containing...", "Highlight")
    
    'Select the search range based on whether there is a selection or not
    If Selection.Cells.count > 1 Then
        Set search_within = Selection
    Else
        Set search_within = ActiveSheet.UsedRange
    End If
    
    Set LastCell = search_within.Cells(search_within.Cells.count)
    
    With ActiveSheet
        'find first cell that contains "rec"
        Set c = search_within.Find(What:=fnd, After:=LastCell)
        
        'if the search returns a cell
        If Not c Is Nothing Then
            'note the address of first cell found
            firstaddress = c.Address
            Do
                'FoundCells is the variable that will refer to all of the
                'cells that are returned in the search
                If FoundCells Is Nothing Then
                    Set FoundCells = c
                Else
                    Set FoundCells = Union(c, FoundCells)
                End If
                'find the next instance of "rec"
                Set c = search_within.FindNext(c)
            Loop While Not c Is Nothing And firstaddress <> c.Address
            
            'after entire sheet nsearched, select all found cells
            FoundCells.Select
            FoundCells.Interior.Color = RGB(255, 255, 0)
            Application.ScreenUpdating = True
            
            c.Activate
        Else
            'if no cells were found in search, display msg
            Application.ScreenUpdating = True
            MsgBox "No cells found."
        End If
    End With
End Sub





' -------------- Highlight all Duplicates in the selected range -------------- '

Sub Tool_DuplicatesHighlight()
    'highglights the duplicates in the selected range. Doesnt work for non-contiguous selection
    Dim cell
    For Each cell In Selection
        If WorksheetFunction.CountIf(Selection, cell.Value) > 1 Then
            cell.Interior.ColorIndex = 6
        End If
    Next cell
End Sub




' ------------------- Remove duplicates with error handling ------------------ '

Sub Tool_DuplicatesRemove()
    'PURPOSE: Remove duplicate cell values within a selected cell range
    'SOURCE: www.TheSpreadsheetGuru.com

    Dim rng    As Range
    Dim x      As Integer

    'Optimize code execution speed
    Application.ScreenUpdating = False
    
    'Determine range to look at from user's selection
    On Error GoTo InvalidSelection
    Set rng = Selection
    On Error GoTo 0
    
    'Determine if multiple columns have been selected
    If rng.Columns.count > 1 Then
        On Error GoTo InputCancel
        x = InputBox("Multiple columns were detected in your selection. " & _
            "Which column should I look at? (Number only!)", "Multiple Columns Found!", 1)
        On Error GoTo 0
    Else
        x = 1
    End If
    
    'Optimize code execution speed
    Application.Calculation = xlCalculationManual
    
    'Remove entire row
    rng.RemoveDuplicates Columns:=x
    
    'Change calculation setting to Automatic
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
    'ERROR HANDLING
InvalidSelection:
    MsgBox "You selection is not valid", vbInformation
    Exit Sub
    
InputCancel:
    
End Sub




' ------------------------ Clean and trim the selection ---------------------- '

Sub Tool_CleanTrimCells_Evaluate()
    'PURPOSE: A Fast way to Clean/Trim cell values in user selection
    'AUTHOR: Armando Montes
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim rng    As Range
    Dim Area   As Range

    If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
    Else
        Application.ReferenceStyle = xlA1
    End If
    
    'Weed out any formulas from selection
    If Selection.Cells.count = 1 Then
        Set rng = Selection
    Else
        Set rng = Selection.SpecialCells(xlCellTypeConstants)
    End If
    
    'Trim and Clean cell values
    For Each Area In rng.Areas
        Area.Value = Evaluate("IF(ROW(" & Area.Address & "),CLEAN(TRIM(" & Area.Address & ")))")
    Next Area
    
End Sub




' --------------- Ask to save workbook before running the macro -------------- '

Sub Fnc_AskToSave()
    'PURPOSE: Ask user if he would like to save before executing rest of code
    'SOURCE: www.TheSpreadsheetGuru.com

    Dim UserAnswer As Long
    Dim SaveAsChoice As Long
    Dim SavePath As String
    Dim FileExt As String
    Dim ExtNumber As Long

    'Ask user if he wants to save before executing
    If ThisWorkbook.Saved = False Then
        UserAnswer = MsgBox("Would you like to save before running?", vbYesNoCancel, "Save?")
        
        If UserAnswer = vbCancel Then Exit Sub    'User clicked cancel button
        
        If UserAnswer = vbYes Then
            If ThisWorkbook.Path <> "" Then
                'Need to SaveAs
                SaveAsChoice = Application.FileDialog(msoFileDialogSaveAs).Show
                If SaveAsChoice <> 0 Then
                    SavePath = Application.FileDialog(msoFileDialogSaveAs).SelectedItems(1)
                    
                    'Determine File Extension Number for SaveAs
                    FileExt = Right(SavePath, Len(SavePath) - InStrRev(SavePath, "."))
                    
                    'Get File Format Number (based off of extension)
                    Select Case FileExt
                        Case "xlsx": ExtNumber = 51
                        Case "xlsm": ExtNumber = 52
                        Case "xlsb": ExtNumber = 50
                        Case "xls": ExtNumber = 56
                    End Select
                    
                    ThisWorkbook.SaveAs SavePath, ExtNumber
                Else
                    Exit Sub                      'User clicked cancel button
                End If
            Else
                ThisWorkbook.Save
            End If
        End If
    End If
    
    'Insert the rest of you code here...
    
End Sub



' --------------------- Load brand colors to the recent colors --------------------- '


Sub Tool_Load2RecentColors()
    'PURPOSE: Use A List Of RGB Codes To Load Colors Into Recent Colors Section of Color Palette
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim ColorList As Variant
    Dim CurrentFill As Variant

    Application.ScreenUpdating = False
    
    
    'Array List of RGB Color Codes to Add To Recent Colors Section (Max 10)
    ColorList = Array("000,168,168", "052,024,082", "001,180,184", "119,094,136")
    
    'Store ActiveCell's Fill Color (if applicable)
    If ActiveCell.Interior.ColorIndex <> xlNone Then CurrentFill = ActiveCell.Interior.Color
    
    'Loop Through List Of RGB Codes And Add To Recent Colors
    For x = LBound(ColorList) To UBound(ColorList)
        ActiveCell.Interior.Color = RGB(Left(ColorList(x), 3), Mid(ColorList(x), 5, 3), Right(ColorList(x), 3))
        DoEvents
        SendKeys "%hhm~"
        DoEvents
    Next x
    
    'Return ActiveCell Original Fill Color
    If CurrentFill = Empty Then
        ActiveCell.Interior.ColorIndex = xlNone
    Else
        ActiveCell.Interior.Color = CurrentFill
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "Colors Loaded to the Color Palette"
    
End Sub




'--------------------------------------------------------

'Sub Tool_PasteValues_V()
    ' NOTE: THIS IS REDUNDANT. EXCEL IMPLEMENTED THIS NATIVELY
    ' PasteValues Macro
    ' Paste only Values from copied cell data.
    '
    ' Keyboard Shortcut: Ctrl+Shift+V
    '
'    On Error Resume Next
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                           :=False, Transpose:=False
'    ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:= _
'                           False
' End Sub

'--------------------------------------------------------

Sub Fmt_AllGreyBorders_C()
Attribute Fmt_AllGreyBorders_C.VB_ProcData.VB_Invoke_Func = "C\n14"
    '
    ' Keyboard Shortcut: Ctrl+Shift+E
    '
    '

    If Selection.Borders.LineStyle = xlNone Then
        Selection.Borders.Color = RGB(220, 220, 220)
        Selection.Borders.LineStyle = xlContinuous
    Else
        Selection.Borders.LineStyle = xlNone
    End If
    
    
End Sub


'-------------------- Convert the selection to numbers -------------------- */

Sub Tool_Convert2Numbers_values()
    'specify the range which suits your purpose
    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With
End Sub

' /* -------------- Convert the selection to numbers Old Version -------------- */

Sub Tool_Convert2Numbers_txt2Col()
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                            :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.NumberFormat = "0"
End Sub



' /* ------------------------------ Wrap if error ----------------------------- */

Sub Tool_WrapIfError_v2()

    'PURPOSE: Add an IFERROR() Function around all the selected cells' formulas. _
    Also handles if IFERROR is already wrapped around formula.
    'SOURCE: www.TheSpreadsheetGuru.com
    
    Dim rng    As Range
    Dim cell   As Range
    Dim AlreadyIFERROR As Boolean
    Dim RemoveIFERROR As Boolean
    Dim TestEnd1 As String
    Dim TestEnd2 As String
    Dim TestStart As String
    Dim MyFormula As String
    Dim x      As String
    
    'Determine if a single cell or range is selected
    If Selection.Cells.count = 1 Then
        Set rng = Selection
        If Not rng.HasFormula Then GoTo NoFormulas
    Else
        'Get Range of Cells that Only Contain Formulas
        On Error GoTo NoFormulas
        Set rng = Selection.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
    End If
    
    'Get formula from First cell in Selected Range
    MyFormula = rng(1, 1).Formula
    
    'Create Test Strings To Determine if IFERROR formula has already been added
    TestEnd1 = Chr(34) & Chr(34) & ")"
    TestEnd2 = ",0)"
    TestStart = Left(MyFormula, 9)
    
    'Determine How we want to modify formula
    If Right(MyFormula, 3) = TestEnd1 And TestStart = "=IFERROR(" Then
        Beg_String = ""
        End_String = "0)"                         '=IFERROR([formula],0)
        AlreadyIFERROR = True
    ElseIf Right(MyFormula, 3) = ",0)" And TestStart = "=IFERROR(" Then
        RemoveIFERROR = True
    Else
        Beg_String = "=IFERROR("
        End_String = "," & Chr(34) & Chr(34) & ")" '=IFERROR([formula],"")
        AlreadyIFERROR = False
    End If
    
    'Loop Through Each Cell in Range and modify formula
    For Each cell In rng.Cells
        x = cell.Formula
        
        If RemoveIFERROR = True Then
            cell = "=" & Mid(x, 10, Len(x) - 12)
        ElseIf AlreadyIFERROR = False Then
            cell = Beg_String & Right(x, Len(x) - 1) & End_String
        Else
            cell = Left(x, Len(x) - 3) & End_String
        End If
        
    Next cell
    
    Exit Sub
    
    'Error Handler
NoFormulas:
    MsgBox "There were no formulas found in your selection!"
    
End Sub


' /* ----------------- List unique values in first two columns ---------------- */

Sub Tool_Duplicates_ListUniqueValues()
    'Create a list of unique values from the selected column

    Dim rSelection As Range
    Dim ws     As Worksheet
    Dim vArray() As Long
    Dim i      As Long
    Dim iColCount As Long
    Dim allCellsBlank As Boolean
    
    


    'Check that a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbOKOnly, "List Unique Values Macro"
        Exit Sub
    End If
    
    'Store the selected range
    Set rSelection = Selection
   

    'Add a new worksheet
    Set ws = Worksheets.Add
    
    'Copy/paste selection to the new sheet
    rSelection.Copy
    
    With ws.Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        '.PasteSpecial xlPasteValuesAndNumberFormats
    End With
    
    'Load array with column count
    'For use when multiple columns are selected
    iColCount = rSelection.Columns.count
    ReDim vArray(1 To iColCount)
    For i = 1 To iColCount
        vArray(i) = i
    Next i
    
    'Remove duplicates
    ws.UsedRange.RemoveDuplicates Columns:=vArray(i - 1), Header:=xlGuess
    
    'Remove blank cells (optional)
    On Error Resume Next
    ws.UsedRange.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlShiftUp
    On Error GoTo 0
    
    'Autofit column
    ws.Columns("A").AutoFit
    
    'Exit CutCopyMode
    Application.CutCopyMode = False
    
End Sub



' /* ----------------------- Remove same sheet reference ---------------------- */

Sub Tool_RemoveSameSheetReferences()
    'PURPOSE: Removes Sheet References from formulas when not needed
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim sht    As Worksheet
    Dim fndList As Variant
    Dim rplcList As Variant
    Dim x      As Long

    Set sht = ActiveSheet
    
    fndList = Array("'" & sht.Name & "'!", sht.Name & "!")
    rplc = ""
    
    'Optimize Code
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'Loop through each item in Array lists
    For x = LBound(fndList) To UBound(fndList)
        sht.Cells.Replace What:=fndList(x), Replacement:=rplc, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
                          SearchFormat:=False, ReplaceFormat:=False
    Next x
    
    'Optimize Code
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub




' /* ------------------ Functions to extract numbers and text ----------------- */

'This VBA code will create a function to get the numeric part from a string
Function Fnc_GetNumeric(CellRef As String)
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
    Next i
    GetNumeric = Result
    
    Set StringLength = Nothing
    
End Function


'This VBA code will create a function to get the text part from a string
Function Fnc_GetText(CellRef As String)
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If Not (IsNumeric(Mid(CellRef, i, 1))) Then Result = Result & Mid(CellRef, i, 1)
    Next i
    GetText = Result
End Function



' /* ------- Turn Off Pivot Table Autofit Column Width On Update Setting ------ */

Sub Tool_Pivot_TurnAutoFitOff()
    'PURPOSE: Turn off Autofit Column Width On Update Setting for every Pivot Table
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim pvt    As PivotTable

    For Each pvt In ActiveSheet.PivotTables
        pvt.HasAutoFormat = False
    Next pvt
    
End Sub


' /* ----------- Change Pivot table fields to sum and set formating ----------- */

Sub Tool_Pivot_ChangeFields()
    'Update 20141127
    'select all the privot table and execute

    Dim xPF    As PivotField
    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    With WorkRng.PivotTable
        .ManualUpdate = True
        For Each xPF In .DataFields
            With xPF
                .Function = xlSum
                .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* """"-""""??_);_(@_)"
            End With
        Next
        .ManualUpdate = False
    End With
End Sub




' /* ------------------------ Swap two selected ranges ------------------------ */

Sub Tool_SwapTwoAreas()
    'PURPOSE: Swap two selected ranges' formulas/values with each other
    'SOURCE: www.TheSpreadsheetGuru.com

    Dim rng    As Range
    Dim StoredRng As Variant

    Set rng = Selection
    
    If rng.Areas.count <> 2 Or rng.Areas(1).Cells.count <> rng.Areas(2).Cells.count Then
        MsgBox "Please select two ranges that are the same size before running this macro"
        Exit Sub
    End If
    
    'Store first selected cell area
    StoredRng = rng.Areas(1).Cells.Formula
    
    'Swap first area with the second
    rng.Areas(1).Cells.Formula = rng.Areas(2).Cells.Formula
    
    'Populate second area with the first
    rng.Areas(2).Cells.Formula = StoredRng
    
End Sub



' /* ------------------------------ Backup a copy ----------------------------- */

Sub Tool_WorkBook_Backup()

    Dim BaseFileName As String
    Dim FileNameArray() As String
    Dim Comment     As String


    ' to check whether the workbook is saved at least once. if a workbook is not saved at least once (ie. book1, book2 etc) it will not have a path.
    If ActiveWorkbook.Path = "" Then
        MsgBox "The workbook is not saved"
        Exit Sub
    End If
    
    
    'Step 1: Check whether the user wants to create a comment in file name
    Comment = InputBox("Insert the comment", 1)
    
    ' if user has entered a comment, format it with the paranthesis
    If Comment <> "" Then
        Comment = " (" & Comment & ")"
    End If
    ' if user cancels, stop the macro without saving a file
    If StrPtr(Comment) = 0 Then Exit Sub
    
    'Step 2: Create a Backup of a Workbook with Current Date in the Same folder
    ' Preapare the file name and extension
    FileNameArray = Split(ActiveWorkbook.Name, ".")
    Debug.Print FileNameArray(0)
    Debug.Print FileNameArray(1)
    
    If ActiveWorkbook.Name = "Personal.xlsb" Then Exit Sub
    
    ' Save a copy
    ActiveWorkbook.SaveCopyAs _
    Filename:=ActiveWorkbook.Path & "\" & _
    FileNameArray(0) & " " & _
    Format(Now(), "YYMMDD_hhmmss") & _
    Comment & "." & _
    FileNameArray(1)
    
End Sub


' /* ----------------------- Documenting and commenting ----------------------- */


' this macro is added on 201119

Sub Tool_Documenting()
    ' add sheets to the workbook
    Documenting_Create_sheet ("|| ")
    Documenting_Create_sheet (" ||")
    Documenting_Create_sheet ("||")
    Documenting_Create_sheet ("VC")
    Documenting_Create_sheet ("BG")
    
    
End Sub


Function Documenting_Create_sheet(WorkSheet_Name As String)

    Dim shtColor As Integer
    shtColor = 1                                  '1 - black, 2-White ,3-Red, 4-Green , 5-Blue, 6-yellow, 7-Pink, 8-Lightblue, 9-Browsn, 10-Dark Green

    If Fnc_Sheet_Exists(WorkSheet_Name) = False Then
        Sheets.Add(before:=Sheets(1)).Name = WorkSheet_Name
        Sheets(WorkSheet_Name).Tab.ColorIndex = shtColor
        Fnc_NoGridZoom (WorkSheet_Name)           ' calls the nogrid zoom function
    Else
        Sheets(WorkSheet_Name).Tab.ColorIndex = shtColor
        Fnc_NoGridZoom (WorkSheet_Name)
    End If
    
End Function


Function Fnc_Sheet_Exists(WorkSheet_Name As String) As Boolean
    ' checks whether a sheet exists in the current workbook by the same name
    Dim ws     As Worksheet

    Fnc_Sheet_Exists = False
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = WorkSheet_Name Then
            Fnc_Sheet_Exists = True
        End If
    Next
    
End Function



' /* --------------------------- Remove white fills --------------------------- */


Sub Fmt_WhiteFill_Toggle()
    'PURPOSE: Change empty cell fill colors to white based on selection (Toggle)
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim rng    As Range
    Dim cell   As Range
    Dim GatherRange As Range

    'Test for Empty Filled Cells in User Selection
    For Each cell In Selection.Cells
        If cell.Interior.ColorIndex = xlNone Then
            If GatherRange Is Nothing Then
                Set GatherRange = cell
            Else
                Set GatherRange = Union(GatherRange, cell)
            End If
        End If
    Next cell
    
    'Were any empty filled cells found?
    If Not GatherRange Is Nothing Then
        'Whiteout all applicable cells
        GatherRange.Interior.Color = RGB(255, 255, 255)
    Else
        'Test for White Fills
        For Each cell In Selection.Cells
            If cell.Interior.Color = RGB(255, 255, 255) Then
                If GatherRange Is Nothing Then
                    Set GatherRange = cell
                Else
                    Set GatherRange = Union(GatherRange, cell)
                End If
            End If
        Next cell
        
        'Remove White Fills
        If Not GatherRange Is Nothing Then
            GatherRange.Interior.ColorIndex = xlNone
        End If
    End If
    
End Sub



' /* ------------------------ Insert Picture as Comment ----------------------- */

Sub Tool_InsertPictureComment()
    'PURPOSE: Insert an Image into the ActiveCell's Comment
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim PicturePath As String
    Dim CommentBox As Comment

    '[OPTION 1] Explicitly Call Out The Image File Path
    'PicturePath = "C:\Users\chris\Desktop\Image1.png"

    '[OPTION 2] Pick A File to Add via Dialog (PNG or JPG)
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Title = "Select Comment Image"
        .ButtonName = "Insert Image"
        .Filters.Clear
        .Filters.Add "Images", "*.png; *.jpg"
        .Show
        
        'Store Selected File Path
        On Error GoTo UserCancelled
        PicturePath = .SelectedItems(1)
        On Error GoTo 0
    End With
    
    'Clear Any Existing Comment
    Application.ActiveCell.ClearComments
    
    'Create a New Cell Comment
    Set CommentBox = Application.ActiveCell.AddComment
    
    'Remove Any Default Comment Text
    CommentBox.Text Text:=""
    
    'Insert The Image and Resize
    CommentBox.Shape.Fill.UserPicture (PicturePath)
    CommentBox.Shape.ScaleHeight 6, msoFalse, msoScaleFromTopLeft
    CommentBox.Shape.ScaleWidth 4.8, msoFalse, msoScaleFromTopLeft
    
    'Ensure Comment is Hidden (Swith to TRUE if you want visible)
    CommentBox.Visible = False
    
    Exit Sub
    
    'ERROR HANDLERS
UserCancelled:
    
End Sub


' /* --------------- Finding the celll fill color and text color -------------- */



Sub Tool_ex_GetRGBColor_Font()
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

Sub Tool_ex_GetRGBColor_Fill()
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



' /* --------------- Save as a Text File in the current location -------------- */

Sub Tool_Saveastxtfile()

    'Save the file as a text file in the current folder. If the source file is not saved, saves a txt file in current users desktop

    ' Copy activesheet to the new workbook
    Dim pth    As String, wbname As String

    pth = ActiveWorkbook.Path
    
    If pth = "" Then
        pth = Environ("USERPROFILE") & "\Desktop"
    End If
    
    
    wbname = InputBox("Please input a name for the textfile")
    
    ActiveSheet.Copy
    MsgBox "This new workbook will be saved in" & pth
    
    'Save new workbook as MyWb.xls(x) into the folder where ActiveWorkbook is stored
    ActiveWorkbook.SaveAs pth & "\" & wbname & "txt", FileFormat:=xlText, CreateBackup:=False
    
    MsgBox "It is saved as " & ActiveWorkbook.FullName & vbLf & "Press OK to close it"
    
    ' Close the saved copy
    ActiveWorkbook.Close False
    
End Sub



' /* -------------------------------------------------------------------------- */
' /*                                Shape Section                               */
' /* -------------------------------------------------------------------------- */

' --- Change whehter the shapes/charts resize with row/column widths change -- '
Sub Shape_ResizeMoveProperty()
    'PURPOSE: Change All Shapes Object Placement Property (User Input)
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim shp    As Shape
    Dim cht    As ChartObject
    Dim PropertyOption As Integer

    'Retrieve Input from User
    PropertyOption = Application.InputBox("Change Everything To What Placement Property?" & _
    "(Must be 1, 2, or 3)" & vbCr & vbCr & _
    "   [1] Move and Size with Cells" & vbCr & _
    "   [2] Move but Don't Size with Cells" & vbCr & _
    "   [3] Don't Move or Size with Cells" & _
    vbCr & " ", Type:=1, Title:="Placement Property For All")
    
    'Handle If User Cancels
    If PropertyOption = 0 Then Exit Sub
    
    'Loop Through Shapes & Controls
    For Each shp In ActiveSheet.Shapes
        shp.Placement = PropertyOption
    Next shp
    
    'Loop Through Charts
    For Each cht In ActiveSheet.ChartObjects
        cht.Placement = PropertyOption
    Next cht
    
End Sub

' /* -------- cycle shapes/charts resize with row/column widths change -------- */

Sub Shape_ResizeMoveProperty_Cycle()
    'PURPOSE: Change All Shapes Object Placement Property (Cycle)
    'ENUMERATIONS: xlMoveAndSize = 1, xlMove = 2, xlFreeFloating = 3

    Dim shp    As Shape
    Dim cht    As ChartObject
    Dim PropertyOption As Integer

    'Determine which Placement to Apply. will exit after finding the 1st shape
    For Each shp In ActiveSheet.Shapes
        PropertyOption = Choose(shp.Placement, 2, 3, 1)
        GoTo PlacementChoosen
    Next shp
    
    For Each cht In ActiveSheet.ChartObjects
        PropertyOption = Choose(cht.Placement, 2, 3, 1)
        GoTo PlacementChoosen
    Next cht
    
    'Nothing Found
    MsgBox "No objects were found to adjust the placement property"
    Exit Sub
    
PlacementChoosen:
    
    'Handle If User Cancels
    If PropertyOption = 0 Then Exit Sub
    
    'Loop Through Shapes & Controls
    For Each shp In ActiveSheet.Shapes
        shp.Placement = PropertyOption
    Next shp
    
    'Loop Through Charts
    For Each cht In ActiveSheet.ChartObjects
        cht.Placement = PropertyOption
    Next cht
    
    'Report action taken to user
    Select Case PropertyOption
        Case 1: MsgBox "All Charts & Shapes set to: " & Chr(34) & "Move and Size with Cells" & Chr(34)
        Case 2: MsgBox "All Charts & Shapes set to: " & Chr(34) & "Move but Don't Size with Cells" & Chr(34)
        Case 3: MsgBox "All Charts & Shapes set to: " & Chr(34) & "Don't Move or Size with Cells" & Chr(34)
    End Select
    
End Sub



' /* ------------------------------ Align Shapes ------------------------------ */

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



' /* ----------- Make all charts in workbook plot non-visible cells ----------- */

Sub Shape_Chart_PlotNonVisibleCells()
    'PURPOSE: Make all charts in workbook plot non-visible cells
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim sht    As Worksheet
    Dim cht    As ChartObject
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





