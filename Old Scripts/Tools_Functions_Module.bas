Attribute VB_Name = "Tools_Functions_Module"
Sub Convert2Numbers()
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                            :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.NumberFormat = "0"
End Sub

Sub ConvertoNumbers()
    'specify the range which suits your purpose
    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With
End Sub


Sub RemoveSameSheetReferences()
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


Sub Email_sheet()
    '
    ' Send Macro
    '

    '
    ActiveSheet.Select
    ActiveSheet.Copy
    Application.Dialogs(xlDialogSendMail).Show
End Sub


Sub List_Unique_Values()
    'Create a list of unique values from the selected column
    'Source: https://www.excelcampus.com/vba/remove-duplicates-list-unique-values

    Dim rSelection As Range
    Dim ws     As Worksheet
    Dim vArray() As Long
    Dim i      As Long
    Dim iColCount As Long

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



'This VBA code will create a function to get the numeric part from a string
Function GetNumeric(CellRef As String)
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
    Next i
    GetNumeric = Result
    
    Set StringLength = Nothing
    
End Function


'This VBA code will create a function to get the text part from a string
Function GetText(CellRef As String)
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If Not (IsNumeric(Mid(CellRef, i, 1))) Then Result = Result & Mid(CellRef, i, 1)
    Next i
    GetText = Result
End Function

Sub ResetUsedRange()
    Dim myLastRow As Long
    Dim myLastCol As Long
    Dim dummyRng As Range
    Dim AnyMerged As Variant
    Dim curUR  As String
    'http://www.contextures.on.ca/xlfaqApp.html#Unused
    'Helps to reset the usedrange by deleting rows and columns AFTER your true used range

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

Sub Highlight_Duplicates()
    Dim cell
    For Each cell In Selection
        If WorksheetFunction.CountIf(Selection, cell.Value) > 1 Then
            cell.Interior.ColorIndex = 6
        End If
        
    Next cell
End Sub


Sub WrapIfError_v2()

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
