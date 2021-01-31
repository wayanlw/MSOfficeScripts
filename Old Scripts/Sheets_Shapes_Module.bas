Attribute VB_Name = "Sheets_Shapes_Module"
Option Explicit


Sub SheetsUnhide()
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
        checksheet = sheetExists("temphidden")
        
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


Sub SheetsHideBack()
    Dim cell   As Range
    Dim sheetname As String


    If sheetExists("temphidden") = True And Worksheets("temphidden").Range("A1").Value <> "" Then
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

Sub SheetsList()

    Dim ws     As Worksheet
    Dim x      As Integer
    Dim wsName As String

    x = 3
    
    If sheetExists("SheetsList") = True Then
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
                                                Anchor:=Sheets("SheetsList").Cells(x, 2), Address:="", SubAddress:= _
                                                "'" & ws.Name & "'!A1", TextToDisplay:=wsName
            x = x + 1
        End If
    Next ws
    
    Worksheets("SheetsList").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    Application.GoTo Reference:=Range("a1"), Scroll:=True
    
End Sub

Sub SheetsAutoColor()

    Dim ws     As Worksheet
    Dim cell   As Range
    Dim constantCell As Range
    Dim formulaCells As Range
    Dim cellFormula As String

    For Each ws In Worksheets
        On Error Resume Next
        Set constantCell = ws.Cells.SpecialCells(xlCellTypeConstants, xlNumbers)
        Set formulaCells = ws.Cells.SpecialCells(xlCellTypeFormulas, 23)
        On Error GoTo 0
        
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
        
    Next ws
        
End Sub

Function sheetExists(sheetToFind As String) As Boolean
    Dim sheet  As Worksheet

    sheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
End Function

Sub Sheet_insertwithName_F()
Attribute Sheet_insertwithName_F.VB_ProcData.VB_Invoke_Func = "F\n14"

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
    If sheetname = "" Then Exit Sub
    
    Set NewSheet = Sheets.Add(After:=ActiveSheet)
    NewSheet.Name = sheetname
    
    
End Sub

Sub sht_FullCleanAllsheets()

    'For current sheeet
    'Remove gridlines, set zoom to 70%, go to A1
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
        
        With ActiveSheet
            
            
            'Remove gridlines, set zoom to 70%, go to A1
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = 70
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
        
    Else
        Exit Sub
        
    End If
    
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
'################################################# Change whehter the shapes and charts resize with change in row/column widths #####################################################
'################################################################################################################################

Sub Shape_ResizeMoveProperty()
    'PURPOSE: Change All Shapes Object Placement Property (User Input)
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim shp    As Shape
    Dim cht    As ChartObject
    Dim PropertyOption As Integer

    'Retrieve Input from User
    PropertyOption = Application.InputBox("Change Everything To What Placement Property?" & _
    "(Must be 1, 2, or 3)" & vbCr & vbCr & "   [1] Move and Size with Cells" & vbCr & _
    "   [2] Move but Don't Size with Cells" & vbCr & "   [3] Don't Move or Size with Cells" & _
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

