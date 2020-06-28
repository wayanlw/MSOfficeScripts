'#####################################################################################################################
'#####################################################################################################################
'#####################################################################################################################
'#####################################################################################################################
'#####################################################################################################################
'#####################################################################################################################
'#####################################################################################################################
'Saample use of array functions - Worksheet order

Sub Reorder_Sheets()

'PURPOSE: Order Worksheets in a custom way (works even if some of the worksheets are missing)
'SOURCE: www.TheSpreadsheetGuru.com

Dim x As Long
Dim myOrder As Variant

Application.DisplayAlerts = False
Application.ScreenUpdating = False

myOrder = Array("Sheet1", "Sheet4", "Sheet3", "Sheet6", "Sheet9", "Sheet11", "Sheet10", "Sheet5")

On Error Resume Next
  For x = UBound(myOrder) To LBound(myOrder) Step -1
    Worksheets(myOrder(x)).Move Before:=Worksheets(1)
  Next x
On Error GoTo 0

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub




'#####################################################################################################################
' Delete Blank Cells Or Entire Blank Rows Within A Range
Sub RemoveBlankCells()
'PURPOSE: Deletes single cells that are blank located inside a designated range
'SOURCE: www.TheSpreadsheetGuru.com

Dim rng As Range

'Store blank cells inside a variable
  On Error GoTo NoBlanksFound
    Set rng = Range("A1:A10").SpecialCells(xlCellTypeBlanks)
  On Error GoTo 0

'Delete blank cells and shift upward
  rng.Rows.Delete Shift:=xlShiftUp

Exit Sub

'ERROR HANLDER
NoBlanksFound:
  MsgBox "No Blank cells were found"

End Sub

Sub RemoveBlankRows()
'PURPOSE: Deletes any row with blank cells located inside a designated range
'SOURCE: www.TheSpreadsheetGuru.com

Dim rng As Range

'Store blank cells inside a variable
  On Error GoTo NoBlanksFound
    Set rng = Range("A1:A10").SpecialCells(xlCellTypeBlanks)
  On Error GoTo 0

'Delete entire row of blank cells found
  rng.EntireRow.Delete
  
Exit Sub

'ERROR HANLDER
NoBlanksFound:
  MsgBox "No Blank cells were found"

End Sub



'#####################################################################################################################




Sub CleanData()

'PURPOSE:Clean up selected data by trimming spaces, converting dates,
'and converting numbers to appropriate formats from text format
'AUTHOR: Ejaz Ahmed (www.StrugglingToExcel.Wordpress.com)
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim MessageAnswer As VbMsgBoxResult
Dim EachRange As Range
Dim TempArray As Variant
Dim rw As Long
Dim col As Long
Dim ChangeCase As Boolean
Dim ChangeCaseOption As VbStrConv
Dim rng As Range

'User Preferences
  ChangeCaseOption = vbProperCase
  ChangeCase = False

Set rng = Application.Selection

'Warn user if Range has Formulas
  If RangeHasFormulas(rng) Then
    MessageAnswer = MsgBox("Some of the cells contain formulas. " _
      & "Would you like to proceed and overwrite formulas with values?", _
      vbQuestion + vbYesNo, "Formulas Found")
    If MessageAnswer = vbNo Then Exit Sub
  End If

'Loop through each separate area the selected range may have
  For Each EachRange In rng.Areas
    TempArray = EachRange.Value2
      If IsArray(TempArray) Then
        For rw = LBound(TempArray, 1) To UBound(TempArray, 1)
          For col = LBound(TempArray, 2) To UBound(TempArray, 2)
            'Check if value is a date
              If IsDate(TempArray(rw, col)) Then
                TempArray(rw, col) = CDate(TempArray(rw, col))
              
            'Check if value is a number
              ElseIf IsNumeric(TempArray(rw, col)) Then
                TempArray(rw, col) = CDbl(TempArray(rw, col))
                  
            'Otherwise value is Text. Let's Trim it! (Remove any extraneous spaces)
              Else
                TempArray(rw, col) = Application.Trim(TempArray(rw, col))
                      
                'Change Case if the user wants to
                  If ChangeCase Then
                    TempArray(rw, col) = StrConv( _
                    TempArray(rw, col), ChangeCaseOption)
                  End If
              End If
          Next col
        Next rw
      Else
        'Handle with Single Cell selected areas
          If IsDate(TempArray) Then 'If Date
            TempArray = CDate(TempArray)
          ElseIf IsNumeric(TempArray) Then 'If Number
            TempArray = CDbl(TempArray)
          Else 'Is Text
            TempArray = Application.Trim(TempArray)
              'Handle case formatting (if necessary)
                If ChangeCase Then
                  TempArray = StrConv(TempArray, ChangeCaseOption)
                End If
          End If
      End If
      
    EachRange.Value2 = TempArray
    
  Next EachRange

'Code Ran Succesfully!
  MsgBox "Your data cleanse was successful!", vbInformation, "All Done!"

End Sub


Function RangeHasFormulas(ByRef rng As Range) As Boolean

'PURPOSE: Determine if given range has any formulas in it
'AUTHOR: Ejaz Ahmed (www.StrugglingToExcel.Wordpress.com)
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim TempVar As Variant

TempVar = rng.HasFormula

'Test Range
  If IsNull(TempVar) Then
    'Some of cells have fromulas
      RangeHasFormulas = True
  Else
    If TempVar = True Then
      'All cells have formulas
        RangeHasFormulas = True
    Else
      'None of cells have formulas
        RangeHasFormulas = False
    End If
  End If

End Function

'#####################################################################################################################

Sub SetShapeProperties()
'PURPOSE: Loop through all shapes on the ActiveSheet and adjust
'Object Positioning property from Size and Properties Dialog box
'SOURCE: www.TheSpreadsheetGuru.com/The-Code_Vault

Dim shp As Shape
Dim Counter As Long

'Loop through each shape (image) in ActiveSheet
  For Each shp In ActiveSheet.Shapes
    shp.Placement = xlFreeFloating
    Counter = Counter + 1
  Next shp
  
'Completion Message
  MsgBox Counter & " shapes were changed"

'OTHER OPTIONS:
  'xlMoveAndSize    Object is moved and sized with the cells.
  'xlMove           Object is moved with the cells.
  'xlFreeFloating   Object is free floating.

End Sub

Sub BreakExternalLinks()
'PURPOSE: Breaks all external links that would show up in Excel's "Edit Links" Dialog Box
'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

Dim ExternalLinks As Variant
Dim wb As Workbook
Dim x As Long

Set wb = ActiveWorkbook

'Create an Array of all External Links stored in Workbook
  ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)

'Loop Through each External Link in ActiveWorkbook and Break it
  For x = 1 To UBound(ExternalLinks)
    wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
  Next x

End Sub
