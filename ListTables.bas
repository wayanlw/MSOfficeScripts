

Option Explicit
Const sheetZoomLevel As Integer = 80

Sub ListTables()

    'Declare variables and data types
    Dim tbl    As ListObject
    Dim ws     As Worksheet
    Dim i      As Single, j As Single

    'Check whther the sheet "wayTableList" exists,clear the sheet. Else create one
    If Sht_Fnc_SheetExists("wayTableList") = True Then
        Worksheets("wayTableList").Cells.Clear
    Else
        Sheets.Add(Before:=Worksheets(1)).Name = "wayTableList"
    End If

    'Creating the header section and formatting the sheet
    With Worksheets("wayTableList")

            ActiveWindow.DisplayGridlines = False
            ActiveWindow.Zoom = sheetZoomLevel

        With .Range("B2")
            .Value = "List of Worksheets"
            .Font.Bold = True
            .Font.Underline = True
        End With

        Range("B3").Value = "Range"
        Range("c3").Value = "TableName"
        Range("d3").Value = "Column Names >>"
        Range("B3:d3").Font.Bold = True

    End With

    ' initialize the starting row
    i = 4

    'Go through each worksheet in the worksheets object collection
    For Each ws In Worksheets

        'Go through all Excel defined Tables located in the current WS worksheet object
        For Each tbl In ws.ListObjects

            'Save Excel defined Table name to cell in column A
            Sheets("wayTableList").Hyperlinks.Add _
            Anchor:=Sheets("wayTableList").Cells(i, 2), Address:="", SubAddress:= _
                        "'" & ws.Name & "'!" & tbl.Range.Address, TextToDisplay:=ws.Name
            Worksheets("wayTableList").Range("A1").Cells(i, 3).Value = tbl.Name
            Worksheets("wayTableList").Range("A1").Cells(i, 4).Value = tbl.Range.Rows.count & "x" & tbl.Range.Columns.count



            'Iterate through columns in Excel defined Table
            For j = 4 To tbl.Range.Columns.count

                'Save header name to cell next to table name
                Worksheets("wayTableList").Range("A1").Cells(i, j + 1).Value = tbl.Range.Cells(1, j)

                'Continue with next column
            Next j

            'Add 1 to variable i
            i = i + 1

            'Continue with next Excel defined Table
        Next tbl

        'Continue with next worksheet
    Next ws

    'Exit macro
End Sub

Sub tool_ListPivotsWithDetails()
    Dim ws     As Worksheet
    Dim wsSD   As Worksheet
    Dim lstSD  As ListObject
    Dim pt     As PivotTable
    Dim rngPT  As Range
    Dim wsPL   As Worksheet
    Dim rngSD  As Range
    Dim rngHead As Range
    Dim pt2    As PivotTable
    Dim rngPT2 As Range
    Dim rCols  As Range
    Dim rRows  As Range
    Dim RowPL  As Long
    Dim RptCols As Long
    Dim SDCols As Long
    Dim SDHead As Long
    Dim lBang  As Long
    Dim nm     As Name
    Dim strSD  As String
    Dim strRefRC As String
    Dim strRef As String
    Dim strWS  As String
    Dim strAdd As String
    Dim strFix As String
    Dim lRowsInt As Long
    Dim lColsInt As Long
    Dim CountPT As Long
    On Error Resume Next

    RptCols = 13
    RowPL = 2

    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            CountPT = CountPT + 1
            If CountPT > 0 Then Exit For
        Next pt
        If CountPT > 0 Then Exit For
    Next ws

    If CountPT = 0 Then
        MsgBox "No pivot tables in this workbook"
        GoTo exitHandler
    End If

    If Sht_Fnc_SheetExists("wayPivotTableList") = True Then
        Worksheets("wayPivotTableList").Cells.Clear
    Else
        Sheets.Add(Before:=Worksheets(1)).Name = "wayPivotTableList"
    End If

    Set wsPL = Worksheets("wayPivotTableList")
    wsPL.Activate



    With wsPL
        .Range(.Cells(1, 1), .Cells(1, RptCols)).Value _
                         = Array("Worksheet", _
                         "Ws PTs", _
                         "PT Name", _
                         "PT Range", _
                         "PTs Same Rows", _
                         "PTs Same Cols", _
                         "PivotCache", _
                         "Source Data", _
                         "Records", _
                         "Data Cols", _
                         "Data Heads", _
                         "Head Fix", _
                         "Refreshed")
    End With

    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            lRowsInt = 0
            lColsInt = 0
            Set rngPT = pt.TableRange2

            For Each pt2 In ws.PivotTables
                If pt2.Name <> pt.Name Then
                    Set rngPT2 = pt2.TableRange2
                    Set rRows = Intersect(rngPT.Rows.EntireRow, _
                        rngPT2.Rows.EntireRow)
                    If Not rRows Is Nothing Then
                        lRowsInt = lRowsInt + 1
                    End If
                    Set rCols = Intersect(rngPT.Columns.EntireColumn, _
                        rngPT2.Columns.EntireColumn)
                    If Not rCols Is Nothing Then
                        lColsInt = lColsInt + 1
                    End If
                End If
            Next pt2

            If pt.PivotCache.SourceType = 1 Then  'xlDatabase
                Set nm = Nothing
                strSD = ""
                strAdd = ""
                strFix = ""
                SDCols = 0
                SDHead = 0
                Set rngHead = Nothing
                Set lstSD = Nothing

                strSD = pt.SourceData

                'worksheet range?
                lBang = InStr(1, strSD, "!")
                If lBang > 0 Then
                    strWS = Left(strSD, lBang - 1)
                    strRefRC = Right(strSD, Len(strSD) - lBang)
                    strRef = Application.ConvertFormula( _
                             strRefRC, xlR1C1, xlA1)
                    Set rngSD = Worksheets(strWS).Range(strRef)
                    SDCols = rngSD.Columns.count
                    Set rngHead = rngSD.Rows(1)
                    SDHead = WorksheetFunction.CountA(rngHead)
                    GoTo AddToList
                End If

                'named range?
                Set nm = ThisWorkbook.Names(strSD)
                If Not nm Is Nothing Then
                    strAdd = nm.RefersToRange.Address
                    SDCols = nm.RefersToRange.Columns.count
                    Set rngHead = nm.RefersToRange.Rows(1)
                    SDHead = WorksheetFunction.CountA(rngHead)
                    GoTo AddToList
                End If

                'list object?
                For Each wsSD In ActiveWorkbook.Worksheets
                    Set lstSD = wsSD.ListObjects(strSD)
                    If Not lstSD Is Nothing Then
                        strAdd = lstSD.Range.Address
                        SDCols = lstSD.Range.Columns.count
                        Set rngHead = lstSD.HeaderRowRange
                        SDHead = WorksheetFunction.CountA(rngHead)
                        GoTo AddToList
                    End If
                Next
            End If

AddToList:
            If SDCols <> SDHead Then strFix = "X"
            With wsPL
                .Range(.Cells(RowPL, 1), _
                                     .Cells(RowPL, RptCols)).Value _
                                     = Array(ws.Name, _
                                     ws.PivotTables.count, _
                                     pt.Name, _
                                     pt.TableRange2.Address, _
                                     lRowsInt, _
                                     lColsInt, _
                                     pt.CacheIndex, _
                                     pt.SourceData, _
                                     pt.PivotCache.RecordCount, _
                                     SDCols, _
                                     SDHead, _
                                     strFix, _
                                     pt.PivotCache.RefreshDate)

                'add hyperlink to pt range
                .Hyperlinks.Add _
                Anchor:=.Cells(RowPL, 4), _
                Address:="", _
                SubAddress:="'" & ws.Name _
                & "'!" & pt.TableRange2.Address, _
                ScreenTip:=pt.TableRange2.Address, _
                TextToDisplay:=pt.TableRange2.Address
            End With

            RowPL = RowPL + 1
        Next pt
    Next ws

    'format headers and autofit the range
    With wsPL
        .Rows(1).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, RptCols)) _
                         .EntireColumn.AutoFit
    End With

exitHandler:
    Set wsPL = Nothing
    Set ws = Nothing
    Set pt = Nothing
    Exit Sub

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




