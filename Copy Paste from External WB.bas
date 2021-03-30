
' Used to copy a sheet from another workbook and value paste them to the current workbook



Sub getdataofclosedbook()

    Dim src    As Workbook
    Dim f      As String
    Dim sname  As String
    Dim i      As Integer

    For i = 1 To 2 'previously when consumer was also there this was 1 To 3

        fName = ""


        f = "\\xxx\Users$\xx\My Documents\SAP\SAP GUI\" & i & ".xls"
        If Dir(f) = "" Then
            MsgBox (f & "  doesn't exist")
        Else
            Select Case i
                Case 1
                        sname = "AUPH ZFIR"
                Case 2
                        sname = "NZPH ZFIR"
'               Case 3
'                       sname = "NZCT ZFIR"
            End Select


            Set src = Workbooks.Open("\\s0000783\Users$\auwwaya\My Documents\SAP\SAP GUI\" & i & ".xls", True, True)
            src.Activate
            src.Worksheets(1).Cells.Copy

            ThisWorkbook.Activate

            Worksheets(sname).Range("A1").PasteSpecial Paste:=xlValues
            Application.DisplayAlerts = False
            src.Close
            Application.DisplayAlerts = True
         End If
    Next i

End Sub
