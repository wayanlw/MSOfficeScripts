Option Explicit

Sub copyEPFG()
    Dim rngPart As Range
    Dim copied As String
    Dim firstCell As Boolean

    firstCell = True

    For Each rngPart In Application.Selection.Areas
        If firstCell = True Then
            copied = Right(rngPart.Value, Len(rngPart.Value) - 1)
            firstCell = False
        Else
            rngPart = copied
        End If
    Next

End Sub

Sub Fnc_CountIf()
    Dim src, item, dest As Range
    Dim wbname, Ssrc, Sitem As String

    Set src = Application.InputBox("Select the Range" & _
                vbCr & " ", Type:=8, Title:="Source range")
    wbname = ActiveWorkbook.Name
    Ssrc = "'[" & wbname & "]" & src.Parent.Name & "'!" & src.Address

    Set item = Application.InputBox("Select the Count value" & _
                vbCr & " ", Type:=8, Title:="Count Item")
    Sitem = item.Address(False, False)

    Set dest = Application.InputBox("Select the Destination" & _
                vbCr & " ", Type:=8, Title:="Destination")
    dest.Value = "=CountIf(" & Ssrc & "," & Sitem & ")"


End Sub

Sub insert_Lines()
    Dim noOfLines As Variant

    noOfLines = Application.InputBox("Insert the Number of Lines to Enter" & _
                vbCr & " ", Type:=1, Title:="Insert Rows Below")

    'The Inputbox types and their input types
    '0   A Formula
    '1   A Number
    '2   Text (Default)
    '4   A logical value (True or False)
    '8   A cell reference, as a Range object
    '16  An error value, such as #N/A
    '64  An array of values

    If noOfLines = "" Then                        ' if user doesnt enter anything exit sub
        Exit Sub
    ElseIf noOfLines = "False" Then               ' if user cancels exit sub
        Exit Sub
    ElseIf noOfLines <= 0 Then
        Exit Sub
    End If

    ActiveCell.EntireRow.Resize(noOfLines).Insert Shift:=xlDown
End Sub



