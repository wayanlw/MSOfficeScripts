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
