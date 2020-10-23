Sub BeautifyFormula()
    For Each c In Selection.Cells
        fmla = c.Formula
        If InStr(fmla, "    ") > 0 Then GoTo NextC
        newf = ""
        tabs = 0
        pos = 1
        inquotes = False
        While pos <= Len(fmla)
            S = Mid(fmla, pos, 1)
            newf = newf & S
            If S = """" Then
                inquotes = Not inquotes
            End If
            If Not inquotes Then
            Select Case S
            Case "("
                tabs = tabs + 1
                newf = newf & Chr(10) & gettabs(tabs)
            Case ")"
                tabs = tabs - 1
                newf = newf & Chr(10) & gettabs(tabs)
            Case ","
                newf = newf & Chr(10) & gettabs(tabs)
            Case Else
            End Select
            End If
            pos = pos + 1
        Wend
        c.Formula = newf
NextC:
    Next c
End Sub

Function gettabs(ByVal tabs) As String
    t = ""
    For i = 1 To tabs
        t = t & "    "
        't = t & Chr(9)
    Next i
    gettabs = t
End Function
