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
