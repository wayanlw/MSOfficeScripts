' ---------------------------------------------------------------------------- '
'Excel is showing a hidden name range

' Excel.Currentworkbook() in powerquery shows an error
' Expression.Error: We couldn't find an Excel table named ''To BU'!_FilterDatabase'.
' Details:
' ---------------------------------------------------------------------------- '


' ---------------------------------------------------------------------------- '
' Solution. Run the below macro to unhide all named ranges

Sub unhideAllNames()
Dim tempName As Variant
For Each tempName In ActiveWorkbook.Names
        tempName.Visible = True
    Next
End Sub
' ---------------------------------------------------------------------------- '

' to get a list of hidden ranges

Sub Rprt()
Dim nm As Name, n As Long, y As Range, z As Worksheet
Application.ScreenUpdating = False
Set z = ActiveSheet
n = 2
With z
    .[a1:g65536].ClearContents
    .[a1:D1] = [{"Name","Sheet Name","Starting Range","Ending Range"}]
    For Each nm In ActiveWorkbook.Names
         On Error Resume Next
        .Cells(n, 1) = nm.Name
        .Cells(n, 2) = Range(nm).Parent.Name
        .Cells(n, 3) = nm.RefersToRange.Address(False, False)
        n = n + 1
        On Error GoTo 0
    Next nm
End With

Set y = z.Range("c2:c" & z.[c65536].End(xlUp).Row)
y.TextToColumns Destination:=z.[C2], DataType:=xlDelimited, _
    OtherChar:=":", FieldInfo:=Array(Array(1, 1), Array(2, 1))
[a:d].EntireColumn.AutoFit

Application.ScreenUpdating = True
End Sub
