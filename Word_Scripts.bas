

'################################################################################################################################
' Cleanup a PDF Copy. Merge all paragraph breaks
'################################################################################################################################'

Sub Paragraph_Merger()
    'PURPOSE: Get rid of unecessary paragraph breaks (typically caused when copying PDF text)
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    'Remove unecessary paragraphs
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
    End With

    'Replace All Instances
    Selection.Find.Execute Replace:=wdReplaceAll

    'Remove any double spaces
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
    End With

    'Replace All Instances
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub

'----------------------------------------------- Revmove consecutive blanklines

Sub RemoveConsecutiveBlankLines()
    'PURPOSE: Remove Consecutive Blank Paragraphs Throughout the Entire Word Document
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim ParagraphCount As Long
    Dim Doc    As Document
    Dim rng    As Range

    Set Doc = ActiveDocument
    Set rng = Doc.Range
    ParagraphCount = Doc.Paragraphs.Count

    'Loop Through Each Paragraph (in reverse order)
    For x = ParagraphCount To 1 Step -1
        If x - 1 > 1 Then
            If rng.Paragraphs(x).Range.Text = vbCr And rng.Paragraphs(x - 1).Range.Text = vbCr Then
                rng.Paragraphs(x).Range.Delete
            End If
        End If
    Next x

End Sub

'################################################################################################################################
'################################# Revmove consecutive blanklines #####################################################
'################################################################################################################################

Sub RemoveConsecutiveBlankLines()
    'PURPOSE: Remove Consecutive Blank Paragraphs Throughout the Entire Word Document
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim ParagraphCount As Long
    Dim Doc    As Document
    Dim rng    As Range

    Set Doc = ActiveDocument
    Set rng = Doc.Range
    ParagraphCount = Doc.Paragraphs.Count

    'Loop Through Each Paragraph (in reverse order)
    For x = ParagraphCount To 1 Step -1
        If x - 1 > 1 Then
            If rng.Paragraphs(x).Range.Text = vbCr And rng.Paragraphs(x - 1).Range.Text = vbCr Then
                rng.Paragraphs(x).Range.Delete
            End If
        End If
    Next x

End Sub
