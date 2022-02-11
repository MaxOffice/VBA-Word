Attribute VB_Name = "HeadingWordCountModule"
Option Explicit

Private Const MACROTITLE = "Heading Word Count"

Public Sub CountHeadingContentWords()
    Dim headingRange As Range
    Set headingRange = ActiveDocument.Content
    
    Dim findHeading As Find
    Set findHeading = headingRange.Find
    
    With findHeading
        .ClearFormatting
        .Style = ActiveDocument.Styles("Heading 1")
        .Forward = True
        .Wrap = wdFindStop
        .Execute
    End With
    
    If Not findHeading.Found Then
        MsgBox "Could not find a heading in this document.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    Dim headingSelectionStart As Integer
    Dim headingSelectionEnd As Integer
    Dim currentHeading As String
    
    headingSelectionStart = -1
    
    Do While findHeading.Found
    
        If headingSelectionStart = -1 Then
            headingSelectionStart = headingRange.Start
        Else
            headingSelectionEnd = headingRange.Start - 1
            
            countWordsInRange headingSelectionStart, headingSelectionEnd
            
            headingSelectionStart = headingSelectionEnd + 1
        End If
        
        currentHeading = headingRange.Text
        findHeading.Execute
        
    Loop
    
    headingSelectionEnd = ActiveDocument.Range.End
    If headingSelectionStart < ActiveDocument.Range.End Then
        countWordsInRange headingSelectionStart, headingSelectionEnd
    End If
End Sub

Private Sub countWordsInRange(ByVal rangeStart As Integer, ByVal rangeEnd As Integer)
    Dim countRange As Range
    Set countRange = ActiveDocument.Range(rangeStart, rangeEnd)
    MsgBox "Heading: " & currentHeading & vbCrLf & _
            "Word count: " & countRange.Words.Count & vbCrLf & _
            "Accurate word count: " & countRange.ComputeStatistics(wdStatisticWords), _
            , MACROTITLE
End Sub
