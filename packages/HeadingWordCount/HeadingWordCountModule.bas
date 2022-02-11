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
    Dim headingsForm As HeadingWordCountForm
    
    headingSelectionStart = -1
    
    Set headingsForm = New HeadingWordCountForm
    headingsForm.Clear
    
    Do While findHeading.Found
    
        If headingSelectionStart = -1 Then
            headingSelectionStart = headingRange.Start
        Else
            headingSelectionEnd = headingRange.Start - 1
            
            countWordsInRange headingSelectionStart, headingSelectionEnd, currentHeading, headingsForm
            
            headingSelectionStart = headingSelectionEnd + 1
        End If
        
        currentHeading = headingRange.Text
        findHeading.Execute
        
    Loop
    
    headingSelectionEnd = ActiveDocument.Range.End
    If headingSelectionStart < ActiveDocument.Range.End Then
        countWordsInRange headingSelectionStart, headingSelectionEnd, currentHeading, headingsForm
    End If
    
    headingsForm.Finalize
    headingsForm.Show vbModal
    
    Unload headingsForm
    Set headingsForm = Nothing
End Sub

Private Sub countWordsInRange( _
                ByVal rangeStart As Integer, _
                ByVal rangeEnd As Integer, _
                ByVal currentHeading As String, _
                ByVal headingsForm As HeadingWordCountForm _
            )
    Dim countRange As Range
    Set countRange = ActiveDocument.Range(rangeStart, rangeEnd)
    headingsForm.Append "Heading: " & currentHeading & vbCrLf & vbCrLf & _
            "Word count: " & countRange.Words.Count & vbCrLf & _
            "Accurate word count: " & countRange.ComputeStatistics(wdStatisticWords) & vbCrLf & _
            vbCrLf
End Sub
