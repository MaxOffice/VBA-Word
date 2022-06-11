Attribute VB_Name = "HeadingWordCountModule"
'Version 0.0.3
'Created by Raj Chaudhuri

Option Explicit

Public outputstr As String

Private Const MACROTITLE = "Heading Word Count"
Public Sub CountHeadingContentWordsAction(rb As IRibbonControl)
    CountHeadingContentWords
End Sub
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
        MsgBox "Could not find any paragraph with Heading1 style in this document.", vbExclamation, MACROTITLE
        Exit Sub
    End If
    
    Dim headingSelectionStart As Long
    Dim headingSelectionEnd As Long
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
    
    headingsForm.Prepare
    headingsForm.Show vbModal
    
    Unload headingsForm
    
    Set headingsForm = Nothing
End Sub

Private Sub countWordsInRange( _
                ByVal rangeStart As Long, _
                ByVal rangeEnd As Long, _
                ByVal currentHeading As String, _
                ByVal headingsForm As HeadingWordCountForm _
            )
    Dim countRange As Range
    Set countRange = ActiveDocument.Range(rangeStart, rangeEnd)
            
    ' Info: Word count, Lines, Paragraphs, Characters, Pages - <Heading 1 Text>
    Dim tempStr As String
    
    ' Full stats
    tempStr = Replace(currentHeading, Chr(13), "") & vbTab & _
        countRange.ComputeStatistics(wdStatisticWords) & vbTab & _
        countRange.ComputeStatistics(wdStatisticParagraphs) & vbTab & _
        countRange.ComputeStatistics(wdStatisticLines) & vbTab & _
        countRange.ComputeStatistics(wdStatisticCharacters) & vbTab & _
        countRange.ComputeStatistics(wdStatisticPages) & vbCrLf
    
    ' Only word count <Heading><tab><WordCount>
    'tempStr = Replace(currentHeading, Chr(13), "") & vbTab & _
        countRange.ComputeStatistics(wdStatisticWords) & _
        vbCrLf
    
    headingsForm.Append tempStr
End Sub
