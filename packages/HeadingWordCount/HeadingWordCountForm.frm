VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HeadingWordCountForm 
   Caption         =   "Heading Word Count"
   ClientHeight    =   4970
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8340.001
   OleObjectBlob   =   "HeadingWordCountForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HeadingWordCountForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private headingsSummary As String

Private Sub cmdCopyAndClose_Click()
    With txtHeadingWordCount
        .SelStart = 0
        .SelLength = Len(headingsSummary)
        .Copy
        .SelLength = 0
    End With
    Me.Hide
End Sub

Public Sub Append(ByVal data As String)
    headingsSummary = headingsSummary & data
End Sub

Public Sub Prepare()
    With txtHeadingWordCount
        .Text = headingsSummary
        .SelStart = 0
        .SelLength = 0
    End With
End Sub

Public Sub Clear()
    headingsSummary = ""
End Sub


Private Sub Label1_Click()

End Sub
