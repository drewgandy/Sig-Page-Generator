VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmInsertSigPage 
   Caption         =   "Mark Page For Signatures"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   OleObjectBlob   =   "FrmInsertSigPage.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmInsertSigPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
    If Trim(TxtNewParty.Text) <> "" Then
        LstParties.AddItem Trim(TxtNewParty.Text)
        TxtNewParty.Text = ""
    End If
End Sub


Private Sub LstParties_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim SnippetFoundOnCurPage As Boolean
    Dim currentPosition As Range
    Set currentPosition = Selection.Range
    currentpage = Selection.Information(wdActiveEndAdjustedPageNumber)
    Selection.WholeStory   'Select entire document

    With Selection.Find
      .ClearFormatting
      .MatchWholeWord = True
      .MatchCase = False
      .Font.Hidden = True
      .MatchWildcards = True
      Do
        bResult = .Execute(FindText:="##Signature Page-*##")
        If bResult Then
            Debug.Print currentpage & "," & Selection.Information(wdActiveEndPageNumber)
          If currentpage = Selection.Information(wdActiveEndPageNumber) Then
            Debug.Print "found on page"
            Selection.Move Unit:=wdCharacter, Count:=1
            Selection.Font.Hidden = True
            If CmbSigPageLimit.Text <> "No Limit" Then
                Selection.TypeText "##Signature Page-" & LstParties.Text & " [Limit=" & Trim(CmbSigPageLimit.Text) & "]##"
            Else
                Selection.TypeText "##Signature Page-" & LstParties.Text & "##"
            End If
            Selection.Font.Hidden = False
'            SnippetFoundOnCurPage = True
            currentPosition.Select
            Exit Sub
          End If
        End If
      Loop Until Not bResult
    End With
    
    currentPosition.Select
'    If SnippetFoundOnCurPage = False Then
        Selection.Font.Hidden = True
        Selection.TypeText "NOTE: This text and the below snippet are hidden text and will not appear when printed.  Do not edit the below snippet, which is used to generate signature pages." & Chr(11)
        If CmbSigPageLimit.Text <> "No Limit" Then
            Selection.TypeText "##Signature Page-" & LstParties.Text & " [Limit=" & Trim(CmbSigPageLimit.Text) & "]##"
        Else
            Selection.TypeText "##Signature Page-" & LstParties.Text & "##"
        End If
        Selection.Font.Hidden = False
    'End If
End Sub



Private Sub TxtNewParty_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
         CmdAdd_Click
    End If
End Sub

Private Sub TxtNewParty_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub UserForm_Click()
Dim SigPageProperties() As String
Dim PropName As String
Dim PropValue As String
Dim PageLimit As Integer

' trim excess spaces around hashtags
txt = "## Signature Page-Borrower [Limit=1, prop2=5]  ##"
Do While txt <> Replace(txt, " ##", "##")
    txt = Replace(txt, " ##", "##")
Loop
Do While txt <> Replace(txt, "## ", "##")
    txt = Replace(txt, "## ", "##")
Loop

If Len(txt) - InStrRev(txt, "]") = 2 Then
    PropStr = Mid(txt, InStrRev(txt, "["), (InStrRev(txt, "]") - InStrRev(txt, "[")) + 1)
    
    txt = Replace(txt, PropStr, "")
    Do While txt <> Replace(txt, " ##", "##")
        txt = Replace(txt, " ##", "##")
    Loop
    
    'Trim Properties list of brackets, populate SigPageProperties array and trim any spaces around each property
    PropStr = UCase(Replace(Mid(PropStr, 2, Len(PropStr) - 2), " ", ""))
    SigPageProperties = Split(PropStr, ",")
    For i = LBound(SigPageProperties) To UBound(SigPageProperties)
        PropName = Left(SigPageProperties(i), InStr(SigPageProperties(i), "=") - 1)
        PropValue = Right(SigPageProperties(i), Len(SigPageProperties(i)) - InStr(SigPageProperties(i), "="))
        Select Case PropName
            Case "LIMIT"
                PageLimit = Int(PropValue)
            Case "PROP2"
        End Select
    Next i
End If
End Sub

Private Sub UserForm_Initialize()
With LstParties
    .AddItem "Borrower"
    .AddItem "Lender"
    .AddItem "Guarantor"
    .AddItem "General Partner"
    .AddItem "Equity Investor"
End With
With CmbSigPageLimit
    .AddItem "No Limit"
    .Text = "No Limit"
    For i = 1 To 20
        .AddItem Trim(Str(i))
    Next
End With
End Sub
