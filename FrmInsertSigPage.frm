VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmInsertSigPage 
   Caption         =   "Mark Page For Signatures"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4890
   OleObjectBlob   =   "FrmInsertSigPage.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmInsertSigPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MacroFilename As String
Private Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1 'bring to top and stay there
Private Const SWP_NOMOVE = &H2 'don't move window
Private Const SWP_NOSIZE = &H1 'don't size window

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Dim lHwnd As Long

Private Sub CmdAdd_Click()
    If Trim(TxtNewParty.Text) <> "" Then
        LstParties.AddItem Trim(TxtNewParty.Text)
        TxtNewParty.Text = ""
    End If
End Sub


Private Sub Label2_Click()

End Sub

Private Sub LstParties_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim SnippetFoundOnCurPage As Boolean
    Dim currentPosition As Range
    Dim PropList As String
    
    Set currentPosition = Selection.Range
    currentpage = Selection.Information(wdActiveEndAdjustedPageNumber)
    Selection.WholeStory   'Select entire document
    Selection.Expand unit:=wdStory
With Selection.Find
    .ClearFormatting
    .MatchWholeWord = False
    .MatchCase = False
    .Font.Hidden = True
    .MatchWildcards = True
    Do
        bResult = .Execute(FindText:="##Signature Page-*##")
        If bResult Then
            Debug.Print currentpage & "," & Selection.Information(wdActiveEndPageNumber)
            If currentpage = Selection.Information(wdActiveEndPageNumber) Then
                Debug.Print "found on page"
                Selection.Move unit:=wdCharacter, Count:=1
                Selection.Font.Hidden = True
                If CmbSigPageLimit.Text <> "No Limit" Then
                    PropList = PropList & ", LIMIT=" & Trim(CmbSigPageLimit.Text) & ", "
                End If
                If CmbPageRange.Text <> "0" Then
                    PropList = PropList & ", PAGES=" & Trim(Str(Int(CmbPageRange.Text) + 1)) & ", "
                End If
                PropList = Replace(PropList, ", ", "", , 1)
                If Len(PropList) <> 0 Then
                    Selection.TypeText "##Signature Page-" & LstParties.Text & " [" & PropList & "]##"
                Else
                    Selection.TypeText "##Signature Page-" & LstParties.Text & "##"
                End If
                Selection.Font.Hidden = False
                currentPosition.Select
                Exit Sub
            End If
        End If
    Loop Until Not bResult
End With
    
    currentPosition.Select
        Selection.Font.Hidden = True
        Selection.TypeText "NOTE: This text and the below snippet are hidden text and will not appear when printed.  Do not edit the below snippet, which is used to generate signature pages." & Chr(11)
        If CmbSigPageLimit.Text <> "No Limit" Then
            PropList = PropList & "LIMIT=" & Trim(CmbSigPageLimit.Text) & ", "
        End If
        If CmbPageRange.Text <> "0" Then
                PropList = PropList & ", PAGES=" & Trim(Str(Int(CmbPageRange.Text) + 1))
        End If
        PropList = Replace(PropList, ", ", "", , 1)
        If Len(PropList) > 0 Then
            Selection.TypeText "##Signature Page-" & LstParties.Text & " [" & PropList & "]##"
        Else
            Selection.TypeText "##Signature Page-" & LstParties.Text & "##"
        End If 'If CmbSigPageLimit.Text <> "No Limit" Then
    '    Selection.TypeText "##Signature Page-" & LstParties.Text & " [Limit=" & Trim(CmbSigPageLimit.Text) & "]##"
        'Else
        '    Selection.TypeText "##Signature Page-" & LstParties.Text & "##"
        'End If
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

End Sub

Private Sub UserForm_Initialize()
If Right(ActiveDocument.name, 4) = "docm" Then
    MacroFilename = ActiveDocument.name
    Documents(MacroFilename).ActiveWindow.Visible = False
End If
'Documents(MacroFilename).ActiveWindow.Visible = True



'    UserForm1.Show vbModeless

    lHwnd = FindWindow("ThunderDFrame", "Mark Page For Signatures")
    
    If lHwnd <> GetForegroundWindow Then
        Call SetWindowPos(lHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If


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
With CmbPageRange
    For i = 0 To 20
        .AddItem Trim(Str(i))
    Next
End With
End Sub

Private Sub UserForm_Terminate()
If MacroFilename <> "" Then Documents(MacroFilename).ActiveWindow.Visible = True

End Sub
