VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "Signature Page Generator"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10635
   OleObjectBlob   =   "FrmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SigPageArray() As Variant



Dim xCounter As Integer

 Sub TrackSigPages(ByVal PartyName As String, ByVal Filename As String, ByVal DocName As String)

 Dim intI As Integer, intJ As Integer
 If UBound(SigPageArray, 2) > 0 Then
    For i = (LBound(SigPageArray, 2) + 1) To (UBound(SigPageArray, 2))
        If SigPageArray(1, i) = PartyName Then
            SigPageArray(2, i) = SigPageArray(2, i) & "/" & Filename
            SigPageArray(3, i) = SigPageArray(3, i) & "/" & DocName
            GoTo PartyNameInArray
        End If
    Next
    ReDim Preserve SigPageArray(3, UBound(SigPageArray, 2) + 1)
    SigPageArray(1, UBound(SigPageArray, 2)) = PartyName
    SigPageArray(2, UBound(SigPageArray, 2)) = Filename
    SigPageArray(3, UBound(SigPageArray, 2)) = DocName
    GoTo PartyNameInArray
  Else
    ReDim SigPageArray(3, 1)
    SigPageArray(1, 1) = PartyName
    SigPageArray(2, 1) = Filename
    SigPageArray(3, 1) = DocName
  End If
  
PartyNameInArray:

End Sub


Private Sub CmbBrowseOutputFolder_Click()
Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
    If sItem <> "" Then TxtOutputFolder.Text = sItem
End Sub

Private Sub CmbDeleteFiles_Click()
Dim FirstSelected As Integer
FirstSelected = -1
Restart:
For i = 0 To LstFilenames.ListCount
    If LstFilenames.Selected(i) = True Then
        If FirstSelected = -1 Then FirstSelected = i
        LstFilenames.Selected(i) = False
        LstFilenames.RemoveItem (i)
        GoTo Restart:
    End If
Next

If FirstSelected <> 0 Then
    If FirstSelected = LstFilenames.ListCount Then
        LstFilenames.Selected(LstFilenames.ListCount - 1) = True
    ElseIf FirstSelected > -1 Then
        LstFilenames.Selected(FirstSelected) = True
    End If
ElseIf LstFilenames.ListCount > 0 Then
    LstFilenames.Selected(0) = True
End If

End Sub

Private Sub CmdChooseFiles_Click()
Dim dlgOpen As FileDialog
 Set dlgOpen = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
 With dlgOpen
    .AllowMultiSelect = True
    FileChosen = .Show
    If FileChosen = -1 Then
        For i = 1 To dlgOpen.SelectedItems.Count
            LstFilenames.AddItem dlgOpen.SelectedItems(i)
        Next i
    End If

 End With

End Sub



Private Sub CmdGenerateSigPages_Click()
  
  Dim myRange As Range
  Dim bResult As Boolean
  Dim objDocument As Document
  Dim SigPagesDoc As Document
  Dim SigPageFilename As String
  Dim OutputFolder As String
  Dim SigPageProperties() As String
  Dim PropName As String
  Dim PropValue As String
  Dim PageLimit As Integer
  Dim PropStr As String
  Dim SigPagePartyName As String
  Dim DocumentName As String
  Dim SigPageCount As Integer
  
  Dim StartTime As Double
  Dim SecondsElapsed As Double

  'Remember time when macro starts
  StartTime = Timer

  
  Application.Browser.Target = wdBrowsePage
  'Set SigPagesDoc = Documents.Add
  If Right(TxtOutputFolder.Text, 1) <> "\" Then OutputFolder = TxtOutputFolder + "\"
  ReDim SigPageArray(3, 0)
    For ifilename = 0 To LstFilenames.ListCount - 1
        LstFilenames.Selected(ifilename) = False
    Next
  For ifilename = 0 To LstFilenames.ListCount - 1
    LstFilenames.Selected(ifilename) = True
    Set objDocument = Documents.Open(Filename:=LstFilenames.List(ifilename), Visible:=False, ReadOnly:=True)
    Documents(objDocument).Activate
    
    
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
            SigPageString = Selection.Text
            
            ' process and clean up snippet and extract properties
            
            ' trim excess spaces around hashtags
            Do While SigPageString <> Replace(SigPageString, " ##", "##")
                SigPageString = Replace(SigPageString, " ##", "##")
            Loop
            Do While SigPageString <> Replace(SigPageString, "## ", "##")
                SigPageString = Replace(SigPageString, "## ", "##")
            Loop
            
            ' extract properties
            If Len(SigPageString) - InStrRev(SigPageString, "]") = 2 Then
                PropStr = Mid(SigPageString, InStrRev(SigPageString, "["), (InStrRev(SigPageString, "]") - InStrRev(SigPageString, "[")) + 1)
                
                SigPageString = Replace(SigPageString, PropStr, "")
                Do While SigPageString <> Replace(SigPageString, " ##", "##")
                    SigPageString = Replace(SigPageString, " ##", "##")
                Loop
                
                ' Trim Properties list of brackets, populate SigPageProperties array and trim any spaces around each property
                PropStr = UCase(Replace(Mid(PropStr, 2, Len(PropStr) - 2), " ", ""))
                ' split all properties into the properties array
                SigPageProperties = Split(PropStr, ",")
                ' split up property name and property value and check if the property means anything
                For i = LBound(SigPageProperties) To UBound(SigPageProperties)
                    PropName = Left(SigPageProperties(i), InStr(SigPageProperties(i), "=") - 1)
                    PropValue = Right(SigPageProperties(i), Len(SigPageProperties(i)) - InStr(SigPageProperties(i), "="))
                    Select Case PropName
                        Case "LIMIT"
                            PageLimit = Int(PropValue)
                        ' add future signature page properties here
                        'Case "INSERT PROPERTY NAME HERE"
                    End Select
                Next i
            End If
            
            ' get party name for current signature page
            SigPagePartyName = Trim(Mid(SigPageString, 18, Len(SigPageString) - 19))
            ' get name of document being procesed
            DocumentName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
            
            For DupeCount = 1 To Int(CmbDuplicateCount.Text)
                ' generate unique signature page filename using page number, copy number and party name
                SigPageFilename = OutputFolder & SigPagePartyName & " Sig Page - " & DocumentName & " (Page " & Str(Selection.Information(wdActiveEndPageNumber)) & " Copy " & Str(DupeCount) & ").pdf"
                
                ' add sig page name to array of all sig pages for reporting and combining
                TrackSigPages SigPagePartyName, SigPageFilename, DocumentName
                
                ' export each sig page as a PDF
                ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                SigPageFilename, ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
                wdExportCurrentPage, Item:=wdExportDocumentContent, _
                IncludeDocProps:=False, KeepIRM:=False, CreateBookmarks:= _
                wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=False, UseISO19005_1:=False
                
                ' count number of sig pages generated
                SigPageCount = SigPageCount + 1
                
                ' check if the current signature page has any limits and if so exit export loop
                If DupeCount = PageLimit Then
                    Debug.Print "sig page count limited"
                    Exit For
                End If
            Next
          End If
        Loop Until Not bResult
      End With
      

    objDocument.Close SaveChanges:=wdDoNotSaveChanges
    Next
    Debug.Print "The following signature pages were generated for each party:"
    For intI = (LBound(SigPageArray, 2) + 1) To (UBound(SigPageArray, 2))
        Debug.Print "Party Name: " & SigPageArray(1, intI)
        Dim DocNames
        If InStr(SigPageArray(3, intI), "/") <> 0 Then
            DocNames = Split(SigPageArray(3, intI), "/")
            For intX = 0 To (UBound(DocNames))
                Debug.Print "  *", DocNames(intX)
            Next intX
        Else
            Debug.Print "  *  " & SigPageArray(3, intI)
        End If
     Next intI
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)

    'Notify user in seconds
    MsgBox Str(SigPageCount) & " signature pages generated in " & SecondsElapsed & " seconds", vbInformation

End Sub

Private Sub UserForm_Click()


End Sub

Private Sub UserForm_Initialize()
For i = 1 To 100
CmbDuplicateCount.AddItem Trim(Str(i))
Next
getdownloadspath = Environ("USERPROFILE") & "\Downloads"
TxtOutputFolder.Text = getdownloadspath
End Sub
