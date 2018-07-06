VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "Signature Page Generator"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10635
   OleObjectBlob   =   "FrmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Restart:
For i = 0 To LstFilenames.ListCount
    If LstFilenames.Selected(i) = True Then
        LstFilenames.Selected(i) = False
        LstFilenames.RemoveItem (i)
        GoTo Restart:
    End If
Next
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
  Dim OutputFolder As String
  
  Application.Browser.Target = wdBrowsePage
  'Set SigPagesDoc = Documents.Add
  If Right(TxtOutputFolder.Text, 1) <> "\" Then OutputFolder = TxtOutputFolder + "\"
  For ifilename = 1 To LstFilenames.ListCount
    Set objDocument = Documents.Open(FileName:=LstFilenames.List(ifilename - 1), Visible:=False, ReadOnly:=True)
    Documents(objDocument).Activate
    
    
      Selection.WholeStory   'Select entire document
    
      With Selection.Find
    '**** Syntax of Find Execute incase you want to use some of the
    '**** other parameters.
    '         expression .Execute(FindText, MatchCase, MatchWholeWord, _
    '          MatchWildcards, MatchSoundsLike, MatchAllWordForms, _
    '          Forward, Wrap, Format, ReplaceWith, Replace, MatchKashida,_
    '          MatchDiacritics, MatchAlefHamza, MatchControl)
              
        .ClearFormatting
        .MatchWholeWord = True
        .MatchCase = False
        .Font.Hidden = True
        .MatchWildcards = True
        Do
          bResult = .Execute(FindText:="##Signature Page-*##")  '<- Your phrase here!
          If bResult Then
          'the below finds docx or whatever filename and trims it.  but if file is not saved, there is no filename to trim...
          '
           ' MsgBox "Sig Page - " & Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1) & " - " & Trim(Right(Selection.Text, Len(Selection.Text) - InStr(Selection.Text, "-"))) & " - " & Str(Selection.Information(wdActiveEndPageNumber)) & ".pdf"
            'Application.PrintOut FileName:="", Range:=wdPrintCurrentPage, Item:= _
              wdPrintDocumentWithMarkup, Copies:=1, Pages:="", PageType:= _
              wdPrintAllPages, Collate:=True, Background:=True, PrintToFile:=False, _
              PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
              PrintZoomPaperHeight:=0
            SigPageName = Trim(Mid(Selection.Text, 18, Len(Selection.Text) - 19))
            For i = 1 To Int(CmbDuplicateCount.Text)
                ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                OutputFolder & SigPageName & " Sig Page - " & Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1) & " (Page " & Str(Selection.Information(wdActiveEndPageNumber)) & " Copy " & Str(i) & ").pdf", ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
                wdExportCurrentPage, Item:=wdExportDocumentContent, _
                IncludeDocProps:=False, KeepIRM:=False, CreateBookmarks:= _
                wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=False, UseISO19005_1:=False
            Next
          End If
        Loop Until Not bResult
      End With
      

    objDocument.Close SaveChanges:=wdDoNotSaveChanges
    Next
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
