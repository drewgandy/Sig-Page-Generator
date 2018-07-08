VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "Signature Page Generator"
   ClientHeight    =   8745
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

Dim SigPageArray As Variant
Dim FilenameArray() As String
Dim xCounter As Integer
Option Compare Text

' Omit plngLeft & plngRight; they are used internally during recursion
Public Sub QuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then QuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then QuickSort pvarArray, lngFirst, plngRight
End Sub
Sub GenerateReport()
    If UBound(SigPageArray, 2) = 0 Then Exit Sub
    FrmReport.TxtReport.Text = ""
    Dim KeyList() As String
    Dim KeyBasenameList() As String
    Dim KeyFound As Boolean
    Dim SortByColumn As Integer
    Dim iKeyBasename As Integer
    'Dim FilenameArray() As String
    KeyFound = False
    SortByColumn = 3
    If FrmReport.RadioParty.Value = True Then SortByColumn = 1
    If FrmReport.RadioDocument.Value = True Then SortByColumn = 3
    
    ReDim KeyList(1)
    ReDim KeyBasenameList(1)
    
    If FrmReport.RadioSigPages.Value = False Then
        For intI = LBound(SigPageArray, 2) To UBound(SigPageArray, 2)

            For iKey = LBound(KeyList) To UBound(KeyList)
                If KeyList(iKey) = SigPageArray(SortByColumn, intI) Then KeyFound = True
            Next iKey
            
            'For iKeyBasename = LBound(KeyBasenameList) To UBound(KeyBasenameList)
            '    If KeyBasenameList(iKeyBasename) = SigPageArray(4, intI) Then KeyFound = True: Debug.Print "already listed sig page"

            'Next iKeyBasename

            If KeyFound = False Then
                FrmReport.TxtReport.Text = FrmReport.TxtReport.Text & SigPageArray(SortByColumn, intI) & vbCrLf
                For intX = LBound(SigPageArray, 2) To UBound(SigPageArray, 2)


                    If SigPageArray(SortByColumn, intX) = SigPageArray(SortByColumn, intI) Then
                        For intY = intI To intX - 1 'LBound(SigPageArray, 2) To intX - 1
                            If SigPageArray(4, intX) = SigPageArray(4, intY) Then KeyFound = False: GoTo SkipDupes:
                        Next intY
                        If SortByColumn = 1 Then
                            FrmReport.TxtReport.Text = FrmReport.TxtReport.Text & "  *  " & SigPageArray(3, intX) & vbCrLf
                        Else
                            FrmReport.TxtReport.Text = FrmReport.TxtReport.Text & "  *  " & SigPageArray(1, intX) & vbCrLf
                        End If
                    End If
SkipDupes:
                Next intX

                ReDim Preserve KeyList(UBound(KeyList) + 1)
                KeyList(UBound(KeyList)) = SigPageArray(SortByColumn, intI)
                ReDim Preserve KeyBasenameList(UBound(KeyBasenameList) + 1)
                KeyBasenameList(UBound(KeyBasenameList)) = SigPageArray(4, intI)
            Else
                KeyFound = False
            End If
        Next intI
     Else
        ReDim FilenameArray(UBound(SigPageArray, 2))
        For intI = LBound(FilenameArray) To UBound(FilenameArray)
            FilenameArray(intI) = SigPageArray(2, intI)
        Next
        QuickSort FilenameArray
        For intI = LBound(FilenameArray) To UBound(FilenameArray)
            FrmReport.TxtReport.Text = FrmReport.TxtReport.Text & Replace(FilenameArray(intI), TxtOutputFolder.Text, "") & vbCrLf
        Next
    End If

End Sub




Sub TrackSigPages(ByVal PartyName As String, ByVal Filename As String, ByVal DocName As String, ByVal BaseName As String)
    
    ReDim Preserve SigPageArray(4, UBound(SigPageArray, 2) + 1)
    SigPageArray(1, UBound(SigPageArray, 2)) = PartyName
    SigPageArray(2, UBound(SigPageArray, 2)) = Filename
    SigPageArray(3, UBound(SigPageArray, 2)) = DocName
    SigPageArray(4, UBound(SigPageArray, 2)) = BaseName
    
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
  Dim SigPageBasename As String
  Dim OutputFolder As String
  Dim SigPageProperties() As String
  Dim PropName As String
  Dim PropValue As String
  Dim PageLimit As Integer
  Dim PageRange As Integer
  Dim PropStr As String
  Dim SigPagePartyName As String
  Dim DocumentName As String
  Dim SigPageCount As Integer
  Dim PageNumOfSigPage As Integer
  
  Dim StartTime As Double
  Dim SecondsElapsed As Double

  'Remember time when macro starts
  StartTime = Timer

  
  Application.Browser.Target = wdBrowsePage
  'Set SigPagesDoc = Documents.Add
  If Right(TxtOutputFolder.Text, 1) <> "\" Then TxtOutputFolder.Text = TxtOutputFolder.Text + "\"
  ReDim SigPageArray(4, 0)
  For ifilename = 0 To LstFilenames.ListCount - 1
        LstFilenames.Selected(ifilename) = False
    Next
  For ifilename = 0 To LstFilenames.ListCount - 1
    LstFilenames.Selected(ifilename) = True
    LstFilenames.ListIndex = ifilename
    PageLimit = 0
    PageRange = 0
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
                        Case "PAGES"
                            PageRange = Int(PropValue) - 1
                        ' add future signature page properties here
                        'Case "INSERT PROPERTY NAME HERE"
                    End Select
                Next i
            End If
            
            ' get party name for current signature page
            SigPagePartyName = Trim(Mid(SigPageString, 18, Len(SigPageString) - 19))
            ' get name of document being procesed
            DocumentName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
            PageNumOfSigPage = Selection.Information(wdActiveEndPageNumber)

            For DupeCount = 1 To Int(CmbDuplicateCount.Text)
                ' generate unique signature page filename using page number, copy number and party name
                SigPageFilename = TxtOutputFolder.Text & SigPagePartyName & " Sig Page - " & DocumentName & " (Copy " & Str(DupeCount) & " Page " & Str(Selection.Information(wdActiveEndPageNumber)) & ").pdf"
                ' generate unique signature page basename using page number and party name (but not copy number). This is used avoid duplicates in the report
                SigPageBasename = SigPagePartyName & " Sig Page - " & DocumentName & " (Page " & Str(Selection.Information(wdActiveEndPageNumber)) & ").pdf"
                
                ' add sig page name to array of all sig pages for reporting and combining
                TrackSigPages SigPagePartyName, SigPageFilename, DocumentName, SigPageBasename
                
                ' export each sig page as a PDF
                ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                SigPageFilename, ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
                wdExportFromTo, From:=PageNumOfSigPage, To:=(PageNumOfSigPage + PageRange), Item:=wdExportDocumentContent, _
                IncludeDocProps:=False, KeepIRM:=False, CreateBookmarks:= _
                wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=False, UseISO19005_1:=False
                 ' export each sig page as a PDF
               ' ActiveDocument.ExportAsFixedFormat OutputFileName:= _
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
                    Exit For
                End If
            Next
          End If
        Loop Until Not bResult
      End With
      

    objDocument.Close SaveChanges:=wdDoNotSaveChanges
    Next
    
    GenerateReport
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    'Notify user in seconds
    MsgBox Str(SigPageCount) & " signature pages generated in " & SecondsElapsed & " seconds", vbInformation

    
End Sub



Private Sub RadioDocument_Click()
GenerateReport
End Sub

Private Sub RadioParty_Click()
GenerateReport
End Sub

Private Sub RadioSigPages_Click()
GenerateReport
End Sub

Private Sub UserForm_Click()


End Sub

Private Sub UserForm_Initialize()
ReDim SigPageArray(4, 0)

For i = 1 To 100
CmbDuplicateCount.AddItem Trim(Str(i))
Next
getdownloadspath = Environ("USERPROFILE") & "\Downloads"
TxtOutputFolder.Text = getdownloadspath
End Sub
