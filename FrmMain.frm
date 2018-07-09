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
Dim TempFolder As String

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
Sub DeleteItem(Arr As Variant, index As Long)
  Dim i As Long
  For i = index To UBound(Arr) - 1
    Arr(i) = Arr(i + 1)
  Next
  Arr(UBound(Arr)) = ""
End Sub


Sub MergePDFs2()
    Dim i As Integer
    Dim intX
    Dim PartySigPageCount As Integer
    Dim AcroApp As Acrobat.CAcroApp

    Dim Part1Document As Acrobat.CAcroPDDoc
    Dim Part2Document As Acrobat.CAcroPDDoc

    Dim numPages As Integer

    Set AcroApp = CreateObject("AcroExch.App")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Debug.Print "Merging PDFs"
    
    ' Generate list of individual PDF signature pages
    'ReDim FilenameArray(UBound(SigPageArray, 2))
    'For intI = LBound(FilenameArray) To UBound(FilenameArray)
    '    FilenameArray(intI) = SigPageArray(2, intI)
    'Next
    ReDim FilenameArray(0)
    x = 0
    For intI = LBound(SigPageArray, 2) To UBound(SigPageArray, 2)
        If SigPageArray(2, intI) <> "" Then
            FilenameArray(x) = SigPageArray(2, intI)
            x = x + 1
            ReDim Preserve FilenameArray(x)
        End If
    Next
    Do While FilenameArray(UBound(FilenameArray)) = ""
        ReDim Preserve FilenameArray(UBound(FilenameArray) - 1)
    Loop
    ' Sort signature pages alphabetically
    QuickSort FilenameArray
   
    ' Create PDF objects
    Set Part1Document = CreateObject("AcroExch.PDDoc")
    Set Part2Document = CreateObject("AcroExch.PDDoc")
    
    
    i = 0
    PartySigPageCount = 1
    
    ' delete any
    
    For intX = LBound(FilenameArray) To UBound(FilenameArray)
        If FilenameArray(intX) = "" Then Exit For
    
        ' tee up first file in array
        If intX = LBound(FilenameArray) Then
            CurrentParty = Replace(FilenameArray(intX), TempFolder, "")
            CurrentParty = Left(CurrentParty, InStr(CurrentParty, " Sig Page") - 1)
            Debug.Print "Combining the following " & CurrentParty & " sig pages with: " & FilenameArray(intX)
            Part1Document.Open (FilenameArray(intX))
            
        Else      ' if not first file in array, test current file against current party
            CheckParty = Replace(FilenameArray(intX), TempFolder, "")
            CheckParty = Left(CheckParty, InStr(CheckParty, " Sig Page") - 1)
            If CheckParty = CurrentParty Then
                Debug.Print "  *  " & FilenameArray(intX)
                Part2Document.Open (FilenameArray(intX))
                numPages = Part1Document.GetNumPages()
                If Part1Document.InsertPages(numPages - 1, Part2Document, 0, Part2Document.GetNumPages(), True) = False Then
                    MsgBox "Cannot merge: " & FilenameArray(intX)
                End If
                PartySigPageCount = PartySigPageCount + 1
                Part2Document.Close

            Else
                ' check if any additional sig pages have been added to the current sig page packet. If so, save the new packet
                If PartySigPageCount > 1 Then
                If Part1Document.Save(PDSaveFull, TxtOutputFolder.Text & "Signature Pages for " & CurrentParty & ".pdf") = False Then
                    MsgBox "Cannot save signature pages for " & CurrentParty
                End If
                
                Else ' the current party only has 1 sig page, so we should copy and rename the file accordingly
                    Debug.Print " !!! " & CurrentParty & " only has 1 sig page.  copying and renaming file."
                    ' First parameter: original location\file
                    ' Second parameter: new location\file
                    objFSO.CopyFile FilenameArray(intX), TxtOutputFolder.Text & "Signature Pages for " & CurrentParty & ".pdf"

                    Part1Document.Close
                    Part2Document.Close
                    Part1Document.Open (FilenameArray(intX))
                End If
                
                ' different party name, so update currentparty
                CurrentParty = CheckParty
                Part1Document.Close
                Part2Document.Close
                Part1Document.Open (FilenameArray(intX))
                Debug.Print "Combining the following " & CurrentParty & " sig pages with: " & FilenameArray(intX)
                ' reset page count for the next party
                PartySigPageCount = 1
            End If
        End If
        
        If intX = UBound(FilenameArray) And PartySigPageCount = 1 Then
            Debug.Print " !!! " & CurrentParty & " only has 1 sig page.  copying and renaming file."
            ' add 'copy/save as' language here also because the last file in the list is for a party that only has 1 sig page.
            ' First parameter: original location\file
            ' Second parameter: new location\file
            objFSO.CopyFile FilenameArray(intX), TxtOutputFolder.Text & "Signature Pages for " & CurrentParty & ".pdf"


        End If
    
    Next
    AcroApp.Exit
    Set AcroApp = Nothing
    Set Part1Document = Nothing
    Set Part2Document = Nothing
    Debug.Print "DONE"

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
        ReDim FilenameArray(0)
        x = 0
        For intI = LBound(SigPageArray, 2) To UBound(SigPageArray, 2)
            If SigPageArray(2, intI) <> "" Then
                FilenameArray(x) = SigPageArray(2, intI)
                x = x + 1
                ReDim Preserve FilenameArray(x)
            End If
        Next
        QuickSort FilenameArray
        For intI = LBound(FilenameArray) To UBound(FilenameArray)
            If intI = LBound(FilenameArray) Then
                FrmReport.TxtReport.Text = Replace(FilenameArray(intI), TempFolder, "")
            ElseIf intI = UBound(FilenameArray) Then
                FrmReport.TxtReport.Text = FrmReport.TxtReport.Text & Replace(FilenameArray(intI), TempFolder, "")
            Else
                FrmReport.TxtReport.Text = FrmReport.TxtReport.Text & Replace(FilenameArray(intI), TempFolder, "") & vbCrLf
            End If
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
  Dim CurrentViewAllStatus As Boolean
  Dim StartTime As Double
  Dim SecondsElapsed As Double
  
    If LstFilenames.ListCount = 0 Then Exit Sub
    ' add trailing slash for consistency
    If Right(TxtOutputFolder.Text, 1) <> "\" Then TxtOutputFolder.Text = TxtOutputFolder.Text + "\"

    ' make temp folder using seconds since midnight as random number
    temptime = Trim(Str(Timer))
    TempFolder = TxtOutputFolder.Text & Replace((temptime), ".", "") & "\"
    MkDir TempFolder

    'Remember time when macro starts
    StartTime = Timer
    TxtReport.Text = ""
    


  CurrentViewAllStatus = ActiveWindow.ActivePane.View.ShowAll
  ActiveWindow.ActivePane.View.ShowAll = True

  Application.Browser.Target = wdBrowsePage
  'Set SigPagesDoc = Documents.Add
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
    Debug.Print "current doc: " & Selection.Document.Name

      Selection.WholeStory   'Select entire document
      Selection.HomeKey Unit:=wdStory
      With Selection.Find
        .ClearFormatting
        .MatchWholeWord = False
        .MatchCase = False
        .Font.Hidden = True
        .MatchWildcards = True
'        .Text = "##Signature Page-*##"
        Selection.Find.Font.Hidden = True
        .Replacement.Text = ""
        Do 'While .Execute
          bResult = .Execute(FindText:="##Signature Page-*##")
          If bResult Then
            Debug.Print "sig page found"
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
                SigPageFilename = TempFolder & SigPagePartyName & " Sig Page - " & DocumentName & " (Copy " & Str(DupeCount) & " Page " & Str(Selection.Information(wdActiveEndPageNumber)) & ").pdf"
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
    
    ' generate report of signature pages for text box
    GenerateReport
    ' Merge individual PDF signature pages into a single packet
    MergePDFs2
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    'Notify user in seconds
    MsgBox Str(SigPageCount) & " signature pages generated in " & SecondsElapsed & " seconds", vbInformation
        
    ' revert paragraph view button to original state
    ActiveWindow.ActivePane.View.ShowAll = CurrentViewAllStatus

    ' delete temp folder
    On Error Resume Next
    Kill TempFolder & "*.*"    ' delete all files in the folder
    RmDir TempFolder ' delete folder

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
