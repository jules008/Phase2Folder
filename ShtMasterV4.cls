VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMasterV4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Const WorkingDir = "\\imp.lincolnshire.gov.uk\DavWWWRoot\contentserverdav\nodes\22151336\Working Folder\"
Const MasterSheetLocation = "\\imp.lincolnshire.gov.uk\DavWWWRoot\contentserverdav\nodes\21592528\Workforce Development\Phase 2 Development Folder\STATS\"
    
    Dim FSO As New FileSystemObject
    Dim ProgressSheet As Worksheet
    Dim ProgressBook As Workbook
    Dim MasterBook As Workbook
    Dim MasterSheet As Worksheet
    Dim UpdaterSheet As Worksheet
    Dim ProgressArea As Range
    Dim FileExt As String
    Dim Station As String
    Dim CandidateNo As String
    Dim CandidateName As String
    Dim Progress As Integer
    Dim ContinueYN As String
    Dim CandidateRow As String
    Dim NewBook As Workbook
    Dim FileName As String
    Dim Workingfolder As Files
    Dim WBopen As Boolean

Public Sub Handler()
    
    Dim RsltRng As Range
    Dim DlgOpen As FileDialog
    
    On Error GoTo FatalErrHandler
    
    Set MasterBook = ActiveWorkbook
    Set MasterSheet = MasterBook.Worksheets("Master V4")
    Set UpdaterSheet = MasterBook.Worksheets("Updater")
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
       
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .InitialFileName = WorkingDir
        .Filters.Clear
        .Filters.Add "Excel Files (*.xls; *.xlsx; *.xlsm)", "*.xls, *.xlsx, *.xlsm"
        .AllowMultiSelect = True
        .Title = "Open Progress Files"
        .Show
    End With
    
    'exit if no files selected
    If DlgOpen.SelectedItems.Count = 0 Then
        MsgBox "There are no files selected", vbOKOnly, "No Files"
        Exit Sub
    End If
           
     'loop and open the files
    For Each i In DlgOpen.SelectedItems
     
        'open progress sheets
        Workbooks.Open i
        
        'determine file ext
        x = Right(i, 1)
        
        Select Case x
            Case "x"
                FileExt = ".xlsx"
            Case "s"
                FileExt = ".xls"
            Case "m"
                FileExt = ".xlsm"
        End Select
        
        'set first sheet to working file
        Set ProgressBook = ActiveWorkbook
        Set ProgressSheet = ProgressBook.Worksheets(1)
    
        'get candidate number
        ProgressSheet.Activate
        CandidateNo = ProgressSheet.Range("P4")
        CandidateName = ProgressSheet.Range("i4")
        Station = ProgressSheet.Range("s4")
        
        
        'add leading zero if needed
        If Len(Station) = 1 Then
            Station = "0" & Station
        End If
    
        'filter master spreadsheet for candidate
        MasterSheet.Range("A:CL").AutoFilter Field:=4, Criteria1:=CandidateNo
        Set RsltRng = MasterSheet.Range("A3:CL300").SpecialCells(xlCellTypeVisible)
        CandidateRow = RsltRng.Row
        
        'check correct record
        MasterSheet.Activate
        MasterSheet.Range("A" & CandidateRow).Select
        ContinueYN = MsgBox("Is this the correct record?", vbYesNo, "Confirm Record")
        
        If ContinueYN = 6 Then
            TransferData
            
            EnterData
            
                        
        End If
        
    Next
    
EndProgram:
Exit Sub

FatalErrHandler:
    MsgBox "Error " & Err.Number & ": " & "Something has gone wrong, application has shut down", vbOKOnly, "Error"
    Resume EndProgram
End Sub

Public Sub EnterData()
    Dim MyCell As Range
    Dim RsltRng As Range
    
    'backup current data
    MasterSheet.Activate
    MasterSheet.Range("f" & CandidateRow & ":bx" & CandidateRow).Copy
    UpdaterSheet.Activate
    UpdaterSheet.Range("A4").PasteSpecial
    UpdaterSheet.Range("a1").Select
    
    'confirm change ok
    UpdaterSheet.Visible = xlSheetVisible
    UpdaterSheet.Activate
    ContinueYN = MsgBox("Are the changes correct?", vbYesNo, "Confirm Changes")
    UpdaterSheet.Visible = xlSheetHidden

    If ContinueYN = 6 Then
    
        'copy data
        UpdaterSheet.Range("A2:bs2").Copy
    
        'paste data into master
        MasterSheet.Activate
        MasterSheet.Range("f" & CandidateRow).PasteSpecial
        
        
        'get overall progress
        ProgressSheet.Activate
        Progress = ProgressSheet.Range("w51") * 100
        
        Set ProgressArea = MasterSheet.Range("by" & CandidateRow & ":cl" & CandidateRow)
        'ProgressArea.Select
        
        For Each MyCell In ProgressArea
            If MyCell = "" Then Exit For
            Debug.Print MyCell.Value
        Next MyCell
        
        MyCell.Value = Progress
        
        'add comment
        MyCell.AddComment (Format(Now, "mmm yy"))
        
        'add date
        MasterSheet.Range("cm" & CandidateRow).Value = Now()
        
        
        'check entry
        MasterSheet.Activate
        MasterSheet.Range("cm1").Select
        ActiveWindow.Zoom = 70
        
        ContinueYN = MsgBox("Happy?", vbYesNo)
        
        'remove filter
        With MasterSheet
            .Range("A:CL").AutoFilter Field:=4
            .Range("g3").Select
        End With
        
        ActiveWindow.Zoom = 100
        
        If ContinueYN = True Then
        
            SaveFile

     
        End If
        
    End If
End Sub


Public Sub TransferData()
      
    'BA Module
    ProgressSheet.Range("C17:N17").Copy
    UpdaterSheet.Range("B2:M2").PasteSpecial
    
    'exercise
    ProgressSheet.Range("s17").Copy
    UpdaterSheet.Range("i2").PasteSpecial
    
    'Extrication and resuce module
    ProgressSheet.Range("C22:M22").Copy
    UpdaterSheet.Range("O2:Y2").PasteSpecial
       
    'Pumps module
    ProgressSheet.Range("C27:M27").Copy
    UpdaterSheet.Range("AA2:AK2").PasteSpecial
       
    'Water Module
    ProgressSheet.Range("C32:F32").Copy
    UpdaterSheet.Range("AM2:AP2").PasteSpecial
     
    ' Working at Height Module
    ProgressSheet.Range("C37:K37").Copy
    UpdaterSheet.Range("AR2:AZ2").PasteSpecial
       
    'Haz Mats Module
    ProgressSheet.Range("C42:G42").Copy
    UpdaterSheet.Range("BB2:BG2").PasteSpecial

    'Communications Module
    ProgressSheet.Range("C47:G47").Copy
    UpdaterSheet.Range("BI2:BM2").PasteSpecial

    'Casualty Care Module
    ProgressSheet.Range("J47").Copy
    UpdaterSheet.Range("BO2").PasteSpecial
    
    'Misc Module
    ProgressSheet.Range("N47:S47").Copy
    UpdaterSheet.Range("BQ:BU").PasteSpecial
    
End Sub

Private Sub CmdUpdPrg_Click()
    Handler
End Sub

Public Sub SaveFile()
    Dim Fldrs As Folders
    Dim Fld As Folder
    Dim Fls As Files
    Dim FilePath As String
    
    MasterBook.Save
    
    FilePath = "\\imp.lincolnshire.gov.uk\DavWWWRoot\contentserverdav\nodes\21592528\Phase 2 FF Development\"
    Set Fldrs = FSO.GetFolder(FilePath).SubFolders
    
    'build filename
    FileName = "Phase 2 Progress Master Spreadsheet.xlsx"
    
    'save copy
    Application.DisplayAlerts = False
    
    'save copy on t drive
    MasterSheet.Copy
    ActiveWorkbook.SaveAs FilePath & "\" & " " & FileName
    ActiveWorkbook.Close
    'save archive MasterSheet
    
    'MasterBook.SaveAs FileName:=MasterSheetLocation & "\" & FileName, FileFormat:=xlOpenXMLWorkbook
    
    
    'save master copy
    'MasterBook.SaveAs FileName:=MasterSheetLocation & "Phase 2 Progress Master Spreadsheet.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    
    Application.DisplayAlerts = True
    
    'look for station folder
    For Each Fld In Fldrs
        Debug.Print Left(Fld.Name, 2)
        
        If Left(Fld.Name, 2) = Station Then Exit For
        
    Next
    
    FilePath = FilePath & Fld.Name

    'save progress sheet
    Application.DisplayAlerts = False
    
    FilePath = FilePath & "\" & Format(Now, "yyyy-mm-dd") & " " & CandidateName & " " & CandidateNo & " v4" & FileExt
    
    ProgressBook.SaveAs (FilePath)
    ProgressBook.Close
    Application.DisplayAlerts = True
    
     With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = FilePath
        .AllowMultiSelect = True
        .Filters.Add "All Files (*.*)", "*.*"
        .Title = "Move old file to archive"
        .Show
    End With
    

End Sub

