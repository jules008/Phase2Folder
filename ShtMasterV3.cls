VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMasterV3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Const WorkingDir = "G:\Departments\Organisational Development\Workforce Development\Phase 2 Development Folder\1. Master Spreadsheets\Working Folder\"
Const MasterSheetLocation = "G:\Departments\Organisational Development\Workforce Development\Phase 2 Development Folder\1. Master Spreadsheets\"
    
    Dim FSO As New FileSystemObject
    Dim ProgressSheet As Worksheet
    Dim ProgressBook As Workbook
    Dim MasterBook As Workbook
    Dim MasterSheetV3 As Worksheet
    Dim MasterSheetV4 As Worksheet
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
    Set MasterSheetV3 = MasterBook.Worksheets("Master V3")
    Set MasterSheetV4 = MasterBook.Worksheets("Master V4")
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
        MasterSheetV3.Range("A:CL").AutoFilter Field:=4, Criteria1:=CandidateNo
        Set RsltRng = MasterSheetV3.Range("A3:CL300").SpecialCells(xlCellTypeVisible)
        CandidateRow = RsltRng.Row
        
        'check correct record
        MasterSheetV3.Activate
        MasterSheetV3.Range("A" & CandidateRow).Select
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
    MasterSheetV3.Activate
    MasterSheetV3.Range("f" & CandidateRow & ":bx" & CandidateRow).Copy
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
        MasterSheetV3.Activate
        MasterSheetV3.Range("f" & CandidateRow).PasteSpecial
        
        
        'get overall progress
        ProgressSheet.Activate
        Progress = ProgressSheet.Range("w51") * 100
        
        Set ProgressArea = MasterSheetV3.Range("by" & CandidateRow & ":cl" & CandidateRow)
        'ProgressArea.Select
        
        For Each MyCell In ProgressArea
            If MyCell = "" Then Exit For
            Debug.Print MyCell.Value
        Next MyCell
        
        MyCell.Value = Progress
        
        'add comment
        MyCell.AddComment (Format(Now, "mmm yy"))
        
        'add date
        MasterSheetV3.Range("cm" & CandidateRow).Value = Now()
        
        
        'check entry
        MasterSheetV3.Activate
        MasterSheetV3.Range("cm1").Select
        ActiveWindow.Zoom = 70
        
        ContinueYN = True
        
        'remove filter
        With MasterSheetV3
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
      
    'Module 1
    'Diary
    x = ProgressSheet.Range("y18")
    x = x * 100
    With UpdaterSheet.Range("A2")
        .Value = x
        .NumberFormat = "0"
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = True
    End With
    
    
    'Observations
    ProgressSheet.Range("c17:i17").Copy
    UpdaterSheet.Range("B2:H2").PasteSpecial
    
    'exercise
    ProgressSheet.Range("s17").Copy
    UpdaterSheet.Range("i2").PasteSpecial
    
    
    'Module 2
    'Diary

    x = ProgressSheet.Range("y23")
    x = x * 100
    With UpdaterSheet.Range("j2")
        .Value = x
        .NumberFormat = "0"
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = True
    End With
    
    'Observations
    ProgressSheet.Range("c22:j22").Copy
    UpdaterSheet.Range("k2:R2").PasteSpecial
       
    'exercise
    ProgressSheet.Range("s22").Copy
    UpdaterSheet.Range("s2").PasteSpecial
 
    'Module 3
    'Diary
    x = ProgressSheet.Range("y28")
    x = x * 100
    With UpdaterSheet.Range("t2")
        .Value = x
        .NumberFormat = "0"
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = True
    End With
    
    'Observations
    ProgressSheet.Range("c27:s27").Copy
    UpdaterSheet.Range("u2:ak2").PasteSpecial
       
    'Module 4
    'Diary
    x = ProgressSheet.Range("y33")
    x = x * 100
    With UpdaterSheet.Range("Al2")
        .Value = x
        .NumberFormat = "0"
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = True
    End With
    
    'Observations
    ProgressSheet.Range("c32:h32").Copy
    UpdaterSheet.Range("am2:ar2").PasteSpecial
       
    'exercise
    ProgressSheet.Range("s32").Copy
    UpdaterSheet.Range("as2").PasteSpecial
    
     'Module 5
    'Diary
    x = ProgressSheet.Range("y38")
    x = x * 100
    With UpdaterSheet.Range("At2")
        .Value = x
        .NumberFormat = "0"
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = True
    End With
    
    'Observations
    ProgressSheet.Range("c37:d37").Copy
    UpdaterSheet.Range("au2:av2").PasteSpecial
       
    'snatch rescue exercise
    ProgressSheet.Range("s37").Copy
    UpdaterSheet.Range("aw2").PasteSpecial

    'Observations
    ProgressSheet.Range("e37:f37").Copy
    UpdaterSheet.Range("ax2:ay2").PasteSpecial

    'decontam exercise
    ProgressSheet.Range("r37").Copy
    UpdaterSheet.Range("az2").PasteSpecial
    
    'Module 6
    'Diary
    x = ProgressSheet.Range("y43")
    x = x * 100
    With UpdaterSheet.Range("bA2")
        .Value = x
        .NumberFormat = "0"
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = True
    End With
    
    'Observations
    ProgressSheet.Range("c42:d42").Copy
    UpdaterSheet.Range("bb2:bc2").PasteSpecial
       
    'exercise
    ProgressSheet.Range("s42").Copy
    UpdaterSheet.Range("bd2").PasteSpecial
    UpdaterSheet.Range("be2").Interior.Color = RGB(255, 255, 153)

    'Module 7
    
    'Observations
    ProgressSheet.Range("c47:j47").Copy
    UpdaterSheet.Range("bf2:bm2").PasteSpecial
    
    'test for new V6 sheet
    If ProgressSheet.Range("S1").Value = "Version 6" Then
    
        'exercise
        ProgressSheet.Range("K47").Copy
        UpdaterSheet.Range("bo2").PasteSpecial
        
        'test
        ProgressSheet.Range("L47").Copy
        UpdaterSheet.Range("bn2").PasteSpecial
        
    Else
        'exercise
        ProgressSheet.Range("r47").Copy
        UpdaterSheet.Range("bo2").PasteSpecial
        
        'test
        ProgressSheet.Range("s47").Copy
        UpdaterSheet.Range("bn2").PasteSpecial
    End If
    
    'Module 8
    
    'Observations
    ProgressSheet.Range("n47:p47").Copy
    UpdaterSheet.Range("bq2:bs2").PasteSpecial
    
    UpdaterSheet.Activate
    

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
    
    FilePath = "\\lcc-fs-010\lincoln-hq\FB-Central\Phase 2 FF Development\"
    Set Fldrs = FSO.GetFolder(FilePath).SubFolders
    
    'build filename
    FileName = "Phase 2 Progress Master Spreadsheet V3.xlsx"
    
    'save copy
    Application.DisplayAlerts = False
    
    'save V3 copy on t drive
    MasterSheetV3.Copy
    ActiveWorkbook.SaveAs FilePath & "\" & " " & FileName
    ActiveWorkbook.Close
    
    'build filename
    FileName = "Phase 2 Progress Master Spreadsheet V4.xlsx"
    
    'save copy
    Application.DisplayAlerts = False
    
    'save V4 copy on t drive
    MasterSheetV4.Copy
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

