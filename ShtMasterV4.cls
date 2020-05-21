VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMasterV4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'    Const WORKING_DIR = "G:\Development Areas\Phase 2 Folder\WorkingFolder\"
'    Const MASTER_SHEET = "G:\Development Areas\Phase 2 Folder\"
'    Const SAVE_LOCATION = "G:\Development Areas\Phase 2 Folder\T_Drive\"

    Const WORKING_DIR = "G:\Departments\Organisational Development\Workforce Development\Phase 2 Development Folder\1. Master Spreadsheets\Working Folder\"
    Const MASTER_SHEET = "G:\Departments\Organisational Development\Workforce Development\Phase 2 Development Folder\1. Master Spreadsheets\"
    Const SAVE_LOCATION = "\\lcc-fs-010\lincoln-hq\FB-Central\Phase 2 FF Development\"
   
    Dim FSO As New FileSystemObject
    Dim ProgressSheet As Worksheet
    Dim ProgressBook As Workbook
    Dim MasterBook As Workbook
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
    Dim i As Variant
    Dim x As String
    
Public Sub Handler()
    
    Dim RsltRng As Range
    Dim DlgOpen As FileDialog
    
    On Error GoTo FatalErrHandler
    
    Set MasterBook = ActiveWorkbook
    Set MasterSheetV4 = MasterBook.Worksheets("Master V4")
    Set UpdaterSheet = MasterBook.Worksheets("Updater")
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
       
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .InitialFileName = WORKING_DIR
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
        MasterSheetV4.Range("A:CL").AutoFilter Field:=4, Criteria1:=CandidateNo
        Set RsltRng = MasterSheetV4.Range("A3:CL300").SpecialCells(xlCellTypeVisible)
        CandidateRow = RsltRng.Row
        
        'check correct record
        MasterSheetV4.Activate
        MasterSheetV4.Range("A" & CandidateRow).Select
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
    MasterSheetV4.Activate
    MasterSheetV4.Range("F" & CandidateRow & ":CK" & CandidateRow).Copy
    UpdaterSheet.Activate
    UpdaterSheet.Range("A4").PasteSpecial
    UpdaterSheet.Range("A1").Select
    
    'confirm change ok
    UpdaterSheet.Visible = xlSheetVisible
    UpdaterSheet.Activate
    ActiveWindow.Zoom = 80
    ContinueYN = MsgBox("Are the changes correct?", vbYesNo, "Confirm Changes")
    UpdaterSheet.Visible = xlSheetHidden

    If ContinueYN = 6 Then
    
        'copy data
        UpdaterSheet.Range("A2:CF2").Copy
    
        'paste data into master
        MasterSheetV4.Activate
        MasterSheetV4.Range("F" & CandidateRow).PasteSpecial
        
        
        'get overall progress
        ProgressSheet.Activate
        Progress = ProgressSheet.Range("w51") * 100
        
        Set ProgressArea = MasterSheetV4.Range("CL" & CandidateRow & ":CY" & CandidateRow)
        
        For Each MyCell In ProgressArea
            If MyCell = "" Then Exit For
            Debug.Print MyCell.Value
        Next MyCell
        
        MyCell.Value = Progress
        
        'add comment
        MasterSheetV4.Activate
        MyCell.AddComment (Format(Now, "mmm yy"))
        
        'add date
        MasterSheetV4.Range("CZ" & CandidateRow).Value = Now()
        
        
        'check entry
        MasterSheetV4.Activate
        MasterSheetV4.Range("A1").Select
        ActiveWindow.Zoom = 70
        
        ContinueYN = True
        
        'remove filter
        With MasterSheetV4
            .Range("A:CQ").AutoFilter Field:=4
            .Range("g3").Select
        End With
        
        ActiveWindow.Zoom = 100
        
        If ContinueYN = True Then
        
            SaveFile

     
        End If
        
    End If
End Sub


Public Sub TransferData()
    Dim TechSht As Boolean
    Dim RngSel As Range
    
    If Right(ProgressSheet.Range("S1"), 1) = "t" Then TechSht = True Else TechSht = False
    
    'BA Theory
    
    Set RngSel = ProgressSheet.Range("Y16")
    
    RngSel.Copy
    With UpdaterSheet.Range("A2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    'BA Module
    ProgressSheet.Range("C17:N17").Copy
    UpdaterSheet.Range("B2:M2").PasteSpecial
        
    'RTC Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y19") Else Set RngSel = ProgressSheet.Range("Y20")
    
    RngSel.Copy
    With UpdaterSheet.Range("N2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    'Extrication and resuce module
    ProgressSheet.Range("C22:N22").Copy
    UpdaterSheet.Range("O2:Z2").PasteSpecial
       
    'Pumps Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y22") Else Set RngSel = ProgressSheet.Range("Y24")
    
    RngSel.Copy
    With UpdaterSheet.Range("AA2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    'Pumps module
    ProgressSheet.Range("C27:O27").Copy
    UpdaterSheet.Range("AB2:AN2").PasteSpecial
       
    'Water Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y25") Else Set RngSel = ProgressSheet.Range("Y28")
    
    RngSel.Copy
    With UpdaterSheet.Range("AO2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    'Water Module
    ProgressSheet.Range("C32:F32").Copy
    UpdaterSheet.Range("AP2:AS2").PasteSpecial
     
    'WAH Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y28") Else Set RngSel = ProgressSheet.Range("Y32")
    
    RngSel.Copy
    With UpdaterSheet.Range("AT2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    ' Working at Height Module
    ProgressSheet.Range("C37:Q37").Copy
    UpdaterSheet.Range("AU2:BI2").PasteSpecial
       
    'Haz Mats Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y31") Else Set RngSel = ProgressSheet.Range("Y36")
    
    RngSel.Copy
    With UpdaterSheet.Range("BJ2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    'Haz Mats Module
    ProgressSheet.Range("C42:H42").Copy
    UpdaterSheet.Range("BK2:BP2").PasteSpecial

    'Comms Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y34") Else Set RngSel = ProgressSheet.Range("Y40")
    
    RngSel.Copy
    With UpdaterSheet.Range("BQ2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    'Communications Module
    ProgressSheet.Range("C47:G47").Copy
    UpdaterSheet.Range("BR2:BV2").PasteSpecial

    'Cas Care Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y37") Else Set RngSel = ProgressSheet.Range("Y44")
    
    RngSel.Copy
    With UpdaterSheet.Range("BW2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    'Casualty Care Module
    ProgressSheet.Range("J47").Copy
    UpdaterSheet.Range("BX2").PasteSpecial
    
    'Misc Theory
    
    If TechSht Then Set RngSel = ProgressSheet.Range("Y40")
    
    RngSel.Copy
    With UpdaterSheet.Range("BY2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        If ProgressSheet.Range("P45") <> 1 Then .Interior.Color = RGB(183, 222, 232)
        .Font.Size = 10
    End With
    
    'Misc Module
    ProgressSheet.Range("N47:S47").Copy
    UpdaterSheet.Range("BZ2:CE2").PasteSpecial
    
    'MOK Theory
    
    Set RngSel = ProgressSheet.Range("Y47")
    
    RngSel.Copy
    With UpdaterSheet.Range("CF2")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
        .Font.Size = 10
    End With
    
    Set RngSel = Nothing

End Sub

Private Sub CmdUpdPrg_Click()
    Handler
End Sub

Public Sub SaveFile()
    Dim Fldrs As Folders
    Dim Fld As Folder
    Dim Fls As Files
    Dim SAVE_LOCATION As String
    
    If Application.UserName <> "Julian Turner" Then
        MasterBook.Save
        
        SAVE_LOCATION = "\\lcc-fs-010\lincoln-hq\FB-Central\Phase 2 FF Development\"
        Set Fldrs = FSO.GetFolder(SAVE_LOCATION).SubFolders
        
        'build filename
        FileName = "Phase 2 Progress Master Spreadsheet V4.xlsx"
        
        'save copy
        Application.DisplayAlerts = False
                
        'build filename
        FileName = "Phase 2 Progress Master Spreadsheet V4.xlsx"
        
        'save copy
        Application.DisplayAlerts = False
        
        'save V4 copy on t drive
        MasterSheetV4.Copy
        ActiveWorkbook.SaveAs SAVE_LOCATION & "\" & " " & FileName
        ActiveWorkbook.Close
        
        'save archive MasterSheetV4
        
        'MasterBook.SaveAs FileName:=MasterSheetV4Location & "\" & FileName, FileFormat:=xlOpenXMLWorkbook
        
        
        'save master copy
        'MasterBook.SaveAs FileName:=MASTER_SHEET & "Phase 2 Progress Master Spreadsheet.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
        
        Application.DisplayAlerts = True
        
        'look for station folder
        For Each Fld In Fldrs
            Debug.Print Left(Fld.Name, 2)
            
            If Left(Fld.Name, 2) = Station Then Exit For
            
        Next
        
        SAVE_LOCATION = SAVE_LOCATION & Fld.Name
    
        'save progress sheet
        Application.DisplayAlerts = False
        
        SAVE_LOCATION = SAVE_LOCATION & "\" & Format(Now, "yyyy-mm-dd") & " " & CandidateName & " " & CandidateNo & " v4" & FileExt
        
        ProgressBook.SaveAs (SAVE_LOCATION)
        ProgressBook.Close
        Application.DisplayAlerts = True
        
         With Application.FileDialog(msoFileDialogOpen)
            .InitialFileName = SAVE_LOCATION
            .AllowMultiSelect = True
            .Filters.Add "All Files (*.*)", "*.*"
            .Title = "Move old file to archive"
            .Show
        End With
    End If

End Sub

