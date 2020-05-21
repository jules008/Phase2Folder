VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPrintObs 
   Caption         =   "Print Utility"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   OleObjectBlob   =   "FrmPrintObs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPrintObs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Option Explicit
'
'Private Sub BtnRemoveAll_Click()
'
'End Sub
'
'Private Sub ChkPrintAll_Click()
'    Dim Item As Integer
'
'    If Me.ChkPrintAll.Value = True Then
'        Me.TxtObservation.Enabled = False
'        Me.LstPrintJobs.Enabled = False
'        Me.LstPrintJobs.BackColor = RGB(242, 242, 242)
'        Me.CmdAdd.Enabled = False
'
'        'clear all jobs
'        Me.LstPrintJobs.Clear
'
'        'add all observations to print job
'        For Item = 0 To Me.TxtObservation.ListCount - 1
'            With Me.LstPrintJobs
'                .AddItem
'                .List(.ListCount - 1, 0) = Me.TxtObservation.List(Item)
'                .List(.ListCount - 1, 1) = Me.TxtObsColour
'                .List(.ListCount - 1, 2) = Me.TxtObsCopies
'            End With
'
'        Next
'    Else
'        Me.TxtObservation.Enabled = True
'        Me.LstPrintJobs.Enabled = True
'        Me.LstPrintJobs.BackColor = RGB(255, 255, 255)
'        Me.CmdAdd.Enabled = True
'    End If
'
'End Sub
'
'Private Sub CmdAdd_Click()
'    With Me.LstPrintJobs
'        .AddItem
'        .List(.ListCount - 1, 0) = Me.TxtObservation
'        .List(.ListCount - 1, 1) = Me.TxtObsColour
'        .List(.ListCount - 1, 2) = Me.TxtObsCopies
'    End With
'
'    If Me.ChkPrintAll = False Then Me.TxtObservation.SetFocus
'
'End Sub
'
'Private Sub CmdCancel_Click()
'    Me.Hide
'End Sub
'
'Private Sub CmdPrint_Click()
'
'    Dim Observation As clsobservation
'    Dim Observations As ClsObservations
'    Dim ObsBook As Workbook
'    Dim PrintBook As Workbook
'    Dim ObsSheet As Worksheet
'    Dim FilePath As String
'    Dim DefaultPrinter As String
'    Dim i As Integer
'
'    Set PrintBook = ActiveWorkbook
'    Set Observations = New ClsObservations
'
'    Application.ActivePrinter = Excel.Range("DefaultPrinter")
'
'    'create observation print jobs
'    For i = 0 To Me.LstPrintJobs.ListCount - 1
'
'        'create new object
'        Set Observation = New clsobservation
'
'        'get print details
'        Observation.Name = Me.LstPrintJobs.List(i, 0)
'        Observation.Colour = Me.LstPrintJobs.List(i, 1)
'        Observation.Copies = Me.LstPrintJobs.List(i, 2)
'
'        'save to collection
'        Observations.Add Observation
'
'    Next
'
'    'process each observation
'    For i = 1 To Observations.Count
'
'        'take copy of Obs
'        Set Observation = Observations.Item(i)
'
'        FilePath = Observation.FileLocation
'
'        'open file
'        Set ObsBook = Workbooks.Open(FilePath)
'        Set ObsSheet = ObsBook.Worksheets(1)
'
'        'take a copy of the observation sheet
'        ObsSheet.Copy before:=PrintBook.Sheets(2)
'
'        'close source workbook
'        ObsBook.Close
'
'        'rename sheet so that it can be identified for formatting
'        PrintBook.Sheets(2).Name = "Obs"
'        PrintBook.Sheets("Obs").Range("A1").Select
'
'        'format sheet
'        Observation.Format Observation.Colour
'
'        'print sheet
'        Observation.PrintObs Observation.Copies
'
'        'delete sheet
'        Application.DisplayAlerts = False
'        PrintBook.Sheets("Obs").Delete
'        Application.DisplayAlerts = True
'    Next
'
'    'reset to defaut printer
'    Application.ActivePrinter = Excel.Range("DefaultPrinter")
'    Me.Hide
'End Sub
'
'Private Sub CmdRemove_Click()
'    Dim i As Integer
'
'    Do While i < Me.LstPrintJobs.ListCount
'        If Me.LstPrintJobs.Selected(i) Then
'            Me.LstPrintJobs.RemoveItem (i)
'        Else
'            i = i + 1
'        End If
'    Loop
'End Sub
'
'Private Sub CmdRemoveAll_Click()
'    Me.LstPrintJobs.Clear
'End Sub
'
'Private Sub UserForm_Initialize()
'    Dim RngObs As Range
'
'    For Each RngObs In ShtAssessors.Range("obslist")
'
'        'observation list
'        Me.TxtObservation.AddItem RngObs.Value
'    Next
'
'    'colour list
'    Me.TxtObsColour.AddItem "Green"
'    Me.TxtObsColour.AddItem "Grey"
'    Me.TxtObsColour.AddItem "Lilac"
'    Me.TxtObsColour.AddItem "Blue"
'    Me.TxtObsColour.AddItem "Amber"
'    Me.TxtObsColour.AddItem "Red"
'    Me.TxtObsColour.AddItem "Black"
'
'    'presets
'    Me.TxtObsColour.Value = "Green"
'    Me.TxtObsCopies.Value = 1
'    Me.ChkPrintAll.Value = False
'    Me.TxtObservation.Enabled = True
'    Me.LstPrintJobs.Enabled = True
'
'End Sub
