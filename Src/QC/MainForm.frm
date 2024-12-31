VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{1416D7C5-8A28-11CF-9236-444553540000}#8.0#0"; "PVXPLORE8.ocx"
Begin VB.Form MainForm 
   BackColor       =   &H80000005&
   Caption         =   "ISA DVP-2 Quality Control"
   ClientHeight    =   13965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   13965
   ScaleWidth      =   18330
   StartUpPosition =   2  'CenterScreen
   Begin PVExplorerLib.PVExplorer MainExplorer 
      Height          =   13545
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   18285
      _Version        =   524288
      LabelEdit       =   0   'False
      Indentation     =   0
      SourceChannel1  =   ""
      TargetChannel1  =   ""
      PathSeparator   =   ""
      Image1          =   "MainForm.frx":0000
      SingleExpand    =   -1  'True
      SourceChannel2  =   ""
      TargetChannel2  =   ""
      Image2          =   "MainForm.frx":037E
      Image3          =   "MainForm.frx":227C
      FileName        =   ""
      DataMember      =   ""
      DataField0      =   ""
      DataField1      =   ""
      DataField2      =   ""
      DataField3      =   ""
      DataField4      =   ""
      DataField5      =   ""
      DataField6      =   ""
      DataField7      =   ""
      DataField8      =   ""
      DataField9      =   ""
      DataField10     =   ""
      DataField11     =   ""
      DataField12     =   ""
      DataField13     =   ""
      DataField14     =   ""
      DataField15     =   ""
      DataField16     =   ""
      DataField17     =   ""
      DataField18     =   ""
      DataField19     =   ""
      PaneDisplay     =   1
      CaptionMode     =   1
      Appearance      =   1
      _ExtentX        =   32253
      _ExtentY        =   23892
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars MainToolBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   4
      Tools           =   "MainForm.frx":25FA
      ToolBars        =   "MainForm.frx":4CE7
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   13710
      Width           =   18330
      _ExtentX        =   32332
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm DensComm 
      Left            =   17730
      Top             =   13110
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2 Quality Control                                     **
'**                                                                        **
'** Module.....: MainForm                                                  **
'**                                                                        **
'** Description: This module provides the main application interface.      **
'**                                                                        **
'** History....:                                                           **
'**    10/02/03 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Private CommInputBuffer As String                       'Serial communications input buffer

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Load                                             **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    '---- Initialize form controls
    UpdatePrinterList
    CommConnect
    Me.Caption = "ISA DVP-2 Q.C. Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_QueryUnload                                      **
'**                                                                        **
'**  Description..:  This routine prompts to exit & shuts down the program.**
'**                                                                        **
'****************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo, "Exit DVP-2 Q.C. Program?") = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    DB.CloseDataBase
    Set DB = Nothing
    Unload ErrorForm
    Unload LutControlForm
    Unload OffsetControlForm
    Unload ColorControlForm
    Unload SettingsControlForm
    Unload PrinterControlForm
    Unload Me
End Sub

Private Sub Form_Resize()
    If Me.Width - 100 > 0 Then
        MainExplorer.Width = Me.Width - 100
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MainToolBar_ToolClick                                 **
'**                                                                        **
'**  Description..:  This routine handles user selection on tool bar.      **
'**                                                                        **
'****************************************************************************
Private Sub MainToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_New"
            AddPrinter
        Case "ID_Erase"
            Dim PrinterDesc As String
            If MainExplorer.SelectedNode.Level = 0 Then
                PrinterDesc = MainExplorer.SelectedNode.Text
            Else
                PrinterDesc = MainExplorer.SelectedNode.Parent.Text
            End If
            If MsgBox("Permanently Erase Printer '" & PrinterDesc & "' (" & PrinterName & ")?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Erase Printer") = vbYes Then
                If MsgBox("You are about to permanently erase printer '" & PrinterDesc & "' (" & PrinterName & ")." & vbCrLf & "Are you sure?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Erase Printer - WARNING") = vbYes Then
                    With DB.rsPrinterList
                        If .RecordCount > 0 Then .MoveFirst
                        .Find "PrinterName='" & PrinterName & "'"
                        If Not .EOF Then
                            .Delete adAffectCurrent
                            UpdatePrinterList
                        Else
                            MsgBox "ERROR!!!"
                        End If
                    End With
                End If
            End If
        Case "ID_Exit"
            Unload Me
    End Select
End Sub

Private Sub MainExplorer_AfterNodeSelectionChange(ByVal NewNode As Object)
    Select Case MainExplorer.SelectedNode.Level
        Case 0                                          'This is the printer-level selection
            PrinterName = MainExplorer.SelectedNode.Key
            PrinterControlForm.Setup
        Case 1                                          'This is the printer-specific selection
            PrinterName = MainExplorer.SelectedNode.Parent.Key
            MainExplorer.Caption = MainExplorer.SelectedNode.Parent.Text
            DB.GetPrinterRecordsets PrinterName
            Select Case MainExplorer.SelectedNode.Key
                Case "Settings"
                    SettingsControlForm.Setup
                Case "LUT"
                    LutControlForm.Setup                    'Setup LUT Control
                Case "Offset"
                    OffsetControlForm.Setup                 'Setup Offset Control
                Case "Color"
                    ColorControlForm.Setup                  'Setup Color Control
            End Select
    End Select
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdatePrinterList                                     **
'**                                                                        **
'**  Description..:  This routine updates the printer list.                **
'**                                                                        **
'****************************************************************************
Private Sub UpdatePrinterList()
    '---- Configure the Data Explorer
    Dim RootNode As pvxNode
    Dim Node As pvxNode
    'MainExplorer.TreeView.ImageList = MainImageList
    With MainExplorer.Nodes
        .RemoveAll
        DB.rsPrinterList.MoveFirst
        Do While Not DB.rsPrinterList.EOF
        
            If DB.rsPrinterList.Fields("PrinterName").Value <> "Default" Then
                Set RootNode = .AddRootNode(DB.rsPrinterList.Fields("Description").Value, 0, 1)
                RootNode.WindowObject = PrinterControlForm.hWnd
                RootNode.Key = DB.rsPrinterList.Fields("PrinterName").Value
                
                Set Node = .AddChild(RootNode, "Color", 0, 1)
                Node.WindowObject = ColorControlForm.hWnd
                Node.Key = "Color"
            
                Set Node = .AddChild(RootNode, "LUT", 0, 1)
                Node.WindowObject = LutControlForm.hWnd
                Node.Key = "LUT"
            
                Set Node = .AddChild(RootNode, "Offset", 0, 1)
                Node.WindowObject = OffsetControlForm.hWnd
                Node.Key = "Offset"
            
                Set Node = .AddChild(RootNode, "Settings", 0, 1)
                Node.WindowObject = SettingsControlForm.hWnd
                Node.Key = "Settings"
            End If
            DB.rsPrinterList.MoveNext
        Loop
    End With
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CommConnect                                           **
'**                                                                        **
'**  Description..:  This routine connects to the serial port.             **
'**                                                                        **
'****************************************************************************
Public Sub CommConnect()
    On Error GoTo ErrorHandler
    With DensComm
        If .PortOpen = True Then .PortOpen = False
        DoEvents
        
        .CommPort = 1               'Needs to be configurable  RDR
        
        
        .Settings = "9600,N,8,1"
        .RThreshold = 1
        .SThreshold = 0
        .PortOpen = True
        AppLog DebugMsg, "Form_Activate,Opened densitometer serial port on comm " & DensComm.CommPort
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "MainForm,CommConnect,Could not open comm port."
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  DensComm_OnComm                                       **
'**                                                                        **
'**  Description..:  This routine handles serial comm events               **
'**                                                                        **
'****************************************************************************
Private Sub DensComm_OnComm()
    On Error GoTo ErrorHandler
    '--- Handle Communications With Densitometer
    Dim Msgs() As String, MsgNum As Integer, DensiValues() As String
    Select Case DensComm.CommEvent
        Case comEvReceive
            If MainExplorer.SelectedNode.Key = "Color" Then                  'Color Ring Calibration
                ColorControlForm.SetColorTab 2
                Sleep 3000
                CommInputBuffer = DensComm.Input
                If Trim(CommInputBuffer) <> "" Then
                    Msgs = Split(Trim(CommInputBuffer), vbCrLf)
                    If UBound(Msgs) <> 11 Then                      'WILL BECOME 11 WHEN STRIP IS FIXED!!
                        MsgBox "Invalid Strip Reading."
                    Else
                        If UBound(Msgs) >= 1 Then
                            '---- if this is the first strip, confirm to add color ring
                            If ColorControlForm.ColorRingPass = 0 Then
                                If MsgBox("Scan new Color Ring Around?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Densitometer data received") = vbYes Then
                                    'DB.AddNewLUT
                                Else
                                    Exit Sub
                                End If
                            End If
                            For MsgNum = 0 To UBound(Msgs) - 1
                                DensiValues = Split(Msgs(MsgNum), " ")
                                With DB.rsRingAroundsAutoScan
                                    .MoveFirst
                                    .Find "BlockNum=" & (ColorControlForm.ColorRingPass + 1)
                                    If .EOF Then
                                        MsgBox "Color Pass definition not found.", vbApplicationModal + vbInformation + vbOKOnly, "Error"
                                        ColorControlForm.ColorRingPass = 0
                                        Exit Sub
                                    End If
                                    .Fields("ExpRed" & Trim(Str(MsgNum + 1))).Value = Val(Mid(DensiValues(0), 2, Len(DensiValues(0))))
                                    .Fields("ExpGreen" & Trim(Str(MsgNum + 1))).Value = Val(Mid(DensiValues(1), 2, Len(DensiValues(1))))
                                    .Fields("ExpBlue" & Trim(Str(MsgNum + 1))).Value = Val(Mid(DensiValues(2), 2, Len(DensiValues(2))))
                                    .UpdateBatch adAffectCurrent
                                End With
                            Next
                            ColorControlForm.UpdateColorCircleAutoScanGrid
                        End If
                        ColorControlForm.ColorRingPass = ColorControlForm.ColorRingPass + 1
                        If ColorControlForm.ColorRingPass > 6 Then
                            ColorControlForm.ColorRingPass = 0
                            'Update the Color Ring Table
                            
                        End If
                    End If
                End If
            End If
                
            If MainExplorer.SelectedNode.Key = "LUT" Then                       'LUT Densitometer Control
                Sleep 3000
                CommInputBuffer = DensComm.Input
                If Trim(CommInputBuffer) <> "" Then
                    Msgs = Split(Trim(CommInputBuffer), vbCrLf)
                    If UBound(Msgs) <> 16 Then
                        MsgBox "Invalid Strip Reading."
                    Else
                        If UBound(Msgs) >= 1 Then
                            '---- if this is the first strip, confirm to add LUT
                            If LutControlForm.DensitometerLutPass = 0 Then
                                If MsgBox("Add new LUT?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Densitometer data received") = vbYes Then
                                    DB.AddNewLUT
                                Else
                                    Exit Sub
                                End If
                            End If
                            For MsgNum = 0 To UBound(Msgs) - 1
                                DensiValues = Split(Msgs(MsgNum), " ")
                                With DB.rsLutDensiValues
                                    .MoveFirst
                                    .Find "BlockNum=" & (LutControlForm.DensitometerLutPass * 16) + MsgNum + 1
                                    If .EOF Then
                                        .AddNew
                                    End If
                                    .Fields("PrinterName").Value = PrinterName
                                    .Fields("LutNum").Value = DB.rsLuts.Fields("LutNum").Value
                                    .Fields("DensRed").Value = Val(Mid(DensiValues(0), 2, Len(DensiValues(0))))
                                    .Fields("DensGreen").Value = Val(Mid(DensiValues(1), 2, Len(DensiValues(1))))
                                    .Fields("DensBlue").Value = Val(Mid(DensiValues(2), 2, Len(DensiValues(2))))
                                    .UpdateBatch adAffectCurrent
                                End With
                            Next
                            LutControlForm.UpdateLutDensiGrid
                        End If
                        LutControlForm.DensitometerLutPass = LutControlForm.DensitometerLutPass + 1
                        If LutControlForm.DensitometerLutPass > 2 Then
                            LutControlForm.DensitometerLutPass = 0
                            LutControlForm.LoadLutTable 0
                        End If
                    End If
                End If
            End If
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError ":DensComm_OnComm", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddPrinter                                            **
'**                                                                        **
'**  Description..:  This routine adds a new printer to the list and       **
'**  and creates the subdirectory structure for all of it's setup files.   **
'**                                                                        **
'****************************************************************************
Public Sub AddPrinter()
    Dim PrinterID As String, Descript As String, TargetFolder As String
    '---- Get next printer ID & Description
    PrinterID = DB.GetNextPrinterID
    Descript = InputBox("Enter description for printer " & PrinterID & ":", "Add Printer", "")
    If Descript <> "" Then
        If MsgBox("Add new printer?", vbApplicationModal + vbQuestion + vbYesNoCancel + vbDefaultButton2, "Add Printer") = vbYes Then
            '---- Copy Default Printer Records to new Printer ID
            With DB.rsPrinterList
                .AddNew
                .Fields("PrinterName").Value = PrinterID
                .Fields("Description").Value = Descript
                .UpdateBatch adAffectCurrent
            End With
            UpdatePrinterList                                'Copy Options, Settings, & Hot Folders
        End If
    End If
End Sub

