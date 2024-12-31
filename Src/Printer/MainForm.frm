VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{22BE512E-E6B6-11D2-9BB5-00A0CC3AD9E7}#1.0#0"; "PVOutlookBar.ocx"
Object = "{1416D7C5-8A28-11CF-9236-444553540000}#8.0#0"; "PVXPLORE8.ocx"
Begin VB.Form MainForm 
   Caption         =   "ISA DVP2"
   ClientHeight    =   14505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   Picture         =   "MainForm.frx":0000
   ScaleHeight     =   14505
   ScaleWidth      =   19080
   Begin PVExplorerLib.PVExplorer MainExplorer 
      Height          =   735
      Left            =   5850
      TabIndex        =   1
      Top             =   14190
      Visible         =   0   'False
      Width           =   375
      _Version        =   524288
      Indentation     =   0
      SourceChannel1  =   ""
      TargetChannel1  =   ""
      PathSeparator   =   ""
      Image1          =   "MainForm.frx":0342
      SourceChannel2  =   ""
      TargetChannel2  =   ""
      Image2          =   "MainForm.frx":06C0
      Image3          =   "MainForm.frx":25BE
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
      PaneDisplay     =   3
      CaptionMode     =   1
      Appearance      =   1
      _ExtentX        =   661
      _ExtentY        =   1296
      _StockProps     =   70
   End
   Begin MSCommLib.MSComm DensComm 
      Left            =   0
      Top             =   13920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin OUTLOOKBARLibCtl.PVOutlookBar OutlookBarMain 
      Height          =   14145
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   19005
      _Version        =   131073
      SoundEffects    =   -1  'True
      Appearance      =   1
      BorderWidth     =   1
      SplitterWindow  =   -1  'True
      SplitterWidth   =   6
      GroupPopupMenu  =   -1  'True
      RenameGroups    =   -1  'True
      AddGroups       =   -1  'True
      RenameItems     =   -1  'True
      RemoveGroups    =   -1  'True
      AddItems        =   -1  'True
      RemoveItems     =   -1  'True
      HideOutlookBar  =   -1  'True
      ItemPopupMenu   =   -1  'True
      OpenItem        =   -1  'True
      Properties      =   -1  'True
      SizeIcons       =   -1  'True
      BackColor       =   -2147483636
      TextColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DyanmicResize   =   -1  'True
      UseChildWindows =   0   'False
   End
   Begin MSComctlLib.ImageList MainImageList 
      Left            =   18150
      Top             =   13920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":293C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":30A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":33C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":36DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":39F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":402A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4344
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4666
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4980
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5636
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5EDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer DiagnosticTimer 
      Interval        =   100
      Left            =   18630
      Top             =   14070
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   16650
      TabIndex        =   12
      Top             =   14220
      Width           =   1725
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   14940
      TabIndex        =   11
      Top             =   14220
      Width           =   1725
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   13230
      TabIndex        =   10
      Top             =   14220
      Width           =   1725
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   11520
      TabIndex        =   9
      Top             =   14220
      Width           =   1725
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   9810
      TabIndex        =   8
      Top             =   14220
      Width           =   1725
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8100
      TabIndex        =   7
      Top             =   14220
      Width           =   1725
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6900
      TabIndex        =   6
      Top             =   14220
      Width           =   1215
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5700
      TabIndex        =   5
      Top             =   14220
      Width           =   1215
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4410
      TabIndex        =   4
      Top             =   14220
      Width           =   1305
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2970
      TabIndex        =   3
      Top             =   14220
      Width           =   1455
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   14220
      Width           =   2955
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Digital VP-2                                              **
'**                                                                        **
'** Module.....: MainForm - The main application form.                     **
'**                                                                        **
'** Description: This form provides the main application interface.        **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 2002-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Private PrinterStatusText As String                         'Status text displayed in main status bar
Private CommInputBuffer As String                           'Serial communications input buffer

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Load                                             **
'**                                                                        **
'**  Description..:  This routine loads forms and initializes globals      **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    SetupOutlookBar
    If DB.StepperMaskInstalled = True Then
        Me.Caption = "ISA DVP-2 Version " & App.Major & "." & App.Minor & "." & App.Revision & IIf(DemoMode = True, " ** DEMO MODE ** ", "")
    Else
        Me.Caption = "ISA NORD Digital Printer Controller Version " & App.Major & "." & App.Minor & "." & App.Revision & IIf(DemoMode = True, " ** DEMO **", "")
    End If
    PrinterStatusText = PrinterIdleMessage
    CommConnect
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Form_Load", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Resize                                           **
'**                                                                        **
'**  Description..:  This routine resizes controls on form resize.         **
'**                                                                        **
'****************************************************************************
Private Sub Form_Resize()
    On Error GoTo ErrorHandler
    If Me.Width - 100 > 0 Then
        OutlookBarMain.Width = Me.Width - 100
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_QueryUnload                                      **
'**                                                                        **
'**  Description..:  This routine disconnects hardware and exits program.  **
'**                                                                        **
'****************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo, "Exit DVP-2") = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    DiagnosticTimer.Enabled = False                     'Stop the timer before unloading anything!
    DoEvents
    Sleep 100
    SizeSettingsForm.ClearImage                         'Clear the image
    Unload UsbKeyDiagnostics
    Unload PunchDiagnostics
    Unload frmSplash
    Unload FileErrorsForm
    Unload PasswordForm
    Unload ColorControlForm
    Unload EmulsionForm
    Unload LutControlForm
    Unload OffsetControlForm
    Unload SizeSettingsForm
    Unload BackWriterDiagnostics
    Unload MotorDiagnostics
    Unload HotFolderSettingsForm
    Unload PrinterQcForm
    Unload PrinterStatisticsForm
    Unload PrintQueHistoryForm
    Unload SettingsControlForm
    Unload DiagnosticsForm
    Unload PrinterConsole
    Unload ErrorForm
    CloseOutputDevice                                   'Close the LCD output device
    If FileSystemHandle.FileExists(CurrentPrintFile) = True Then
        Kill CurrentPrintFile
    End If
    CloseLogFile
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SetupOutlookBar                                       **
'**                                                                        **
'**  Description..:  This routine initializes the Outlook bar.             **
'**                                                                        **
'****************************************************************************
Public Sub SetupOutlookBar()
    On Error GoTo ErrorHandler
    '---- Configure the Data Explorer
    Dim RootNode As pvxNode
    Dim Node As pvxNode
    'MainExplorer.TreeView.ImageList = MainImageList
    With MainExplorer.Nodes
        .RemoveAll
        
        Set RootNode = .AddRootNode("Console", 0, 1)
        RootNode.WindowObject = PrinterConsole.hWnd
        RootNode.Key = "Console"
        
        Set Node = .AddChild(RootNode, "File Errors", 0, 1)
        Node.WindowObject = FileErrorsForm.hWnd
        Node.Key = "FileErrors"
        
        Set Node = .AddChild(RootNode, "Print History", 0, 1)
        Node.WindowObject = PrintQueHistoryForm.hWnd
        Node.Key = "PrintHistory"
        
        Set Node = .AddChild(RootNode, "Q.C. Mode", 0, 1)
        Node.WindowObject = PrinterQcForm.hWnd
        Node.Key = "QC"
    
        Set Node = .AddChild(RootNode, "Statistics", 0, 1)
        Node.WindowObject = PrinterStatisticsForm.hWnd
        Node.Key = "Statistics"
    
        Set Node = .AddChild(RootNode, "Settings", 0, 1)
        Node.WindowObject = SettingsControlForm.hWnd
        Node.Key = "Settings"
    
        Set Node = .AddChild(RootNode, "Color Settings", 0, 1)
        Node.WindowObject = ColorControlForm.hWnd
        Node.Key = "Color"
    
        Set Node = .AddChild(RootNode, "Offset", 0, 1)
        Node.WindowObject = OffsetControlForm.hWnd
        Node.Key = "Offset"
    
        Set Node = .AddChild(RootNode, "Hot Folders", 0, 1)
        Node.WindowObject = HotFolderSettingsForm.hWnd
        Node.Key = "HotFolders"
    
        Set Node = .AddChild(RootNode, "Emulsion", 0, 1)
        Node.WindowObject = EmulsionForm.hWnd
        Node.Key = "Emulsion"
    
        Set Node = .AddChild(RootNode, "LUT", 0, 1)
        Node.WindowObject = LutControlForm.hWnd
        Node.Key = "LUT"
    
        Set Node = .AddChild(RootNode, "Size Settings", 0, 1)
        Node.WindowObject = SizeSettingsForm.hWnd
        Node.Key = "SizeSettings"
    
        Set Node = .AddChild(RootNode, "Settings", 0, 1)
        Node.WindowObject = SettingsControlForm.hWnd
        Node.Key = "Settings"
    
        Set Node = .AddChild(RootNode, "Diagnostics", 0, 1)
        Node.WindowObject = DiagnosticsForm.hWnd
        Node.Key = "Diagnostics"
        
        Set Node = .AddChild(RootNode, "Backwriters", 0, 1)
        Node.WindowObject = BackWriterDiagnostics.hWnd
        Node.Key = "Backwriters"
        
        Set Node = .AddChild(RootNode, "Paper Advance & Mask", 0, 1)
        Node.WindowObject = MotorDiagnostics.hWnd
        Node.Key = "Motors"
   
        Set Node = .AddChild(RootNode, "Punch", 0, 1)
        Node.WindowObject = PunchDiagnostics.hWnd
        Node.Key = "Punch"
   
        Set Node = .AddChild(RootNode, "USB Key", 0, 1)
        Node.WindowObject = UsbKeyDiagnostics.hWnd
        Node.Key = "USBKey"
   
   
    End With

    Dim Group As PVOutlookGroup
    Dim Item As PVOutlookItem
    
    OutlookBarMain.LargeImageList = MainImageList
    
    '---- Main Group
    Set Group = OutlookBarMain.Groups.Add("Main")
    
    Set Item = Group.Items.Add("Console", 0)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Console")
    Item.Display
    
    Set Item = Group.Items.Add("File Errors", 8)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("FileErrors")

    Set Item = Group.Items.Add("Print History", 9)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("PrintHistory")

    Set Item = Group.Items.Add("Q.C. Mode", 11)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("QC")
    
    Set Item = Group.Items.Add("Statistics", 10)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Statistics")
    
    
    '---- Preferences Group
    Set Group = OutlookBarMain.Groups.Add("Preferences")
    
    Set Item = Group.Items.Add("Color Control", 3)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Color")
    
    Set Item = Group.Items.Add("Hot Folders", 7)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("HotFolders")
    
    Set Item = Group.Items.Add("Emulsion", 4)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Emulsion")
    
    Set Item = Group.Items.Add("LUT", 4)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("LUT")
    
    Set Item = Group.Items.Add("Offset", 5)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Offset")
    
    Set Item = Group.Items.Add("Size Settings", 1)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("SizeSettings")
    
    Set Item = Group.Items.Add("Settings", 6)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Settings")
    
    
    '---- Maintenance Group
    Set Group = OutlookBarMain.Groups.Add("Maintenance")
    
    Set Item = Group.Items.Add("Digital I/O", 2)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Diagnostics")
    
    Set Item = Group.Items.Add("Advance & Mask", 14)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Motors")
    
    Set Item = Group.Items.Add("Punch", 15)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Punch")
    
    Set Item = Group.Items.Add("USB Key", 13)
    Item.ActiveXObject1 = MainExplorer.object
    Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("USBKey")
    
    If DB.BackWritersInstalled = True Then
        Set Item = Group.Items.Add("Back Writers", 12)
        Item.ActiveXObject1 = MainExplorer.object
        Item.ExplorerNode = MainExplorer.Nodes.NodeFromKey("Backwriters")
    End If
    Exit Sub

ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SetupOutlookBar", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  DiagnosticTimer_Timer                                 **
'**                                                                        **
'**  Description..:  This routine returns monitors the machine.            **
'**                                                                        **
'****************************************************************************
Private Sub DiagnosticTimer_Timer()
    If DiagnosticsForm.PanelConnected = False Then
        Exit Sub
    End If
    DiagnosticsForm.ScanInputs                                 'Update digital I/O screen
    
    '---- Check for Image Time Out (Save LCD from burning in...)
    'If WatchForImageTimeOut = True Then
    '    If ImageTimeOut < 9000 Then                     '100ms timer = 600/min wait 5min @ 3000
    '        ImageTimeOut = ImageTimeOut + 1
    '    Else
    '        ClearImage
    '    End If
    'End If
    
    '---- Update the Status Bar
    Me.StatusLabel(0).Caption = PrinterStatusText
    If PrinterConsole.GetTimerEnabled = True Then
        StatusLabel(1).Caption = "Server Running"
        StatusLabel(1).BackColor = vbGreen
    Else
        StatusLabel(1).Caption = "Server Paused"
        StatusLabel(1).BackColor = vbRed
    End If
    StatusLabel(2).Caption = "Digital Control"
    StatusLabel(2).BackColor = IIf(DiagnosticsForm.PanelConnected = True, vbGreen, vbRed)
    
    StatusLabel(3).Caption = "Paper Drive"
    StatusLabel(3).BackColor = IIf(MotorDiagnostics.PaperAdvanceConnected, vbGreen, vbRed)
        
    StatusLabel(4).Caption = "RAMDISK"
    StatusLabel(4).BackColor = IIf(RamDiskConnected = True, vbGreen, vbRed)
    
    StatusLabel(5).Caption = "Paper Out"
    If DiagnosticsForm.IsInputEnabled(DiagnosticsForm.PaperOut) Then
        If DB.StepperMaskInstalled = True Then          'Sense is reversed in the Nord
            StatusLabel(5).BackColor = IIf(DiagnosticsForm.IsInputON(DiagnosticsForm.PaperOut) = True, vbGreen, vbRed)
        Else
            StatusLabel(5).BackColor = IIf(DiagnosticsForm.IsInputON(DiagnosticsForm.PaperOut) = True, vbRed, vbGreen)
        End If
    Else
        StatusLabel(5).BackColor = vbYellow
    End If
            
    If DB.StepperMaskInstalled = True Then              'Only for DVP2 with Stepper Mask
        StatusLabel(6).Caption = "Left Platen"
        If DiagnosticsForm.IsInputEnabled(DiagnosticsForm.PlattenLeft) Then
            StatusLabel(6).BackColor = IIf(DiagnosticsForm.IsInputON(DiagnosticsForm.PlattenLeft) = True, vbGreen, vbRed)
        Else
            StatusLabel(6).BackColor = vbYellow
        End If
        
        StatusLabel(7).Caption = "Right Platen"
        If DiagnosticsForm.IsInputEnabled(DiagnosticsForm.PlattenRight) Then
            StatusLabel(7).BackColor = IIf(DiagnosticsForm.IsInputON(DiagnosticsForm.PlattenRight) = True, vbGreen, vbRed)
        Else
            StatusLabel(7).BackColor = vbYellow
        End If
        
        StatusLabel(8).Caption = "P-Roller"
        If DiagnosticsForm.IsInputEnabled(DiagnosticsForm.PressureRollerEngaged) Then
            StatusLabel(8).BackColor = IIf(DiagnosticsForm.IsInputON(DiagnosticsForm.PressureRollerEngaged) = False, vbGreen, vbRed)
        Else
            StatusLabel(8).BackColor = vbYellow
        End If
        
        StatusLabel(9).Caption = "Door"
        If DB.DoorSwitchInstalled = False Then
            If DiagnosticsForm.IsInputEnabled(DiagnosticsForm.DoorClosed) Then
                StatusLabel(9).BackColor = IIf(DiagnosticsForm.IsInputON(DiagnosticsForm.DoorClosed) = True, vbGreen, vbRed)
            Else
                StatusLabel(9).BackColor = vbYellow
            End If
        Else
            StatusLabel(9).BackColor = vbGreen
        End If
        
        If DB.BackWritersInstalled = True Then
            StatusLabel(10).Caption = "BackWriters"
            StatusLabel(10).BackColor = vbGreen
        Else
            StatusLabel(10).Caption = ""
            StatusLabel(10).BackColor = RGB(255, 255, 255)
        End If
    
    End If

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
        .CommPort = 3              'Needs to be configurable  RDR
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
                    If UBound(Msgs) <> 16 Then                      'WILL BECOME 11 WHEN STRIP IS FIXED!!
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
                    
                    If UBound(Msgs) < 16 Or UBound(Msgs) > 20 Then
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
                            For MsgNum = 0 To 16
                                
                                Debug.Print Msgs(MsgNum)
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
                            LutControlForm.LoadLutTable DB.rsLuts("LutNum").Value
                        End If
                    End If
                End If
            End If
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError ":DensComm_OnComm", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub




Private Sub OutlookBarMain_BeforeItemChange(ByVal Group As OUTLOOKBARLibCtl.IPVOutlookGroup, ByVal Item As OUTLOOKBARLibCtl.IPVOutlookItem)
    On Error GoTo ErrorHandler
    If OutlookBarMain.CurrentItem.Text = "Size Settings" And Item.Text <> "Size Settings" Then
        '----- Close Shutters....
        With SizeSettingsForm
            If .SizeSettingsToolBars.Tools("ID_MotorOn").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_MotorOn").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Red").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Red").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Green").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Green").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Blue").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Blue").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Lamp").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Lamp").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Iris").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Iris").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Pan").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Pan").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Ruler").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Ruler").State = ssUnchecked
            End If
            If .SizeSettingsToolBars.Tools("ID_Focus").State = ssChecked Then
                .SizeSettingsToolBars.Tools("ID_Focus").State = ssUnchecked
            End If
        End With
    End If
    
    If OutlookBarMain.CurrentItem.Text = "Digital I/O" And Item.Text <> "Digital I/O" Then
        '----- Disable Outputs....
        With DB.rsOutputs
            .MoveFirst
            Do While Not .EOF
                DiagnosticsForm.Outputs.ActiveRow.Cells(3).Value = False
                .MoveNext
            Loop
        End With
    End If
    
    If OutlookBarMain.CurrentItem.Text = "Settings" And Item.Text <> "Settings" Then
        '----- Update Settings from database
        DB.GetPrinterSettings
    End If
    
    
    If OutlookBarMain.CurrentItem.Text = "Console" And Item.Text <> "Console" Then
        '----- Update Settings from database
        
       PrinterConsole.QueTimer.Enabled = False
        
    End If
    
    If OutlookBarMain.CurrentItem.Text <> "Console" And Item.Text = "Console" Then
    
        PrinterConsole.QueTimer.Enabled = True
    
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":OutlookBarMain_BeforeItemChange", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

