VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form BackWriterDiagnostics 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16665
   LinkTopic       =   "Form1"
   ScaleHeight     =   12465
   ScaleWidth      =   16665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter BackWriterSplitter 
      Height          =   12075
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   16605
      _ExtentX        =   29289
      _ExtentY        =   21299
      _Version        =   262144
      SplitterBarJoinStyle=   0
      PaneTree        =   "BackWriterDiagnostics.frx":0000
      Begin Threed.SSFrame BackWriterDiagnosticsFrame 
         Height          =   12015
         Left            =   7590
         TabIndex        =   1
         Top             =   30
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   21193
         _Version        =   262144
         Caption         =   "Communication Diagnostics"
         Begin VB.TextBox SendCommandBuffer 
            Height          =   345
            Left            =   720
            TabIndex        =   3
            Top             =   11430
            Width           =   4755
         End
         Begin VB.TextBox InputBuffer 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   11115
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   210
            Width           =   8805
         End
         Begin VB.Label CommSendImmediateLabel 
            Caption         =   "Send:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Top             =   11460
            Width           =   465
         End
      End
      Begin UltraGrid.SSUltraGrid BackWriterSettings 
         Height          =   12015
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   21193
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "BackWriter Settings"
      End
   End
   Begin MSCommLib.MSComm BackWriterComm 
      Left            =   30
      Top             =   11850
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin ActiveToolBars.SSActiveToolBars BackWriterToolBar 
      Left            =   30
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "BackWriterDiagnostics.frx":0052
      ToolBars        =   "BackWriterDiagnostics.frx":4CF9
   End
End
Attribute VB_Name = "BackWriterDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP2                                                      **
'**                                                                        **
'** Module.....: BackWriterDiagnostics                                     **
'**                                                                        **
'** Description: This form provides backwriter functionality.              **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Private CommBuffer As String                            'Simple buffer for comm port
Private WaitingForResponse As Boolean                   'TRUE if waiting for a response
Private Text1 As String                                 'Backwriter text for head 1
Private Text2 As String                                 'Backwriter text for head 2
Private CharDelay As Integer                            'Delay between characters (micro-sec)
Private CharHeight As Integer                           'Height of characters 7,9,11
Private DotOnTime As Integer                            'Time for DOTs ON (micro-sec)
Private DotOffTime As Integer                           'Time for DOTs OFF (micro-sec)
Private MaxChars As Integer                             'Max # chars
Private DelayFromStart As Integer                       'Delay from start of advance (ms)
Private RibbonMotorSpeed As Integer                     'Speed of Ribbon Motor

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Unload                                           **
'**                                                                        **
'**  Description..:  This routine disconnects serial ports.                **
'**                                                                        **
'****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    CommDisconnect
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
        Me.BackWriterSplitter.Width = Me.Width - 100
        Me.BackWriterDiagnosticsFrame.Width = Me.BackWriterSplitter.Panes("Pane B").Width - 100
        Me.InputBuffer.Width = Me.BackWriterDiagnosticsFrame.Width - 100
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  BackWriterToolBar_ToolClick                           **
'**                                                                        **
'**  Description..:  This routine handles the toolbar.                     **
'**                                                                        **
'****************************************************************************
Private Sub BackWriterToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Select Case Tool.ID
        Case "ID_Print"
            With DB.rsBackWriterSettings
                .MoveFirst      'Text 1
                CommSend Chr(34) & Trim(.Fields("SettingValue").Value) & Chr(34) & ":HEAD1_DATA"
                .MoveNext       'Text 2
                CommSend Chr(34) & Trim(.Fields("SettingValue").Value) & Chr(34) & ":HEAD2_DATA"
            End With
            CommSend ":PRINT_DATA"
            MotorDiagnostics.AdvancePaper 7
        Case "ID_Test"
            CommSend ":TEST_PRINT"
            MotorDiagnostics.AdvancePaper 20
        Case "ID_AdvancePaper"
            MotorDiagnostics.AdvancePaper 10
        Case "ID_GetSettings"
            CommSend ":SETUP?"
        Case "ID_SendSettings"
            SendBackWriterSettings
        Case "ID_CommandHelp"
            CommSend ":HELP"
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":BackWriterToolBar_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Public Sub Setup()
    On Error GoTo ErrorHandler
    If DB.BackWritersInstalled Then
        BackWriterDiagMode DB.BackWritersInstalled     'Determine whether or not to display backwriter frame
        With BackWriterSettings
            Set .DataSource = DB.rsBackWriterSettings
            .Refresh ssRefetchAndFireInitializeRow
            .Bands(0).Columns(0).Hidden = True
            .Bands(0).Columns(1).Hidden = True
            .Bands(0).Columns(2).Header.Caption = "Setting"
            .Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
            .Bands(0).Columns(2).AutoEdit = False
            .Bands(0).Columns(2).Width = 2500
            .Bands(0).Columns(3).Header.Caption = "Value"
            .Bands(0).Columns(3).Width = 3700
        End With
        CommConnect 2, "38400,n,8,1"                    'open Backwriter serial port
        SendBackWriterSettings                          'Send current settings to BackWriter Controller
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SendBackWriterSettings                                **
'**                                                                        **
'**  Description..:  This routine sends settings to the backwriters.       **
'**                                                                        **
'****************************************************************************
Private Sub SendBackWriterSettings()
    On Error GoTo ErrorHandler
    If DB.BackWritersInstalled Then
        With DB.rsBackWriterSettings
            .MoveFirst      'Text 1
            .MoveNext       'Text 2
            .MoveNext       'Char Delay
            CommSend Trim(.Fields("SettingValue").Value) & ":CHAR_DELAY_TIME"
            .MoveNext       'On Time
            CommSend Trim(.Fields("SettingValue").Value) & ":DOT_ON_TIME"
            .MoveNext       'Off Time
            CommSend Trim(.Fields("SettingValue").Value) & ":DOT_OFF_TIME"
            .MoveNext       'Max Chars
            CommSend Trim(.Fields("SettingValue").Value) & ":MAX_NUM_CHAR"
            .MoveNext       'Delay start
            CommSend Trim(.Fields("SettingValue").Value) & ":DELAY_START"
            .MoveNext       'RIBBON SPEED
            CommSend Trim(.Fields("SettingValue").Value) & ":RIBBON_SPEED"
            .MoveNext       'CHAR SIZE
            CommSend ":CHAR_SIZE_" & Trim(.Fields("SettingValue").Value)
            .MoveNext       'STEPS ON
            If Trim(UCase(.Fields("SettingValue").Value)) = "YES" Then
                CommSend ":STEPS_ON"
            Else
                CommSend ":STEPS_OFF"
            End If
        End With
        Sleep 500
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SendBackWriterSettings", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  WriteText                                             **
'**                                                                        **
'**  Description..:  This routine sends text to the backwriters.           **
'**                                                                        **
'****************************************************************************
Public Sub WriteText()
    On Error GoTo ErrorHandler
    If DB.BackWritersInstalled Then
        With DB.rsBackWriterSettings
            .MoveFirst      'Text 1
            CommSend Chr(34) & Trim(.Fields("SettingValue").Value) & Chr(34) & ":HEAD1_DATA"
            .MoveNext       'Text 2
            CommSend Chr(34) & Trim(.Fields("SettingValue").Value) & Chr(34) & ":HEAD2_DATA"
        End With
        CommSend ":PRINT_DATA"
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":WriteText", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  WriteFixedText                                        **
'**                                                                        **
'**  Description..:  This routine sends fixed text to the backwriters.     **
'**                                                                        **
'****************************************************************************
Public Sub WriteFixedText(Text1 As String, Text2 As String)
    On Error GoTo ErrorHandler
    If DB.BackWritersInstalled Then
        CommSend Chr(34) & Trim(Text1) & Chr(34) & ":HEAD1_DATA"
        CommSend Chr(34) & Trim(Text2) & Chr(34) & ":HEAD2_DATA"
        CommSend ":PRINT_DATA"
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":WriteFixedText", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CommConnect                                           **
'**                                                                        **
'**  Description..:  This routine connects to serial ports.                **
'**                                                                        **
'****************************************************************************
Public Function CommConnect(CommPort As Integer, CommSettings As String) As Boolean
    On Error GoTo ErrorHandler
    CommBuffer = ""
    WaitingForResponse = False
    CommConnect = False
    If DB.BackWritersInstalled Then
        With BackWriterComm
            If .PortOpen = True Then .PortOpen = False
            DoEvents                                        'RARE DoEvents to make sure comm port is closed!
            .CommPort = CommPort
            .Settings = CommSettings
            .RThreshold = 1
            .SThreshold = 0
            .PortOpen = True
        End With
        AppLog InfoMsg, "CommConnect,Opened Port " & CommPort & ",Settings " & CommSettings
        CommConnect = True
    End If
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":CommConnect", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CommDisconnect                                        **
'**                                                                        **
'**  Description..:  This routine disconnects serial ports.                **
'**                                                                        **
'****************************************************************************
Public Sub CommDisconnect()
    On Error GoTo ErrorHandler
    If BackWriterComm.PortOpen = True Then
        BackWriterComm.PortOpen = False
        AppLog InfoMsg, "CommDisconnect,Closed Backwriter port."
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":CommDisconnect", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Comm_OnComm                                           **
'**                                                                        **
'**  Description..:  This routine handles all comm events on all ports.    **
'**                                                                        **
'****************************************************************************
Private Sub BackWriterComm_OnComm()
    '---- Handle BackWriter Serial Communications
    On Error GoTo ErrorHandler
    Dim Msgs() As String, MsgNum As Integer, DensiValues() As String
    BackWriterComm.RThreshold = 0                          'This will disable the event (prevent recursion)
    Select Case BackWriterComm.CommEvent
        Case comEventBreak                          'A Break was received.
        Case comEventFrame                          'Framing Error
        Case comEventOverrun                        'Data Lost
        Case comEventRxOver                         'Receive buffer overflow
        Case comEventRxParity                       'Parity Error
        Case comEventTxFull                         'Transmit buffer full
        Case comEventDCB                            'Unexpected error retrieving DCB
        Case comEvCD                                'Change in the CD line
        Case comEvCTS                               'Change in the CTS line
        Case comEvDSR                               'Change in the DSR line
        Case comEvRing                              'Change in the Ring Indicator
        Case comEvReceive                           'Received RThreshold # of chars
            Dim LastMessage As String
            Dim CommInputBuffer As String
            Dim EndChar As Integer
            Dim ByteNum As Integer
            Dim DoneParsing As Boolean
            Dim DevIndex As Integer
            Dim DoneGettingData As Boolean, TimeOut As Integer
            DoneGettingData = False
            TimeOut = 0
            Do While Not DoneGettingData
                If BackWriterComm.InBufferCount > 0 Then
                    CommInputBuffer = CommInputBuffer & BackWriterComm.Input
                Else
                    TimeOut = TimeOut + 1
                    If TimeOut > 15 Then
                        DoneGettingData = True
                    Else
                        Sleep 10
                    End If
                End If
            Loop
            If Trim(CommInputBuffer) <> "" Then
                Msgs = Split(Trim(CommInputBuffer), vbCrLf)
                If UBound(Msgs) >= 1 Then
                    For MsgNum = 0 To UBound(Msgs)
                        SerialLog Msgs(MsgNum)
                    Next
                End If
            End If
        Case comEvSend                                      'SThreshold number of characters in the transmit buffer
        Case comEvEOF                                       'An EOF charater was found in the input stream
    End Select
    BackWriterComm.RThreshold = 1
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":BackWriterComm_OnComm", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CommSend                                              **
'**                                                                        **
'**  Description..:  This routine sends data out a specified serial port.  **
'**                                                                        **
'****************************************************************************
Public Sub CommSend(SendMessage As String)
    On Error GoTo ErrorHandler
    If BackWriterComm.PortOpen = True Then
        If Not DemoMode Then
            BackWriterComm.Output = SendMessage & vbCr      'Markers need carriage return
            Sleep 80                                        'The motors will not accept commands faster than 40ms - they will respond with ? instead of > prompt.
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":CommSend", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SendCommandBuffer_KeyDown                             **
'**                                                                        **
'**  Description..:  This routine sends user text out specified serial port**
'**                                                                        **
'****************************************************************************
Public Sub SendCommandBuffer_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    If Not DemoMode Then
        If KeyCode = vbKeyReturn Then
            With SendCommandBuffer
                CommSend Trim(.Text)
                .Text = ""
                .Refresh
            End With
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SendCommandBuffer_KeyDown", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SerialLog                                             **
'**                                                                        **
'**  Description..:  This routine applog s messages to serial text box.    **
'**                                                                        **
'****************************************************************************
Public Sub SerialLog(MessageText As String)
    On Error GoTo ErrorHandler
    Dim DevName As String
    With InputBuffer
        If Len(.Text) > 8196 Then
            .Text = ""
        End If
        .Text = .Text & Format(Timer, "000000.00") & " " & MessageText & vbCrLf
        .SelStart = Len(.Text)
        .Refresh
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Log", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Public Sub BackWriterDiagMode(Mode As Boolean)
  '  CommDevice(3).Visible = Mode
  '  CommDevice(3).Enabled = Mode
End Sub


