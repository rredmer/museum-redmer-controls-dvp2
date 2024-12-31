VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form MotorDiagnostics 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   13395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   13395
   ScaleWidth      =   17160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars MotorControlToolBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   2
      ToolsCount      =   7
      Tools           =   "MotorDiagnostics.frx":0000
      ToolBars        =   "MotorDiagnostics.frx":594D
   End
   Begin Threed.SSFrame SerialDiagnosticsFrame 
      Height          =   12345
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   17145
      _ExtentX        =   30242
      _ExtentY        =   21775
      _Version        =   262144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Serial Communications Diagnostics"
      Begin VB.TextBox SendCommandBuffer 
         Height          =   345
         Left            =   720
         TabIndex        =   2
         Top             =   11400
         Width           =   4905
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   16965
      End
      Begin Threed.SSOption CommDevice 
         Height          =   315
         Index           =   0
         Left            =   5760
         TabIndex        =   3
         Top             =   11370
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Paper Advance"
      End
      Begin Threed.SSOption CommDevice 
         Height          =   315
         Index           =   1
         Left            =   5760
         TabIndex        =   4
         Top             =   11670
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Left Mask"
      End
      Begin Threed.SSOption CommDevice 
         Height          =   315
         Index           =   2
         Left            =   5760
         TabIndex        =   5
         Top             =   11970
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Right Mask"
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
         TabIndex        =   6
         Top             =   11430
         Width           =   615
      End
   End
   Begin MSCommLib.MSComm Comm 
      Index           =   0
      Left            =   60
      Top             =   12750
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   690
      Top             =   12810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "MotorDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP2                                                      **
'**                                                                        **
'** Module.....: LUTControl                                                **
'**                                                                        **
'** Description: This module provides advance & mask motor control         **
'**                                                                        **
'** History....:                                                           **
'**    10/02/03 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Private Enum SerialDevices
    PaperAdvanceControl = 0                             'Index of paper advance serial control
    PaperMaskLeftControl = 1                            'Index of paper mask left serial control
    PaperMaskRightControl = 2                           'Index of paper mask right serial control
End Enum
Public PaperAdvanceConnected As Boolean
Private CommBuffer As String                            'Simple buffer for comm port
Private WaitingForResponse As Boolean                   'TRUE if waiting for a response
Private PaperMotorStillMoving As Boolean
Private PaperMaskLeftStillMoving As Boolean
Private PaperMaskRightStillMoving As Boolean

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Public Sub Setup()
    On Error GoTo ErrorHandler
    PaperAdvanceConnected = False
    If Not DemoMode Then
        AppLog InfoMsg, "InitializeHardware,Connecting to Motors..."
        If DB.StepperMaskInstalled = True Then
            '--- DVP-2 has paper advance on COM4 & paper mask motors on 5 & 9
            AppLog InfoMsg, "InitializeHardware,Connecting to devices using RS-422 configuration."
            CommConnect PaperAdvanceControl, 1, "9600,N,8,1"    'Connect to paper advance serial port
            CommSend PaperAdvanceControl, vbLf
            CommSend PaperMaskLeftControl, vbLf
            CommSend PaperMaskRightControl, vbLf
        Else
            '--- Nord has paper advance on COM1
            AppLog InfoMsg, "InitializeHardware,Connecting to devices using Nord RS-232 configuration."
            CommConnect PaperAdvanceControl, 1, "9600,N,8,1"    'Connect to paper advance serial port
            CommDevice(1).Visible = False
            CommDevice(1).Enabled = False
            CommDevice(2).Visible = False
            CommDevice(2).Enabled = False
        End If
        InitializeStepperMask
        InitializePaperAdvance
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
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
        SerialDiagnosticsFrame.Width = Me.Width - 100
        InputBuffer.Width = SerialDiagnosticsFrame.Width - 200
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CommDisconnect PaperAdvanceControl                  'Disconnect from paper advance motion controller
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MotorControlToolBar_ToolClick                         **
'**                                                                        **
'**  Description..:  This routine handles the serial command tool bar.     **
'**                                                                        **
'****************************************************************************
Private Sub MotorControlToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Select Case Tool.ID
        Case "ID_Erase"
            InputBuffer.Text = ""
            InputBuffer.Refresh
        Case "ID_Copy"
            Clipboard.Clear                             'Clear the clipboard contents
            Clipboard.SetText InputBuffer.Text, vbCFText  'Save the text to the clipboard
        Case "ID_Save"
            '---- Show the save file dialog
            CommonDialog.ShowSave
            If CommonDialog.FileName <> "" Then
                Dim BufFileSystem As New Scripting.FileSystemObject     'Pointer to Bufor File System Object
                Dim BufFile As Scripting.TextStream                     'Pointer to Bufor Report File
                SaveText CommonDialog.FileName, InputBuffer.Text
                MsgBox "Created Text File: " & CommonDialog.FileName, vbApplicationModal + vbOKOnly + vbInformation, "Finished"
            End If
        Case "ID_GetSettings"
            MsgBox "Coming soon..."
        Case "ID_Initialize"
            InitializePaperAdvance
            InitializeStepperMask
        Case "Home"
            PaperMaskHome
        Case "GetSwitchSettings"
            GetMaskStatus 1
            GetMaskStatus 2
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":MotorControlToolBar_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SerialLog                                             **
'**                                                                        **
'**  Description..:  This routine applog s messages to serial text boxes.  **
'**                                                                        **
'****************************************************************************
Public Sub SerialLog(index As Integer, MessageText As String)
    On Error GoTo ErrorHandler
    Dim DevName As String
    With InputBuffer
        If Len(.Text) > 8196 Then
            .Text = ""
        End If
        Select Case index
            Case PaperAdvanceControl
                DevName = "[Advance   ]"
            Case PaperMaskLeftControl
                DevName = "[LeftMask  ]"
            Case PaperMaskRightControl
                DevName = "[RightMask ]"
        End Select
        .Text = .Text & Format(Timer, "000000.00") & " " & DevName & MessageText & vbCrLf
        .SelStart = Len(.Text)
        .Refresh
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Log", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CommConnect                                           **
'**                                                                        **
'**  Description..:  This routine connects to serial ports.                **
'**                                                                        **
'****************************************************************************
Public Function CommConnect(index As Integer, CommPort As Integer, CommSettings As String) As Boolean
    On Error GoTo ErrorHandler
    CommBuffer = ""
    WaitingForResponse = False
    CommConnect = False
    With Comm(index)
        If .PortOpen = True Then .PortOpen = False
        DoEvents                                        'RARE DoEvents to make sure comm port is closed!
        .CommPort = CommPort
        .Settings = CommSettings
        .RThreshold = 1
        .SThreshold = 0
        .PortOpen = True
    End With
    AppLog InfoMsg, "CommConnect,Opened Index " & index & ",Port " & CommPort & ",Settings " & CommSettings
    CommConnect = True
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
Public Sub CommDisconnect(index As Integer)
    On Error GoTo ErrorHandler
    If Comm(index).PortOpen = True Then
        Comm(index).PortOpen = False
        AppLog InfoMsg, "CommDisconnect,Closed Index " & index
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
Public Sub Comm_OnComm(index As Integer)
    On Error GoTo ErrorHandler
    
    Dim Msgs() As String, MsgNum As Integer, DensiValues() As String
    Comm(index).RThreshold = 0                          'This will disable the event (prevent recursion)
    Select Case Comm(index).CommEvent
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
            
            If DB.RS422MotorConfiguration = True Then
                '---- RS-422 Control
                Select Case index
                    '---- Special handling for stepper motors
                    Case PaperAdvanceControl, PaperMaskLeftControl, PaperMaskRightControl
                        CommInputBuffer = Comm(PaperAdvanceControl).Input
                        If Trim(CommInputBuffer) <> "" Then
                            Msgs = Split(Trim(CommInputBuffer), vbLf)
                            If UBound(Msgs) >= 1 Then
                                For MsgNum = 0 To UBound(Msgs)
                                    LastMessage = Msgs(MsgNum)
                                    If Left(LastMessage, 1) = ">" Then
                                        LastMessage = Mid(LastMessage, 2, Len(LastMessage))
                                    End If
                                    Select Case UCase(Left(LastMessage, 1))
                                        Case "A"
                                            DevIndex = PaperAdvanceControl
                                        Case "B"
                                            DevIndex = PaperMaskLeftControl
                                        Case "C"
                                            DevIndex = PaperMaskRightControl
                                    End Select
                                    If LastMessage <> "" Then
                                        SerialLog DevIndex, LastMessage
                                    End If
                                    If UCase(Mid(LastMessage, 2, 5)) = "PR MV" Then
                                        LastMessage = Msgs(MsgNum + 1)
                                        SerialLog DevIndex, "MV=" & LastMessage
                                        If Val(Left(LastMessage, 1)) = 0 Then
                                            Select Case DevIndex
                                                Case PaperAdvanceControl
                                                    PaperMotorStillMoving = False
                                                Case PaperMaskLeftControl
                                                    PaperMaskLeftStillMoving = False
                                                Case PaperMaskRightControl
                                                    PaperMaskRightStillMoving = False
                                            End Select
                                        Else
                                            Select Case DevIndex
                                                Case PaperAdvanceControl
                                                    PaperMotorStillMoving = True
                                                Case PaperMaskLeftControl
                                                    PaperMaskLeftStillMoving = True
                                                Case PaperMaskRightControl
                                                    PaperMaskRightStillMoving = True
                                            End Select
                                        End If
                                    End If
                                    If UCase(Mid(LastMessage, 2, 8)) = "PRINT I1" Then
                                        If UCase(Left(LastMessage, 1)) = "B" Then
                                            SerialLog index, "Left Home = " & Mid(LastMessage, 12, 1)
                                            'LeftMaskHoValue = Val(Mid(LastMessage, 12, 1))
                                            'LeftMaskHoRefresh
                                        Else
                                            SerialLog index, "Right Home = " & Mid(LastMessage, 12, 1)
                                            'RightMaskHoValue = Val(Mid(LastMessage, 12, 1))
                                            'RightMaskHoRefresh
                                        End If
                                    End If
                                Next
                            End If
                        End If
                End Select
            Else
                CommInputBuffer = Comm(PaperAdvanceControl).Input
                If Trim(CommInputBuffer) <> "" Then
                    Msgs = Split(Trim(CommInputBuffer), vbCrLf)
                    If UBound(Msgs) >= 1 Then
                        For MsgNum = 0 To UBound(Msgs)
                            LastMessage = Msgs(MsgNum)
                            If Left(LastMessage, 1) = ">" Then
                                LastMessage = Mid(LastMessage, 2, Len(LastMessage))
                            End If
                            If LastMessage <> "" Then
                                SerialLog PaperAdvanceControl, LastMessage
                            End If
                            If UCase(Mid(LastMessage, 1, 5)) = "PR MV" Then         'No A prefix for Nord!
                                LastMessage = Msgs(MsgNum + 1)
                                SerialLog PaperAdvanceControl, "MV=" & LastMessage
                                If Val(Left(LastMessage, 1)) = 0 Then
                                    PaperMotorStillMoving = False
                                Else
                                    PaperMotorStillMoving = True
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Case comEvSend                              'SThreshold number of characters in the transmit buffer
        Case comEvEOF                               'An EOF charater was found in the input stream
    End Select
    Comm(index).RThreshold = 1
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Comm_OnComm", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CommSend                                              **
'**                                                                        **
'**  Description..:  This routine sends data out a specified serial port.  **
'**                                                                        **
'****************************************************************************
Public Sub CommSend(index, SendMessage As String)
    On Error GoTo ErrorHandler
    If Not DemoMode Then
        If DB.RS422MotorConfiguration = True Then
            If Comm(PaperAdvanceControl).PortOpen = False Then
                Exit Sub
            End If
            Select Case index
                Case PaperAdvanceControl
                    Comm(PaperAdvanceControl).Output = "A" & SendMessage & vbLf      'Send LF to motors in Party Mode
                Case PaperMaskLeftControl
                    If DB.StepperMaskInstalled = True Then
                        Comm(PaperAdvanceControl).Output = "B" & SendMessage & vbLf      'Send LF to motors in Party Mode
                    End If
                Case PaperMaskRightControl
                    If DB.StepperMaskInstalled = True Then
                        Comm(PaperAdvanceControl).Output = "C" & SendMessage & vbLf      'Send LF to motors in Party Mode
                    End If
            End Select
        Else
            Comm(PaperAdvanceControl).Output = SendMessage & vbCr               'Send CR to motors not in Party Mode
        End If
        Sleep 80                                                                 'The motors will not accept commands faster than 40ms - they will respond with ? instead of > prompt.
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
        Dim index As Integer
        If KeyCode = vbKeyReturn Then
            If CommDevice(0).Value = True Then
                index = PaperAdvanceControl
            End If
            If CommDevice(1).Value = True Then
                index = PaperMaskLeftControl
            End If
            If CommDevice(2).Value = True Then
                index = PaperMaskRightControl
            End If
            With SendCommandBuffer
                CommSend index, Trim(.Text)
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
'**  Procedure....:  InitializePaperAdvance                                **
'**                                                                        **
'**  Description..:  This routine initializes the paper advance motor      **
'**                                                                        **
'****************************************************************************
Public Function InitializePaperAdvance() As Boolean
    If Not DemoMode Then
        AppLog DebugMsg, "InitializePaperAdvance,Sending paper advance motor parameters..."
        CommSend PaperAdvanceControl, "RC=100"                  '236 PPI for DVP-2, 216 PPI for Nord (typically).
        If DB.StepperMaskInstalled = True Then
            CommSend PaperAdvanceControl, "HC=50"
        Else
            CommSend PaperAdvanceControl, "HC=80"
        End If
        CommSend PaperAdvanceControl, "MS=4"
        CommSend PaperAdvanceControl, "A=5000"
        CommSend PaperAdvanceControl, "D=5000"
        CommSend PaperAdvanceControl, "VM=2500"
        CommSend PaperAdvanceControl, "S4=17,0"                 'Set pin high when paper advancing for backwriting  (May need to change to S1)
        CommSend PaperAdvanceControl, "S"                       'Save these parameters to memory
        DoEvents
        Sleep 200
        PaperMotorTorqueOFF                                     'Turn torque off to ease loading
        PaperAdvanceConnected = True
    End If
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  WaitForPaperAdvance                                   **
'**                                                                        **
'**  Description..:  This routine waits for paper advance motor to stop    **
'**                                                                        **
'****************************************************************************
Public Function WaitForPaperAdvance() As Boolean
    If Not DemoMode Then
        Dim wPC As New PerformanceCounter
        wPC.StartTimer True
        WaitForPaperAdvance = True          'We will assume the motor is moving, let it tell us otherwise
        PaperMotorStillMoving = True
        If Comm(PaperAdvanceControl).PortOpen = True Then
            Do While PaperMotorStillMoving = True
                'Log DebugMsg, "WaitForPaperAdvance,Checking for motor stopped..."
                CommSend PaperAdvanceControl, "PR MV"               'Ask if paper motor is still moving
                DoEvents                                            'Handle windows events while we wait - this MUST be here!
            Loop
        End If
        AppLog DebugMsg, "WaitForPaperAdvance,Timed," & Format(wPC.StopTimer, "####.####")
        Set wPC = Nothing
    Else
        WaitForPaperAdvance = False
        PaperMotorStillMoving = False
    End If
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AdvancePaper                                          **
'**                                                                        **
'**  Description..:  This routine advances the paper - button on form.     **
'**                                                                        **
'****************************************************************************
Public Sub AdvancePaper(AdvanceLength As Single)
    If AdvanceLength = 0 Or AdvanceLength > 50 Or DemoMode Then
        Exit Sub
    End If
    '---- We have to make sure the motor is not moving prior to sending another move command
    If WaitForPaperAdvance() = False Then
        MsgBox "Motor error.  Timed out waiting for motor to stop.", vbApplicationModal + vbCritical + vbOKOnly, "ERROR"
    Else
        PaperMotorTorqueON
        CommSend PaperAdvanceControl, "P=0"
        If DB.StepperMaskInstalled = True Then
            If DiagnosticsForm.MarkTextRequired = True Then
                AppLog DebugMsg, "AdvancePaper,BackWriting=" & DiagnosticsForm.MarkText
                BackWriterDiagnostics.WriteFixedText DiagnosticsForm.MarkText, DiagnosticsForm.MarkText
                DiagnosticsForm.MarkTextRequired = False
            Else
                AppLog DebugMsg, "AdvancePaper,Backwriting not required."
            End If
            CommSend PaperAdvanceControl, "MA " & Format(AdvanceLength * DB.MotorPPISetting, "####")    'Standard Poles:  Multiply = 15100
        Else
            CommSend PaperAdvanceControl, "MA -" & Format(AdvanceLength * DB.MotorPPISetting, "####")    'Standard Poles:  Multiply = 15100
        End If
    End If
    'AppLog DebugMsg, "AdvancePaper,Timed," & Format(PC.StopTimer, "####.####") & " seconds."
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PaperMotorTorqueON/OFF                                **
'**                                                                        **
'**  Description..:  These routines set paper advance holding torque.      **
'**                                                                        **
'****************************************************************************
Public Sub PaperMotorTorqueON()

    If DB.StepperMaskInstalled = True Then
        CommSend PaperAdvanceControl, "HC=50"
    Else
        CommSend PaperAdvanceControl, "HC=85"
    End If
    
End Sub

Public Sub PaperMotorTorqueOFF()
    CommSend PaperAdvanceControl, "HC=0"
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  InitializeStepperMask                                 **
'**                                                                        **
'**  Description..:  This routine initializes the stepper mask motors      **
'**                                                                        **
'****************************************************************************
Public Sub InitializeStepperMask()
    If DB.StepperMaskInstalled <> True Then Exit Sub
    If Not DemoMode Then
        If DB.RS422MotorConfiguration = True Then
            AppLog DebugMsg, "InitializeStepperMask,Sending stepper mask motor parameters via RS-422..."
            CommSend PaperMaskLeftControl, Chr(3)
            CommSend PaperMaskRightControl, Chr(3)
            Sleep 500
            CommSend PaperMaskLeftControl, vbLf
            CommSend PaperMaskRightControl, vbLf
            Sleep 250
        Else
            AppLog DebugMsg, "InitializeStepperMask,Sending stepper mask motor parameters via RS-232..."
            CommSend PaperMaskLeftControl, vbCr
            CommSend PaperMaskRightControl, vbCr
        End If
        CommSend PaperMaskLeftControl, "HC=50"
        CommSend PaperMaskLeftControl, "A=10000000"
        CommSend PaperMaskLeftControl, "D=10000000"
        CommSend PaperMaskLeftControl, "VM=75000"
        CommSend PaperMaskLeftControl, "S1=1,0"
        Sleep 250
        CommSend PaperMaskRightControl, "HC=50"
        CommSend PaperMaskRightControl, "A=10000000"
        CommSend PaperMaskRightControl, "D=10000000"
        CommSend PaperMaskRightControl, "VM=75000"
        CommSend PaperMaskRightControl, "S1=1,0"
        Sleep 250
        PaperMaskHome
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  WaitForPaperMask                                      **
'**                                                                        **
'**  Description..:  This routine waits for the stepper mask motors        **
'**                                                                        **
'****************************************************************************
Public Function WaitForPaperMask() As Boolean
    If DB.StepperMaskInstalled <> True Then Exit Function
    Dim mPC As New PerformanceCounter
    mPC.StartTimer True
    If Not DemoMode Then
        If DB.RS422MotorConfiguration = False Then
            '--- If using Quatech exit function if ports are not open
            If Comm(PaperMaskLeftControl).PortOpen = False Or Comm(PaperMaskRightControl).PortOpen = False Then
                GoTo Cleanup
            End If
        End If
        WaitForPaperMask = True          'We will assume the motor is moving, let it tell us otherwise
        PaperMaskRightStillMoving = True
        PaperMaskLeftStillMoving = True
        Do While PaperMaskLeftStillMoving = True Or PaperMaskRightStillMoving = True
            CommSend PaperMaskLeftControl, "PR MV"
            CommSend PaperMaskRightControl, "PR MV"
            Sleep 75
            DoEvents                                            'Handle windows events while we wait
        Loop
    Else
        WaitForPaperMask = False
        PaperMaskRightStillMoving = False
        PaperMaskLeftStillMoving = False
    End If
Cleanup:
    AppLog DebugMsg, "WaitForPaperMask,Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    Set mPC = Nothing
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PaperMaskHome                                         **
'**                                                                        **
'**  Description..:  This routine homes the stepper mask motors            **
'**                                                                        **
'****************************************************************************
Public Sub PaperMaskHome()
    If DB.StepperMaskInstalled <> True Then Exit Sub
    If Not DemoMode Then
        If DB.RS422MotorConfiguration = True Then
            AppLog DebugMsg, "PaperMaskHome,Sending motors home via RS-422..."
            CommSend PaperMaskLeftControl, "MR=-5000"
            CommSend PaperMaskRightControl, "MR=5000"
            WaitForPaperMask
            CommSend PaperMaskLeftControl, "HM=3"
            CommSend PaperMaskRightControl, "HM=1"
            WaitForPaperMask
            CommSend PaperMaskLeftControl, "P=0"
            CommSend PaperMaskRightControl, "P=0"
        Else
            AppLog DebugMsg, "PaperMaskHome,Sending motors home via RS-232..."
            CommSend PaperMaskLeftControl, "MR=-5000"
            CommSend PaperMaskRightControl, "MR=5000"
            WaitForPaperMask
            CommSend PaperMaskLeftControl, "HM=3"
            CommSend PaperMaskRightControl, "HM=1"
            WaitForPaperMask
            CommSend PaperMaskLeftControl, "P=0"
            CommSend PaperMaskRightControl, "P=0"
        End If
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  HomeRightMask                                         **
'**                                                                        **
'**  Description..:  This routine moves right mask to home position.       **
'**                                                                        **
'****************************************************************************
Public Sub HomeRightMask()
    If DB.StepperMaskInstalled <> True Then Exit Sub
    CommSend PaperMaskRightControl, "MR=5000"
    WaitForPaperMask
    CommSend PaperMaskRightControl, "HM=1"
    WaitForPaperMask
    CommSend PaperMaskRightControl, "P=0"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  HomeLeftMask                                          **
'**                                                                        **
'**  Description..:  This routine moves left mask to home position.        **
'**                                                                        **
'****************************************************************************
Public Sub HomeLeftMask()
    If DB.StepperMaskInstalled <> True Then Exit Sub
    CommSend PaperMaskLeftControl, "MR=-5000"
    WaitForPaperMask
    CommSend PaperMaskLeftControl, "HM=3"
    WaitForPaperMask
    CommSend PaperMaskLeftControl, "P=0"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MoveLeftMask                                          **
'**                                                                        **
'**  Description..:  This routine moves left mask to absolute position.    **
'**                                                                        **
'****************************************************************************
Public Sub MoveLeftMask(MaskPosition As String)
    If DB.StepperMaskInstalled <> True Then Exit Sub
    CommSend PaperMaskLeftControl, "MA=" & MaskPosition
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MoveRightMask                                         **
'**                                                                        **
'**  Description..:  This routine moves right mask to absolute position.   **
'**                                                                        **
'****************************************************************************
Public Sub MoveRightMask(MaskPosition As String)
    If DB.StepperMaskInstalled <> True Then Exit Sub
    CommSend PaperMaskRightControl, "MA=" & MaskPosition
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetMaskStatus                                         **
'**                                                                        **
'**  Description..:  This routine returns home status for mask motors      **
'**                                                                        **
'****************************************************************************
Public Function GetMaskStatus(MaskNo As Integer) As Boolean
    If DB.StepperMaskInstalled <> True Then Exit Function
    If Not DemoMode Then CommSend MaskNo, "PRINT I1"
End Function


