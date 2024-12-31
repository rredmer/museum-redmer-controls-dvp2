VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form DiagnosticsForm 
   BorderStyle     =   0  'None
   Caption         =   "DVP-2"
   ClientHeight    =   13350
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   18900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13350
   ScaleWidth      =   18900
   ShowInTaskbar   =   0   'False
   Begin UltraGrid.SSUltraGrid Outputs 
      Height          =   11325
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   19976
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   72352788
      BorderStyle     =   5
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColScrollRegions=   "DiagnosticsForm.frx":0000
      Override        =   "DiagnosticsForm.frx":003E
      Appearance      =   "DiagnosticsForm.frx":00E0
      Caption         =   "Outputs"
   End
   Begin UltraGrid.SSUltraGrid Inputs 
      Height          =   11325
      Left            =   5610
      TabIndex        =   1
      Top             =   30
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   19976
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   72352788
      BorderStyle     =   5
      ScrollBars      =   0
      ViewStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColScrollRegions=   "DiagnosticsForm.frx":011C
      Override        =   "DiagnosticsForm.frx":015A
      Appearance      =   "DiagnosticsForm.frx":01FC
      Caption         =   "Inputs"
   End
End
Attribute VB_Name = "DiagnosticsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Digital VP-2                                              **
'**                                                                        **
'** Module.....: Hardware                                                  **
'**                                                                        **
'** Description: User Control to provide Digital I/O & Serial Interface.   **
'**                                                                        **
'** History....:                                                           **
'**    09/25/03 v1.00 RDR Implemented Class from existing code.            **
'**                                                                        **
'** (c) 2002-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Private Const BoardNum As Integer = 0                   'Measurement Computing PCI-DIO96 Board number (0=First Board)

Public MarkText As String
Public MarkTextRequired As Boolean
Public LastSize As String                              'The index into the exposure settings grid of the last exposure!
Public PanelConnected As Boolean                       'Set TRUE when board is functional
Public CyanBit12 As Integer                            'Digital I/O Bit for Cyan filter
Public CyanBit24  As Integer                           'Digital I/O Bit for Cyan filter
Public YellowBit12  As Integer                         'Digital I/O Bit for Yellow filter
Public YellowBit24 As Integer                          'Digital I/O Bit for Yellow filter
Public MagentaBit12 As Integer                         'Digital I/O Bit for Magenta filter
Public MagentaBit24 As Integer                         'Digital I/O Bit for Magenta filter
Public LampShutterBit12 As Integer                     'Digital I/O Bit for lamp shutter bit
Public LampShutterBit24 As Integer                     'Digital I/O Bit for lamp shutter bit
Public Lens11x14Bit As Integer                         'Digital I/O Bit for 11x14 Lens cylinder
Public Lens8x10Bit As Integer                          'Digital I/O Bit for 8x10 Lens cylinder (Nord Only)
Public IrisShutterBit24 As Integer                     'Digital I/O Bit for IRIS Shutter
Public LcdHBit As Integer                              'Digital I/O Bit for LCD Horizontal
Public LcdVBit As Integer                              'Digital I/O Bit for LCD Vertical
Public RotateBit As Integer                            'Digital I/O Bit for Rotate cylinder
Public RightPanShutterBit As Integer                   'Digital I/O Bit for right pan shutter bit
Public LeftPanShutterBit As Integer                    'Digital I/O Bit for left pan shutter bit
Public PunchDie As Integer
Public PunchExtend As Integer
Public PunchBit0 As Integer
Public PunchBit1 As Integer
Public PunchBit2 As Integer
Public PunchBit3 As Integer
Public PunchBit4 As Integer
Public PunchBit5 As Integer
Public PunchBit6 As Integer
Public PunchBit7 As Integer
Public DoorClosed As Integer
Public PressureRollerEngaged As Integer
Public PaperOut As Integer
Public PlattenLeft As Integer
Public PlattenRight As Integer
Public RightPanOpen As Integer
Public LeftPanOpen As Integer
Public LcdHome As Integer
Public LcdRotated As Integer
Public IrisOpen As Integer
Public Lens11x14Switch As Integer
Public Lens8x10Switch As Integer
Public LensShuttleRight As Integer
Public PunchExtended As Integer
Public PunchRetracted As Integer
Public PunchDisengaged As Integer
Public RightFlap As Integer
Public LeftFlapLarge As Integer
Public LeftFlapSmall As Integer
Public PlattenCylinderBit As Integer
Private LastAdv As Single
Private LastPreAdv As Single
Private InExposureRoutine As Boolean                    'Used to avoid recursion in exposure routine
Private PC As PerformanceCounter

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UserControl_Initialize                                **
'**                                                                        **
'**  Description..:  This routine initializes DemoMode.                    **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Set PC = New PerformanceCounter                     'Initialize performance counter
    InExposureRoutine = False                           'Used to avoid recursion in exposure routine
    MarkTextRequired = False
    MarkText = ""
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Form_Load", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PC = Nothing
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  InitializeHardware                                    **
'**                                                                        **
'**  Description..:  This routine initializes all printer         **
'**                                                                        **
'****************************************************************************
Public Function InitializeHardware() As Boolean
    On Error GoTo ErrorHandler
    Dim HwdPC As New PerformanceCounter, SQLcmd As String
    HwdPC.StartTimer True
    Dim LcdDevName As String
    
    MapDigitalIO                               'Map Variables to Digital I/O
    
    Set Outputs.DataSource = DB.rsOutputs
    Outputs.Refresh ssRefetchAndFireInitializeRow
    Outputs.Bands(0).Columns.Add
    Outputs.Bands(0).Columns(0).Hidden = True
    Outputs.Bands(0).Columns(1).Hidden = True
    Outputs.Bands(0).Columns(2).Header.Caption = "Output"
    Outputs.Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
    Outputs.Bands(0).Columns(2).AutoEdit = False
    Outputs.Bands(0).Columns(2).Width = 2800
    Outputs.Bands(0).Columns(3).Header.Caption = "On/Off"
    Outputs.Bands(0).Columns(3).Width = 650
    Outputs.Bands(0).Columns(3).DataType = ssDataTypeBoolean
    
    Set Inputs.DataSource = DB.rsInputs
    Inputs.Refresh ssRefetchAndFireInitializeRow
    Inputs.Bands(0).Columns.Add
    Inputs.Bands(0).Columns(0).Hidden = True
    Inputs.Bands(0).Columns(1).Hidden = True
    Inputs.Bands(0).Columns(2).Header.Caption = "Input"
    Inputs.Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
    Inputs.Bands(0).Columns(2).AutoEdit = False
    Inputs.Bands(0).Columns(2).Width = 3700
    Inputs.Bands(0).Columns(3).Header.Caption = "Enabled"
    Inputs.Bands(0).Columns(3).Width = 800
    Inputs.Bands(0).Columns(5).Header.Caption = "On/Off"
    Inputs.Bands(0).Columns(5).Width = 650
    Inputs.Bands(0).Columns(5).Activation = ssActivationActivateNoEdit
    Inputs.Bands(0).Columns(5).AutoEdit = False
    Inputs.Bands(0).Columns(5).DataType = ssDataTypeBoolean
    
    If Not DemoMode Then
        AppLog InfoMsg, "InitializeHardware,Connecting to Panel..."
        PanelConnect                                        'Connect to digital I/O using Measurement Computing DLL
        LcdDevName = "\\.\Display2"
        AppLog InfoMsg, "InitializeHardware,Setting LCD Output Device to " & LcdDevName
        SetDeviceName LcdDevName                            'Open the second display for outputting the image - set device name
        AppLog InfoMsg, "InitializeHardware,Opening LCD device..."
        OpenOutputDevice                                    'Open the second display for outputting the image
    End If
    
    AppLog InfoMsg, "InitializeHardware,Timed," & Format(HwdPC.StopTimer, "####.####") & " seconds."
    Set HwdPC = Nothing
    Exit Function
ErrorHandler:
    HwdPC.StopTimer
    ErrorForm.ReportError Me.Name & ":InitializeHardware", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PanelConnect                                          **
'**                                                                        **
'**  Description..:  This routine configures the Digital I/O (Panel).      **
'**                                                                        **
'****************************************************************************
Public Sub PanelConnect()
    On Error GoTo ErrorHandler
    Dim PortNum As Integer, Status As Integer, BoardName As String * BOARDNAMELEN
    PanelConnected = False                              'Default to board not available
    BoardName = ""                                      'Clear the board name
    Status = cbGetBoardName(0, BoardName)               'Call Get Board Name function to see if installed
    If Left(BoardName, 4) = "DEMO" Then                 'If no board name returned, then its not installed
        AppLog ErrorMsg, "PanelConnect,Error connecting to PCI-DIO96 library."
    Else                                                'Else board is installed, configure it
        Status = cbDeclareRevision(CURRENTREVNUM)       'Must set library revision level first
        If Status = 0 Then
            Status = cbErrHandling(DONTPRINT, DONTSTOP)
            If Status <> 0 Then
                AppLog ErrorMsg, "PanelConnect,Error " & Status & " setting PCI-DIO96 error handling."
                Exit Sub
            End If
            For PortNum = 10 To 14
                Status = cbDConfigPort(BoardNum, PortNum, DIGITALOUT)
                If Status <> 0 Then
                    AppLog ErrorMsg, "PanelConnect,Error " & Status & " setting PCI-DIO96 outputs."
                    Exit Sub
                End If
                DoEvents
            Next PortNum
            For PortNum = 15 To 17
                Status = cbDConfigPort(BoardNum, PortNum, DIGITALIN)
                If Status <> 0 Then
                    AppLog ErrorMsg, "PanelConnect,Error " & Status & " setting PCI-DIO96 inputs."
                    Exit Sub
                End If
                DoEvents
            Next PortNum
            PanelConnected = True                       'If we got here then everything is ok, set PanelConnected flag!
        Else
            AppLog ErrorMsg, "PanelConnect,Error " & Status & " setting PCI-DIO96 library revision."
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":PanelConnect", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Outputs_CellChange                                    **
'**                                                                        **
'**  Description..:  This routine toggles outputs through grid interface.  **
'**                                                                        **
'****************************************************************************
Public Sub Outputs_CellChange(ByVal Cell As UltraGrid.SSCell)
    On Error GoTo ErrorHandler
    Dim Row As Integer, BitValue As Integer, Status As Integer
    BitValue = Cell.GetText
    Row = DB.rsOutputs.Fields("OutputNumber").Value
    AppLog DebugMsg, "Outputs_CellChange,Row " & Row & "," & DB.rsOutputs.Fields("Description").Value & "," & IIf(BitValue = 1, "On", "Off")
    If PanelConnected = True Then                                   'Only toggle bit if panel is connected
        Status = cbDBitOut(BoardNum, FIRSTPORTA, Row - 1, BitValue)
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Outputs_CellChange", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckSensesInPrintLoop                                **
'**                                                                        **
'**  Description..:  This routine checks the enabled printer senses        **
'**                                                                        **
'****************************************************************************
Public Function CheckSensesInPrintLoop() As Boolean
    Dim SenseStatus As Boolean
    SenseStatus = True
    ScanInputs                                  'Update digital I/O screen
    If IsInputEnabled(DoorClosed) = True Then
        If IsInputON(DoorClosed) = False Then
            MsgBox "The door is open.", vbInformation + vbOKOnly + vbApplicationModal, "Error"
            AppLog ErrorMsg, "CheckSensesInPrintLoop,The door is open."
            SenseStatus = False
        End If
    End If
    If IsInputEnabled(PressureRollerEngaged) = True Then
        If IsInputON(PressureRollerEngaged) = True Then
            MsgBox "The Pressure Roller is disengaged.", vbInformation + vbOKOnly + vbApplicationModal, "Error"
            AppLog ErrorMsg, "CheckSensesInPrintLoop,The pressure roller is disengaged."
            SenseStatus = False
        End If
    End If
    If IsInputEnabled(PlattenLeft) = True Then
        If IsInputON(PlattenLeft) = False Then
            MsgBox "The left-side of the paper platten is not set properly.", vbInformation + vbOKOnly + vbApplicationModal, "Error"
            AppLog ErrorMsg, "CheckSensesInPrintLoop,The left-side of the paper platten is not set properly."
            SenseStatus = False
        End If
    End If
    If IsInputEnabled(PlattenRight) = True Then
        If IsInputON(PlattenRight) = False Then
            MsgBox "The right-side of the paper platten is not set properly.", vbInformation + vbOKOnly + vbApplicationModal, "Error"
            AppLog ErrorMsg, "CheckSensesInPrintLoop,The right-side of the paper platten is not set properly."
            SenseStatus = False
        End If
    End If
    If IsInputEnabled(PaperOut) = True Then
        If DB.StepperMaskInstalled = True Then
            If IsInputON(PaperOut) = False Then
                MsgBox "The printer is out of paper.", vbInformation + vbOKOnly + vbApplicationModal, "Error"
                AppLog ErrorMsg, "CheckSensesInPrintLoop,The printer is out of paper."
                SenseStatus = False
            End If
        Else
            If IsInputON(PaperOut) = True Then
                MsgBox "The printer is out of paper.", vbInformation + vbOKOnly + vbApplicationModal, "Error"
                AppLog ErrorMsg, "CheckSensesInPrintLoop,The printer is out of paper."
                SenseStatus = False
            End If
        End If
    End If
    AppLog DebugMsg, "CheckSensesInPrintLoop," & SenseStatus
    CheckSensesInPrintLoop = SenseStatus
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetInputDelay                                         **
'**                                                                        **
'**  Description..:  This routine returns input delay from the database.   **
'**                                                                        **
'****************************************************************************
Public Function GetInputDelay(InputNum As Integer) As Long
    On Error GoTo ErrorHandler
    Dim TmpDelay As Long
    With DB.rsInputs
        .MoveFirst
        .Find "InputNumber=" & InputNum
        If Not .EOF Then
            TmpDelay = CLng(.Fields("TimeOut").Value)
            AppLog DebugMsg, "GetInputDelay," & InputNum & "," & TmpDelay
        Else
            TmpDelay = 0
            AppLog DebugMsg, "GetInputDelay," & InputNum & ",Input not defined."
        End If
    End With
    GetInputDelay = TmpDelay
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":GetInputDelay", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  IsInputEnabled                                        **
'**                                                                        **
'**  Description..:  This routine returns input enabled from the database. **
'**                                                                        **
'****************************************************************************
Public Function IsInputEnabled(InputNum As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim EnabledState As Boolean
    With DB.rsInputs
        .MoveFirst
        .Find "InputNumber=" & InputNum
        If Not .EOF Then
            EnabledState = .Fields("InputEnabled").Value
            'AppLog DebugMsg, "IsInputEnabled," & InputNum & "," & .Fields("Description").Value & "," & EnabledState
        Else
            EnabledState = False
            'AppLog DebugMsg, "IsInputEnabled," & InputNum & ",Input not defined."
        End If
    End With
    IsInputEnabled = EnabledState
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":IsInputEnabled", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ScanInputs                                            **
'**                                                                        **
'**  Description..:  This routine updates digital inputs.                  **
'**                                                                        **
'****************************************************************************
Public Sub ScanInputs()
    If PanelConnected = False Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Status As Integer, BitNum As Integer, BitValue As Integer, BookMarkedRecord As Variant
    BookMarkedRecord = DB.rsInputs.Bookmark
    For BitNum = 33 To 48
        Status = cbDBitIn(0, FIRSTPORTA, CInt(BitNum - 1), BitValue)
        With DB.rsInputs
            .MoveFirst
            .Find "InputNumber=" & (BitNum - 32)
            If Not .EOF Then
                Inputs.ActiveRow.Cells(5).Value = IIf(BitValue = 1, True, False)
            Else
                AppLog ErrorMsg, "ScanInputs,Input " & BitNum - 32 & " not found."
            End If
        End With
    Next
    DB.rsInputs.Bookmark = BookMarkedRecord
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ScanInputs", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckForInput                                         **
'**                                                                        **
'**  Description..:  This routine checks for input status & timing.        **
'**                                                                        **
'****************************************************************************
Public Function CheckForInput(SettingNum As Integer, InputState As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim mPC As New PerformanceCounter
    mPC.StartTimer True
    If Not DemoMode Then
        Dim InputTimeOut As Long, ExpectedTimeOut As Long, SettingName As String, startTime As Single
        startTime = Timer
        With DB.rsInputs
            .MoveFirst
            .Find "InputNumber=" & SettingNum
            If Not .EOF Then
                SettingName = .Fields("Description").Value
                ExpectedTimeOut = CLng(.Fields("TimeOut").Value)
            End If
        End With
        InputTimeOut = 0
        If IsInputEnabled(SettingNum) = True Then
            Do While IsInputON(SettingNum) <> InputState
                ScanInputs
                Sleep 1
                InputTimeOut = InputTimeOut + 1
                If InputTimeOut > ExpectedTimeOut Then
                    AppLog ErrorMsg, "CheckForInput," & SettingName & " input timeout."
                    MsgBox SettingName & " timed out at " & ExpectedTimeOut & "ms, please check air pressure.", vbInformation + vbOKOnly + vbApplicationModal, "Error"
                    CheckForInput = False
                    GoTo Cleanup
                End If
            Loop
            AppLog DebugMsg, "CheckForInput," & SettingName & ",Timed," & Format(mPC.StopTimer, "####.####")
        Else                                                'Sense Disabled, the timeout period is simply a wait delay.
            AppLog DebugMsg, "CheckForInput," & SettingName & " disabled, waiting for " & ExpectedTimeOut & "ms."
            Sleep ExpectedTimeOut
        End If
        CheckForInput = True
    End If
Cleanup:
    mPC.StopTimer
    Set mPC = Nothing
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":CheckForInput", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  IsInputOn                                             **
'**                                                                        **
'**  Description..:  This routine returns input hardware status.           **
'**                                                                        **
'****************************************************************************
Public Function IsInputON(InputNum As Integer) As Boolean
    On Error GoTo ErrorHandler
    If PanelConnected = False Then
        Exit Function
    End If
    Dim InputState As Boolean, BookMarkedRecord As Variant
    BookMarkedRecord = DB.rsInputs.Bookmark
    With DB.rsInputs
        .MoveFirst
        .Find "InputNumber=" & InputNum
        If Not .EOF Then
            InputState = Inputs.ActiveRow.Cells(5).Value
            'Log DebugMsg, "IsInputOn," & InputNum & "," & InputState
        Else
            InputState = False
            'Log DebugMsg, "IsInputOn," & InputNum & ",Input not defined."
        End If
        .Bookmark = BookMarkedRecord
    End With
    IsInputON = InputState
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":IsInputOn", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  BitON                                                 **
'**                                                                        **
'**  Description..:  This routine turns on specified digital output.       **
'**                                                                        **
'****************************************************************************
Public Sub BitON(OutputNum As Integer)
    If PanelConnected = False Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    PC.StartTimer True
    Dim Status As Integer, startTime As Single
    startTime = Timer
    With DB.rsOutputs
        .MoveFirst
        .Find "OutputNumber=" & OutputNum
        If Not .EOF Then
            Outputs.ActiveRow.Cells(3).Value = True
            Status = cbDBitOut(BoardNum, FIRSTPORTA, OutputNum - 1, 1)
            'Log DebugMsg, "BitOn," & OutputNum & "," & Format(PC.StopTimer, "####.####") & " sec."
        Else
            'Log DebugMsg, "BitOn," & OutputNum & ",Output not defined."
        End If
    End With
    PC.StopTimer
    Exit Sub
ErrorHandler:
    PC.StopTimer
    ErrorForm.ReportError Me.Name & ":BitON", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  BitOFF                                                **
'**                                                                        **
'**  Description..:  This routine turns off specified digital output.      **
'**                                                                        **
'****************************************************************************
Public Sub BitOFF(OutputNum As Integer)
    If PanelConnected = False Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Status As Integer
    PC.StartTimer True
    With DB.rsOutputs
        .MoveFirst
        .Find "OutputNumber=" & OutputNum
        If Not .EOF Then
            Outputs.ActiveRow.Cells(3).Value = False
            Status = cbDBitOut(BoardNum, FIRSTPORTA, OutputNum - 1, 0)
            'Log DebugMsg, "BitOff," & OutputNum & "," & Format(PC.StopTimer, "####.####") & " sec."
        Else
            'Log DebugMsg, "BitOff," & OutputNum & ",Output not defined."
        End If
    End With
    PC.StopTimer
    Exit Sub
ErrorHandler:
    PC.StopTimer
    ErrorForm.ReportError Me.Name & ":BitOFF", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AllFiltersOn                                          **
'**                                                                        **
'**  Description..:  This routine enables all filters.                     **
'**                                                                        **
'****************************************************************************
Public Sub AllFiltersOn()
    '--- Enable All Filters  24v
    Dim mPC As New PerformanceCounter
    mPC.StartTimer True
    AppLog DebugMsg, "AllFiltersOn,Enabling 24v on all Filters (Bi-Level Hi)"
    BitON CyanBit24
    BitON MagentaBit24
    BitON YellowBit24
    Refresh
    Sleep 20
    AppLog DebugMsg, "AllFiltersOn,Enabling 12v on all Filters (Bi-Level Lo)"
    BitON CyanBit12
    BitON MagentaBit12
    BitON YellowBit12
    AppLog DebugMsg, "AllFiltersOn,Disabling 24v on all Filters (Bi-Level OFF)"
    BitOFF CyanBit24
    BitOFF MagentaBit24
    BitOFF YellowBit24
    AppLog DebugMsg, "AllFiltersOn,Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    Set mPC = Nothing
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AllFiltersOff                                         **
'**                                                                        **
'**  Description..:  This routine disables all filters.                    **
'**                                                                        **
'****************************************************************************
Public Sub AllFiltersOff()
    Dim mPC As New PerformanceCounter
    mPC.StartTimer True
    AppLog DebugMsg, "AllFiltersOff,Disabling 12v on all Filters (All Filters OFF)"
    BitOFF CyanBit12
    BitOFF MagentaBit12
    BitOFF YellowBit12
    Refresh
    AppLog DebugMsg, "AllFiltersOff,Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    Set mPC = Nothing
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MapDigitalIO                                          **
'**                                                                        **
'**  Description..:  This routine maps digital I/O by printer type.        **
'**                                                                        **
'****************************************************************************
Public Sub MapDigitalIO()
    If DB.StepperMaskInstalled = True Then
        '---- DVP2
        CyanBit12 = 1                   'Digital I/O Bit for Cyan filter
        CyanBit24 = 2                   'Digital I/O Bit for Cyan filter
        YellowBit12 = 3                 'Digital I/O Bit for Yellow filter
        YellowBit24 = 4                 'Digital I/O Bit for Yellow filter
        MagentaBit12 = 5                'Digital I/O Bit for Magenta filter
        MagentaBit24 = 6                'Digital I/O Bit for Magenta filter
        LampShutterBit12 = 7            'Digital I/O Bit for lamp shutter bit
        LampShutterBit24 = 8            'Digital I/O Bit for lamp shutter bit
        Lens11x14Bit = 9                'Digital I/O Bit for 11x14 Lens cylinder
        IrisShutterBit24 = 10           'Digital I/O Bit for IRIS Shutter
        LcdHBit = 11                    'Digital I/O Bit for LCD Horizontal
        LcdVBit = 12                    'Digital I/O Bit for LCD Vertical
        RotateBit = 13                  'Digital I/O Bit for Rotate cylinder
        RightPanShutterBit = 14         'Digital I/O Bit for right pan shutter bit
        LeftPanShutterBit = 15          'Digital I/O Bit for left pan shutter bit
        PunchDie = 16
        PunchExtend = 17
        PunchBit0 = 18
        PunchBit1 = 19
        PunchBit2 = 20
        PunchBit3 = 21
        PunchBit4 = 22
        PunchBit5 = 23
        PunchBit6 = 24
        PunchBit7 = 25
        DoorClosed = 1                  'Digital Inputs Start Here!
        PressureRollerEngaged = 2
        PaperOut = 3
        PlattenLeft = 4
        PlattenRight = 5
        RightPanOpen = 6
        LeftPanOpen = 7
        LcdHome = 8
        LcdRotated = 9
        IrisOpen = 10
        Lens11x14Switch = 11
        Lens8x10Switch = 12
        LensShuttleRight = 13
        PunchExtended = 14
        PunchRetracted = 15
        PunchDisengaged = 16
        PlattenCylinderBit = 26
    Else
        '---- NORD
        CyanBit12 = 1                   'Digital I/O Bit for Cyan filter
        CyanBit24 = 2                   'Digital I/O Bit for Cyan filter
        YellowBit12 = 3                 'Digital I/O Bit for Yellow filter
        YellowBit24 = 4                 'Digital I/O Bit for Yellow filter
        MagentaBit12 = 5                'Digital I/O Bit for Magenta filter
        MagentaBit24 = 6                'Digital I/O Bit for Magenta filter
        LampShutterBit12 = 7            'Digital I/O Bit for lamp shutter bit
        LampShutterBit24 = 8            'Digital I/O Bit for lamp shutter bit
        IrisShutterBit24 = 10           'Digital I/O Bit for IRIS Shutter
        LcdHBit = 11                    'Digital I/O Bit for LCD Horizontal
        LcdVBit = 12                    'Digital I/O Bit for LCD Vertical
        RotateBit = 13                  'Digital I/O Bit for Rotate cylinder
        RightFlap = 14
        LeftFlapLarge = 21
        LeftFlapSmall = 20
        Lens8x10Bit = 17
        Lens11x14Bit = 18               'Digital I/O Bit for 11x14 Lens cylinder
        PlattenCylinderBit = 19
        
        '--- These variables are not used in Nord, simply initialize to unused/harmless bit #
        PunchExtend = 25
        RightPanShutterBit = 25         'Digital I/O Bit for right pan shutter bit
        LeftPanShutterBit = 25          'Digital I/O Bit for left pan shutter bit
        PunchDie = 25
        PunchBit0 = 25
        PunchBit1 = 25
        PunchBit2 = 25
        PunchBit3 = 25
        PunchBit4 = 25
        PunchBit5 = 25
        PunchBit6 = 25
        PunchBit7 = 25
        DoorClosed = 1                  'Digital Inputs Start Here!
        PressureRollerEngaged = 1
        
        PaperOut = 2
        PlattenLeft = 4
        PlattenRight = 5
        RightPanOpen = 6
        LeftPanOpen = 7
        LcdHome = 8
        LcdRotated = 9
        IrisOpen = 10
        Lens11x14Switch = 11
        Lens8x10Switch = 12
        LensShuttleRight = 13
        PunchExtended = 14
        PunchRetracted = 15
        PunchDisengaged = 16
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ExposeImage                                           **
'**                                                                        **
'**  Description..:  This routine makes the actual exposure.               **
'**                                                                        **
'****************************************************************************
Public Sub ExposeImage(RedTime As Long, GreenTime As Long, BlueTime As Long)
    Dim mPC As New PerformanceCounter
    mPC.StartTimer True
    AppLog DebugMsg, "ExposeImage,Quadrant Exposure Times R=" & RedTime & ", G=" & GreenTime & ", B=" & BlueTime
    CheckForInput RightPanOpen, True                    'Make sure the pan shutter is open
    
    BitON LcdHBit                                       'Energize H Bit per specification
    BitON LcdVBit                                       'Energize V Bit per specification
    
    Sleep 60                                            'Leave the H&V Activated for 60ms to wake them up
    
    BitOFF LcdHBit                                      'De-energize H Bit per specification
    BitOFF LcdVBit                                      'De-energize V Bit per specification
    OutputFrame 0                                       'Make sure first frame is on LCD prior to opening up!
    MakeColor 1                                         'Put filters in for RED
    
    BitON IrisShutterBit24                              'Open IRIS Shutter
    CheckForInput IrisOpen, True                        'Make sure IRIS opened... Delay should be no more than 50ms or Color Shift will occur
    BitON LampShutterBit24                              'Fire the Lamp Shutter 24v - this is the start of the exposure
    BitON LampShutterBit12                              'Fire the Lamp Shutter 12v - this is the hold voltage (Bi-Level)
    ExposeColor 1, RedTime                              'Make Red exposure
    BitOFF LampShutterBit24                             'Bi-Level the Lamp Shutter
    MakeColor 2
    ExposeColor 2, GreenTime                            'Make Green exposure
    MakeColor 3
    ExposeColor 3, BlueTime                             'Make Blue exposure
    BitOFF LampShutterBit24                             'Close Lamp Shutter - this is the end of the exposure
    BitOFF LampShutterBit12                             'Close Lamp Shutter - this is the end of the exposure
    BitOFF IrisShutterBit24                             'Close IRIS Shutter
    
    Sleep 50
    BitOFF CyanBit12                                    'Open Cyan Filter
    BitOFF MagentaBit12                                 'Open Magenta Filter
    BitOFF YellowBit12                                  'Open Yellow Filter
    BitOFF LcdHBit
    BitOFF LcdVBit
    CheckForInput IrisOpen, False                       'Make sure IRIS is not Stuck Open
    AppLog DebugMsg, "ExposeImage,Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    Set mPC = Nothing
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MakeColor                                             **
'**                                                                        **
'**  Description..:  This routine pulls in filters to provide RGB Color.   **
'**                                                                        **
'****************************************************************************
Public Sub MakeColor(ColorPlane As Long)
    AllFiltersOn
    Select Case ColorPlane
        Case 1                                          'RED = YELLOW+MAGENTA
            AppLog DebugMsg, "MakeColor,Setting RED by disabling Cyan."
            BitOFF CyanBit12
        Case 2                                          'GREEN = CYAN + YELLOW
            AppLog DebugMsg, "MakeColor,Setting GREEN by disabling Magenta."
            BitOFF MagentaBit12
        Case 3                                          'BLUE = CYAN + MAGENTA
            AppLog DebugMsg, "MakeColor,Setting BLUE by disabling Yellow."
            BitOFF YellowBit12
    End Select
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ExposeColor                                           **
'**                                                                        **
'**  Description..:  This routine exposes single color using the LCD.      **
'**                                                                        **
'****************************************************************************
Public Sub ExposeColor(ColorPlane As Long, delaytime As Long)
    Dim mPC As New PerformanceCounter
    mPC.StartTimer True
    If Not DemoMode Then
        On Error GoTo ErrorHandler
        Dim FrameIndex As Integer
        FrameIndex = (ColorPlane - 1) * 4
        OutputFrame FrameIndex + 0
        
        BitOFF LcdHBit
        BitOFF LcdVBit
        
        'log  DebugMsg, vbTab & ColorString & " position " & FrameIndex + 0 & " @ " & Str(DelayTime) & "ms."
        Sleep delaytime                                     'Wait for exposure time
        OutputFrame FrameIndex + 1
        
        BitON LcdHBit
        BitOFF LcdVBit
        
        'log  DebugMsg, vbTab & ColorString & " position " & FrameIndex + 1 & " @ " & Str(DelayTime) & "ms."
        Sleep delaytime                                     'Wait for exposure time
        OutputFrame FrameIndex + 2
        
        BitON LcdHBit
        BitON LcdVBit
        
        'log  DebugMsg, vbTab & ColorString & " position " & FrameIndex + 2 & " @ " & Str(DelayTime) & "ms."
        Sleep delaytime                                     'Wait for exposure time
        OutputFrame FrameIndex + 3
        
        BitOFF LcdHBit
        BitON LcdVBit
        
        'log  DebugMsg, vbTab & ColorString & " position " & FrameIndex + 3 & " @ " & Str(DelayTime) & "ms."
        Sleep delaytime                                     'Wait for exposure time
        BitOFF LcdHBit
        BitOFF LcdVBit
    
    Else
        Sleep 4 * delaytime                                 'Simulate exposure delay in demo mode
        AppLog InfoMsg, "ExposeColor, demo mode exposure..."
    End If
    AppLog DebugMsg, "ExposeColor,Timed," & Format(mPC.StopTimer, "####.####") & ",ExpTime," & 4 * delaytime
    Set mPC = Nothing
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ExposeColor", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MakePunch                                             **
'**                                                                        **
'**  Description..:  This routine activates the punch.                     **
'**                                                                        **
'****************************************************************************
Public Sub MakePunch(PunchCode As Byte)
    Dim mPC As New PerformanceCounter
    mPC.StartTimer True
    Dim BitNum As Integer, BitVal As Byte
    AppLog DebugMsg, "Punching Code=" & PunchCode
    BitON PunchExtend                                   'Extend punch
    Sleep DB.PunchPkgExtendTime                         'Time for package punch to extend
    For BitNum = 0 To 7                                 'Set the bit pattern
        BitVal = PunchCode And (2 ^ BitNum)
        If BitVal = (2 ^ BitNum) Then
            BitON PunchBit0 + BitNum                    'Turn on the necessary punch bit
            'AppLog DebugMsg, "Enabled Punch Bit " & Str(PunchBit0 + BitNum)
        End If
    Next
    Sleep DB.PunchSolenoidTime                          'Time for bit solenoids to engage
    BitON PunchDie                                      'Make the punch
    Sleep DB.PunchEngageTime                            'Time for punch to engage
    BitOFF PunchDie
    For BitNum = PunchBit0 To PunchBit7                 'Clear all punch bits
        BitOFF BitNum
    Next
    Sleep DB.PunchDisengageTime                         'Time for solenoids to disengage
    BitOFF PunchExtend                                  'Retract the punch
    Sleep DB.PunchPkgExtendTime                         'Time for Package Punch to Retract
    AppLog DebugMsg, "MakePunch,Code=" & Str(PunchCode) & ",Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    Set mPC = Nothing
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  MakeExposure                                          **
'**                                                                        **
'**  Description..:  This routine makes an exposure using the LCD.         **
'**                  Returns time consumed in ms, -1 if error.             **
'**                                                                        **
'****************************************************************************
Public Function MakeExposure(ImageFileName As String, SizeString As String, ManualMode As Boolean, Text1 As String, Text2 As String, PunchCode As Byte) As Currency
    If InExposureRoutine = True Then Exit Function
    InExposureRoutine = True
    On Error GoTo ErrorHandler
    Dim mPC As New PerformanceCounter, RedTime As Long, GreenTime As Long, BlueTime As Long, TimedValue As Currency
    MakeExposure = -1
    MainForm.DiagnosticTimer.Enabled = False                            'No diagnostic timer during this routine....
    mPC.StartTimer True                                                 'Total Exposure Time maintained with Performance Counter Class (Windows API Calls)
    AppLog InfoMsg, "MakeExposure,Starting Exposure," & ImageFileName & "," & SizeString
    If PrepareToPrintImage(ImageFileName, SizeString, True, PunchCode) <> -1 Then 'Prepare image and machine (advance, mask, rotate, punch) for printing
        If DB.rsPrintSizes.Fields("Lens11x14").Value = False Then       'The Proper PrintSize record is now positioned from PrepareToPrintImage Call
            RedTime = CLng(CLng(DB.rsExposureTime.Fields("Red8x10").Value) / 4)
            GreenTime = CLng(CLng(DB.rsExposureTime.Fields("Green8x10").Value) / 4)
            BlueTime = CLng(CLng(DB.rsExposureTime.Fields("Blue8x10").Value) / 4)
        Else
            RedTime = CLng(CLng(DB.rsExposureTime.Fields("Red11x14").Value) / 4)
            GreenTime = CLng(CLng(DB.rsExposureTime.Fields("Green11x14").Value) / 4)
            BlueTime = CLng(CLng(DB.rsExposureTime.Fields("Blue11x14").Value) / 4)
        End If
        If DB.PlatenCylinderEnable = True Then              'If the platten has a cylinder on it
            BitOFF PlattenCylinderBit                       'Make sure the platten is down prior to making an exposure
        End If
        MotorDiagnostics.WaitForPaperMask                   'Simply wait for mask motors to stop
        MotorDiagnostics.WaitForPaperAdvance                'Simply wait for advance motor to stop
        ShowCursor 0                                        'Hide the mouse cursor
        ExposeImage RedTime, GreenTime, BlueTime            'Make the exposure (Hardware Interface)
        ShowCursor -1                                       'Show the mouse cursor
        If FileSystemHandle.FileExists(CurrentPrintFile) = True Then
            Kill CurrentPrintFile
        End If
        
        If DB.StepperMaskInstalled = False Then
            BitOFF LeftFlapSmall
            BitOFF LeftFlapLarge
            BitOFF RightFlap
        End If
    End If
    
    DB.StatTotalExposures = DB.StatTotalExposures + 1
    'DB.StatTimeRunning = Timer - startTime
    MainForm.DiagnosticTimer.Enabled = True
    InExposureRoutine = False                           'Clear recursion flag
    TimedValue = mPC.StopTimer
    AppLog DebugMsg, "MakeExposure,Timed," & Format(TimedValue, "####.####")
    MakeExposure = TimedValue
    Set mPC = Nothing
    Exit Function
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
    ErrorForm.ReportError Me.Name & ":MakeExposure", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PrepareToPrintImage                                   **
'**                                                                        **
'**  Description..:  This routine updates the preview image w/ corrections.**
'**                  Returns time consumed in ms, -1 if error.             **
'**                                                                        **
'****************************************************************************
Public Function PrepareToPrintImage(FileName As String, SizeString As String, MakePaperAdvance As Boolean, PunchCode As Byte) As Currency
    On Error GoTo ErrorHandler
    Dim mPC As PerformanceCounter, CurAdv As Single, PreAdv As Single, PunchDist As Single, TimedValue As Currency
    Dim FrameNo As Integer, Position As Integer, lutname As String
    Dim RedLutFile As String, GreenLutFile As String, BlueLutFile As String, RedOffsetFile As String, GreenOffsetFile As String, BlueOffsetFile As String
    Set mPC = New PerformanceCounter
    mPC.StartTimer True                                         'Start the performance timer
    PrepareToPrintImage = -1                                    'Default function to failure
    
    '---- Backwriting needs to be adjusted to start of print....
    AppLog InfoMsg, "PrepareToPrintImage,Configuring," & FileName & "," & SizeString & ",Advance=" & MakePaperAdvance
    
    With DB.rsPrintSizes
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "PrintSize='" & SizeString & "'", 0, adSearchForward, 0
            If Not .EOF Then
                '---- Make paper advance, if necessary
                If MakePaperAdvance = True Then
                    CurAdv = .Fields("PaperFeed").Value         'Get current advance length
                    PreAdv = .Fields("PaperPreFeed").Value      'Get current pre-feed
                    PunchDist = .Fields("PaperPunchFeed").Value 'Get current punch feed (moved outside LastSize block to solve punch issue)
                    If LastSize <> "" Then                      'If a previous exposure was made
                        LastAdv = LastAdv + (LastPreAdv - PreAdv) - PunchDist 'Get the last print out of the way (to the punch)
                        AppLog DebugMsg, "PrepareToPrintImage,Advancing Paper,LastSize=" & LastSize & ",CurAdv=" & CurAdv & ",PreAdv=" & PreAdv & ",PunchDist=" & PunchDist & ",MarkReq=" & IIf(Me.MarkTextRequired, "Yes", "No") & ",Text=" & Me.MarkText
                        AppLog DebugMsg, "PrepareToPrintImage,Advancing " & LastAdv & " inches"
                        MotorDiagnostics.AdvancePaper LastAdv   'Primary Advance
                        DB.StatPaperUsed = DB.StatPaperUsed + (LastAdv / 12)
                    Else
                        AppLog DebugMsg, "PrepareToPrintImage,No Last Size Defined!"
                    End If
                    LastSize = .Fields("PrintSize").Value
                Else
                    AppLog DebugMsg, "PrepareToPrintImage,Not making paper advance."
                End If
                
                '---- Set the paper mask
                If DB.StepperMaskInstalled = True Then
                    MotorDiagnostics.WaitForPaperMask           'Make sure mask is not moving or next command will lock up the machine
                    SizeSettingsForm.LeftMaskPosition.Caption = .Fields("LeftMaskPosition").Value
                    SizeSettingsForm.RightMaskPosition.Caption = .Fields("RightMaskPosition").Value
                    If SizeSettingsForm.LeftMaskPosition.Caption < DB.MaskLeftLimit Then
                        AppLog ErrorMsg, "PrepareToPrintImage,Left Mask Out Of Range " & SizeSettingsForm.LeftMaskPosition.Caption & " Limit=" & DB.MaskLeftLimit
                        SizeSettingsForm.LeftMaskPosition.Caption = DB.MaskLeftLimit
                    End If
                    If SizeSettingsForm.RightMaskPosition.Caption > DB.MaskRightLimit Then
                        AppLog ErrorMsg, "PrepareToPrintImage,Right Mask Out Of Range " & SizeSettingsForm.RightMaskPosition.Caption & " Limit=" & DB.MaskRightLimit
                        SizeSettingsForm.RightMaskPosition.Caption = DB.MaskRightLimit
                    End If
                    MotorDiagnostics.MoveLeftMask SizeSettingsForm.LeftMaskPosition.Caption
                    MotorDiagnostics.MoveRightMask SizeSettingsForm.RightMaskPosition.Caption
                Else
                    '---- Enable Large Flap if set
                    If .Fields("LeftFlapLarge").Value = True Then
                        BitON LeftFlapSmall
                        Sleep 100
                        BitON LeftFlapLarge
                    Else
                        BitOFF LeftFlapLarge
                        BitOFF LeftFlapSmall
                    End If
                    '---- Enable Small Flap is set
                    If .Fields("LeftFlapSmall").Value = True Then
                        BitON LeftFlapSmall
                    Else
                        BitOFF LeftFlapSmall
                    End If
                    '---- Enable Right Flap if set
                    If .Fields("RightFlap").Value = True Then
                        '--- Turn Mask bit on for right
                        BitON RightFlap
                    Else
                        BitOFF RightFlap
                    End If

                End If
                
                '---- Set the Lens according to spreadsheet
                If DB.StepperMaskInstalled = True Then
                    If .Fields("Lens11x14").Value = True Then
                        BitON Lens11x14Bit
                    Else
                        BitOFF Lens11x14Bit
                    End If
                Else
                    '---- NORD LENS CONFIGURATION
                    If .Fields("Lens11x14").Value = True Then
                        BitOFF Lens8x10Bit
                        Sleep 100
                        BitON Lens11x14Bit
                    Else
                        BitOFF Lens11x14Bit
                        Sleep 100
                        BitON Lens8x10Bit
                    End If
                End If
                
                '---- Rotate the deck if necessary
                If .Fields("RotateTable").Value = True Then
                    BitON RotateBit
                Else
                    BitOFF RotateBit
                End If
            Else
                AppLog ErrorMsg, "PrepareToPrintImage,Print size not defined for " & FileName
                mPC.StopTimer
                Set mPC = Nothing
                Exit Function
            End If
        Else
            MsgBox "No print sizes defined."
            Exit Function
        End If
    End With
    
    SizeSettingsForm.ProcessImage FileName                      'This is where the image is processed for LCD
    SetImageFile MakeCstring(CurrentPrintFile)              'Set the image file to the current image
    
    '---- Path to LUT & Offset files
    If PrinterQcForm.PrintingLUT = True Then
        RedLutFile = FastSettingsFolder & "LUT\ClearLUTr.lut"
        GreenLutFile = FastSettingsFolder & "LUT\ClearLUTg.lut"
        BlueLutFile = FastSettingsFolder & "LUT\ClearLUTb.lut"
    Else
        RedLutFile = FastSettingsFolder & "LUT\lutr.lut"
        GreenLutFile = FastSettingsFolder & "LUT\lutg.lut"
        BlueLutFile = FastSettingsFolder & "LUT\lutb.lut"
    End If
    
    RedOffsetFile = FastSettingsFolder & "Offset\offsetr.frm"
    GreenOffsetFile = FastSettingsFolder & "Offset\offsetg.frm"
    BlueOffsetFile = FastSettingsFolder & "Offset\offsetb.frm"
    
    
    '---- Calculate Red Exposure
    lutname = RedLutFile
    If FileSystemHandle.FileExists(lutname) Then
        If FileSystemHandle.FileExists(RedOffsetFile) Then
            AppLog DebugMsg, "CalculateExposures,Calculating red exposure planes using LUT=" & lutname & " Offset=" & RedOffsetFile
            FrameNo = 0
            SetLutFile MakeCstring(lutname)
            SetOffsetFile MakeCstring(RedOffsetFile)
            SetColor 2
            For Position = 0 To 3
                SetPosition Position
                CalcFrame FrameNo
                FrameNo = FrameNo + 1
            Next
        Else
            AppLog ErrorMsg, "Red Offset File not found," & RedOffsetFile
            MsgBox "Red Offset File not found," & RedOffsetFile, vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
            GoTo Cleanup
        End If
    Else
        AppLog ErrorMsg, "Red LUT File not found," & RedLutFile
        MsgBox "Red LUT File not found," & RedLutFile, vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
        GoTo Cleanup
    End If
    
    '---- Calculate Green Exposure
    lutname = GreenLutFile
    If FileSystemHandle.FileExists(lutname) Then
        If FileSystemHandle.FileExists(GreenOffsetFile) Then
            AppLog DebugMsg, "CalculateExposures,Calculating green exposure planes using LUT=" & lutname & " Offset=" & GreenOffsetFile
            SetLutFile MakeCstring(lutname)
            SetOffsetFile MakeCstring(GreenOffsetFile)
            SetColor 1
            For Position = 0 To 3
                SetPosition Position
                CalcFrame FrameNo
                FrameNo = FrameNo + 1
            Next
        Else
            AppLog ErrorMsg, "Green Offset File not found," & GreenOffsetFile
            MsgBox "green Offset File not found," & GreenOffsetFile, vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
            GoTo Cleanup
        End If
    Else
        AppLog ErrorMsg, "Green LUT File not found," & GreenLutFile
        MsgBox "green LUT File not found," & GreenLutFile, vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
        GoTo Cleanup
    End If

    '---- Calculate Blue Exposure
    lutname = BlueLutFile
    If FileSystemHandle.FileExists(lutname) Then
        If FileSystemHandle.FileExists(BlueOffsetFile) Then
            AppLog DebugMsg, "CalculateExposures,Calculating blue exposure planes using LUT=" & lutname & " Offset=" & BlueOffsetFile
            SetLutFile MakeCstring(lutname)
            SetOffsetFile MakeCstring(BlueOffsetFile)
            SetColor 0
            For Position = 0 To 3
                SetPosition Position
                CalcFrame FrameNo
                FrameNo = FrameNo + 1
            Next
        Else
            AppLog ErrorMsg, "Blue Offset File not found," & BlueOffsetFile
            MsgBox "Blue Offset File not found," & BlueOffsetFile, vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
            GoTo Cleanup
        End If
    Else
        AppLog ErrorMsg, "Blue LUT File not found," & BlueLutFile
        MsgBox "Blue LUT File not found," & BlueLutFile, vbApplicationModal + vbOKOnly + vbCritical, "ERROR"
        GoTo Cleanup
    End If
    
    If MakePaperAdvance = True Then
        If DB.PunchEnable = True Or DB.PackagePunchEnable = True Then
            MotorDiagnostics.WaitForPaperAdvance
            AppLog DebugMsg, "PrepareToPrintImage,Punching paper..."
            If PunchCode > 0 Then
                MakePunch CByte(PunchCode)
            End If
        End If
        If PunchDist <> 0 Then
            AppLog DebugMsg, "PrepareToPrintImage,Advancing " & PunchDist & " inches for Punch"
            MotorDiagnostics.AdvancePaper PunchDist         'This is the punch feed (Wait for stop is in MakeExposure)
        End If
        LastAdv = CurAdv                                    'Store current advance setting as last advance
        LastPreAdv = PreAdv                                 'Store current pre-advance as last pre-advance
    End If
    
    If DB.rsPrintSizes.Fields("Lens11x14").Value = True Then
        CheckForInput Lens11x14Switch, True
    Else
        CheckForInput Lens8x10Switch, True
    End If
    If DB.rsPrintSizes.Fields("RotateTable").Value = True Then      'Make sure the table is in position
        CheckForInput LcdRotated, True
    Else
        CheckForInput LcdHome, True
    End If
    
Cleanup:
    TimedValue = mPC.StopTimer
    PrepareToPrintImage = TimedValue
    AppLog DebugMsg, "PrepareToPrintImage,Timed," & Format(TimedValue, "####.####") & " seconds."
    Set mPC = Nothing
    DB.StatTotalImages = DB.StatTotalImages + 1
    DB.StatAverageServerTime = TimedValue
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":PrepareToPrintImage", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

