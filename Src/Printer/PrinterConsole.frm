VERSION 5.00
Object = "{00130063-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltdlg13n.ocx"
Object = "{00130003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltocx13n.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form PrinterConsole 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   13335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13335
   ScaleWidth      =   17640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   13215
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   17595
      _ExtentX        =   31036
      _ExtentY        =   23310
      _Version        =   262144
      PaneTree        =   "PrinterConsole.frx":0000
      Begin UltraGrid.SSUltraGrid PrintQueGrid 
         Height          =   5655
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   17535
         _ExtentX        =   30930
         _ExtentY        =   9975
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   1
         LayoutFlags     =   72351748
         BorderStyle     =   5
         Override        =   "PrinterConsole.frx":0072
         Appearance      =   "PrinterConsole.frx":00F0
         Caption         =   "PrintQueGrid"
      End
      Begin Threed.SSFrame CurrentImageFrame 
         Height          =   7410
         Left            =   30
         TabIndex        =   2
         Top             =   5775
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   13070
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
         Caption         =   "Current Image"
         Begin LEADLib.LEAD ImagePreview 
            Height          =   7005
            Left            =   60
            TabIndex        =   3
            Top             =   210
            Width           =   8565
            _Version        =   65540
            _ExtentX        =   15108
            _ExtentY        =   12356
            _StockProps     =   229
            BorderStyle     =   1
            ScaleHeight     =   465
            ScaleWidth      =   569
            DataField       =   ""
            BitmapDataPath  =   ""
            AnnDataPath     =   ""
            PaintSizeMode   =   3
            PanWinTitle     =   "PanWindow"
            CLeadCtrl       =   0
            LoadCompressed  =   0
         End
      End
   End
   Begin VB.Timer QueTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   60
      Top             =   12780
   End
   Begin ActiveToolBars.SSActiveToolBars MainToolBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   12
      Tools           =   "PrinterConsole.frx":012C
      ToolBars        =   "PrinterConsole.frx":993F
   End
   Begin LEADDlgLibCtl.LEADDlg LEADDlg 
      Left            =   480
      Top             =   12780
      Angle           =   0
      AngleFlag       =   0   'False
      NewWidth        =   0
      NewHeight       =   0
      MaxFileSize     =   0
      LoadCompressed  =   -1  'True
      LoadRotated     =   -1  'True
      Change          =   0
      SaveMulti       =   0
      PageNumber      =   1
      Effect          =   2000
      Grain           =   5
      Delay           =   20
      MaxPass         =   1
      Transparent     =   0   'False
      WandWidth       =   2
      GradientStyle   =   1000
      GradientSteps   =   256
      Transition      =   1000
      Shape           =   1000
      ShapeBackStyle  =   1
      ShapeFillStyle  =   0
      ShapeBorderStyle=   1
      ShapeBorderWidth=   1
      ShapeInnerStyle =   0
      ShapeInnerWidth =   0
      ShapeOuterStyle =   0
      ShapeOuterWidth =   0
      ShadowXDepth    =   5
      ShadowYDepth    =   5
      SampleText      =   "LEADTOOLS!"
      TextStyle       =   1
      TextAlign       =   4
      TextWordWrap    =   -1  'True
      TextUseForegroundImage=   0   'False
      FileDlgFlags    =   0
      FileDialogTitle =   "LEADTOOLS Common Dialog"
      FileName        =   ""
      Filter          =   ""
      FilterIndex     =   0
      InitialDir      =   ""
      UIFlags         =   0
      ShowHelpButton  =   0   'False
      PreviewEnabled  =   -1  'True
      EnableMethodErrors=   -1  'True
      LowBit          =   0
      HighBit         =   0
      LowValue        =   0
      HighValue       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   255
      ForeColor       =   16711680
      StartColor      =   255
      EndColor        =   16711680
      TransparentColor=   0
      WandColor       =   255
      ShapeBorderColor=   0
      ShapeInnerHiliteColor=   16777215
      ShapeInnerShadowColor=   0
      ShapeOuterHiliteColor=   16777215
      ShapeOuterShadowColor=   0
      ShadowColor     =   0
      TextColor       =   16711680
      TextHiliteColor =   16777215
      Directory       =   ""
      BeginProperty DlgFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubTypeIndex    =   0
   End
End
Attribute VB_Name = "PrinterConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Digital VP-2                                              **
'**                                                                        **
'** Module.....: PrintQue                                                  **
'**                                                                        **
'** Description: User Control to provide Print Queue Processing.           **
'**              The Print Queue data is maintained in a seperate database **
'**              connection than the main printer tables.  This provides   **
'**              the ability to split this data into it's own local file   **
'**              on the printer or maintain it seperately on the SQLServer.**
'**              The Print Queue data is extremely dynamic by nature, and  **
'**              therefore must be managed more often than the other data. **
'**                                                                        **
'** History....:                                                           **
'**    09/25/03 v1.00 RDR Implemented Class from existing code.            **
'**                                                                        **
'** (c) 2002-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Public FirstPrint As Boolean
Private StopPrinterButtonPressed As Boolean             'Provides for Stopping Print Queue
Private ProcessingPrintQue As Boolean                   'Prevents recurrsion in print routine
Private Const ControlFileSection As String = "Order_Item"
Private Const ControlFileQuantity As String = "Quantity"
Private Const ControlFileProdCode As String = "Product_Code"
Private Const ControlFilePunchCode As String = "Punch_Code"
Private Const ControlFileBackPrint1 As String = "BackPrint_1"
Private Const ControlFileBackPrint2 As String = "BackPrint_2"
Private Const ControlFileBackPrint3 As String = "BackPrint_3"
Private Const ControlFileBackPrint4 As String = "BackPrint_4"
Private Const ControlFileImageName As String = "FileName"
Private ProcessingQueue As Boolean                      'Flag to avoid recurrsion in print que processor
Private BackWriterBuffer As String

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine connects to recordsets.                  **
'**                                                                        **
'****************************************************************************
Public Sub Setup()
    On Error GoTo ErrorHandler
    Dim Col As Integer
    AppLog InfoMsg, "PrintQue:Setup,configuring grid controls"
    
    DiagnosticsForm.LastSize = ""                                       'Just started - no last size printed!!
    BackWriterBuffer = ""
    
    ProcessingPrintQue = False                                          'Used to avoid recursion in print que processor
    StopPrinterButtonPressed = False                                    'Indicates if STOP printing button pressed
    
    '--- The Print Queue Grid
    With PrintQueGrid
        Set .DataSource = DB.rsPrintQue                                 'Set datasource for Grid
        .Refresh ssRefetchAndFireInitializeRow                          'Refresh grid with data
        .Bands(0).Columns(0).Hidden = True                              'This is the PrinterName column
        .Bands(0).Columns(1).Width = 8000                               'Image File Name Column
        .Bands(0).Columns(1).Activation = ssActivationActivateNoEdit
        .Bands(0).Columns(2).Activation = ssActivationActivateNoEdit    'Print Size
        .Bands(0).Columns(3).Activation = ssActivationAllowEdit         'Quantity
        .Bands(0).Columns(4).Hidden = True                              'PixelsX
        .Bands(0).Columns(5).Hidden = True                              'PixelsY
        .Bands(0).Columns(6).Activation = ssActivationActivateNoEdit    'Control File
        .Bands(0).Columns(7).Activation = ssActivationAllowEdit         'Print Enable
        .Bands(0).Columns(8).Activation = ssActivationAllowEdit         'Hold Enable
        .Bands(0).Columns(9).Activation = ssActivationActivateNoEdit    'Status
        For Col = 10 To .Bands(0).Columns.Count - 1
            .Bands(0).Columns(Col).Hidden = True
        Next
        .Override.SelectTypeRow = ssSelectTypeExtended
        .Override.MaxSelectedRows = 1000
    End With

    UserMode_Click 0, 0                                 'Initialize user mode to simple
    
    EnableToolBarAfterPrinting                          'Initialize Tool Bar - not printing on startup!
    
    AppLog InfoMsg, "PrintQue:Setup,Starting Que Timer..."
    QueTimer.Enabled = True                             'Start the Print Que Timer

    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintQue:Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UserControl_Terminate                                 **
'**                                                                        **
'**  Description..:  This routine closes recordsets on normal exit.        **
'**                                                                        **
'****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler
    QueTimer.Enabled = False
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrinterConsole:Form_Unload", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PrintQueToolbar_ButtonClick                           **
'**                                                                        **
'**  Description..:  This routine handles the print que tool bar.          **
'**                                                                        **
'****************************************************************************
Private Sub MainToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Dim FileName As String, SizeString As String
    
    Select Case Tool.ID
        Case "ID_StartPrinting"
            If Tool.Name = StartPrintingMessage Then
                If ProcessingPrintQue = False Then                         'This will avoid recurrsion!
                    Tool.Name = StopPrintingMessage
                    StopPrinterButtonPressed = False
                    DoEvents
                    ProcessQue
                    Tool.Name = StartPrintingMessage
                End If
            Else
                StopPrinterButtonPressed = True
                MsgBox "Stopping printer production"
                AppLog DebugMsg, "QueButton_Click,Stopping Print Que Processing..."
                Tool.Name = StartPrintingMessage
            End If
        Case "ID_StartServer"
            If Tool.Name = StartServerMessage Then
                Tool.Name = StopServerMessage
                AppLog DebugMsg, "QueButton_Click,Starting Server..."
                RunQueTimer True
            Else
                AppLog DebugMsg, "QueButton_Click,Stopping Server..."
                RunQueTimer False
                Tool.Name = StartServerMessage
            End If
        Case "ID_RemovePaper"
            With MotorDiagnostics
                If DB.PaperAdvanceTearOffLength < 50 Then           'Check for tear off < 50in.
                    .AdvancePaper DB.PaperAdvanceTearOffLength
                Else
                    AppLog ErrorMsg, "TearOffButton_Click,Tear Off Length too long (>50) " & DB.PaperAdvanceTearOffLength
                    .AdvancePaper 33                                'Provide standard tear off length
                End If
                .WaitForPaperAdvance
                .PaperMotorTorqueOFF
            End With
            If DB.PlatenCylinderEnable = True Then
                DiagnosticsForm.BitON DiagnosticsForm.PlattenCylinderBit
            End If
            
            DiagnosticsForm.LastSize = ""
            
        Case "ID_AdvancePaper"
            Dim length As Single
            length = DB.PaperAdvanceLength
            With MotorDiagnostics
                .PaperMotorTorqueON
                .AdvancePaper IIf(length < 40, length, 40)          'Advance the paper - no more than 20 inches
                .WaitForPaperAdvance
                Sleep 500
            End With
        Case "ID_Add"
            With LEADDlg                                            'LEAD DiagnosticsForm.log Control
                .FileName = ""
                .DialogTitle = "Add image to print queue"
                .PreviewEnabled = True
                .FILTER = "All |*.*|BMP|*.bmp|CMP|*.cmp|JPEG|*.jpg|PSD|*.psd|TGA|*.tga|TIF|*.tif"
                .FilterIndex = 2
                .LoadPasses = 0
                .LoadRotated = False
                .LoadCompressed = False
                .Bitmap = 0                                         'free any existing bitmap reference
                .FileDlgFlags = DLG_OFN_NOCHANGEDIR
                .EnableMethodErrors = False
                .EnableLongFilenames = True
                .UIFlags = DLG_FO_RESDLG + DLG_FO_FILEINFO + DLG_FO_SHOWSTAMP + DLG_FO_SHOWPREVIEW
                .FileOpen (hWnd)
                If .FileName <> "" Then
                    AddImageToPrintQue "Manual", 0, .FileName
                End If
            End With
        Case "ID_Erase"
            With DB.rsPrintQue
                If MsgBox("Remove selected images from print queue?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            If PrintQueGrid.ActiveRow.Selected = True Then
                                FileName = Trim(.Fields("ImageFileName").Value)
                                If FileSystemHandle.FileExists(FileName) Then
                                    Kill FileName
                                End If
                                .Delete adAffectCurrent
                            End If
                            .MoveNext
                        Loop
                    End If
                End If
            End With
            PrintQueGrid.Refresh ssRefetchAndFireInitializeRow
        Case "ID_Refresh"
            If MsgBox("Are you sure?", vbApplicationModal + vbDefaultButton2 + vbQuestion + vbYesNo, "Refresh print queue?") = vbYes Then
                AppLog DebugMsg, "Refreshing Print Queue..."
                With DB.rsPrintQue
                    If .RecordCount > 0 Then
                    .MoveFirst
                        Do While Not .EOF
                            .Delete adAffectCurrent
                            .MoveNext
                        Loop
                    End If
                End With
                PrintQueGrid.Refresh ssRefetchAndFireInitializeRow
                With DB.rsHotFolders
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            AppLog DebugMsg, "Setting HotFolder [" & .Fields("HotFolderPath").Value & "] to 0 images for refresh."
                            .Fields("NumberOfFiles").Value = 0
                            .UpdateBatch adAffectCurrent
                            .MoveNext
                        Loop
                    Else
                        AppLog ErrorMsg, "Refreshing Print Queue... no hot folder records to reset."
                    End If
                End With
            End If
        Case "ID_Preview"
            FileName = DB.rsPrintQue.Fields("ImageFileName").Value
            SizeString = DB.rsPrintQue.Fields("PrintSize").Value
            If FileSystemHandle.FileExists(FileName) = True Then
                AppLog DebugMsg, "User Preview of [" & FileName & "]."
                If DiagnosticsForm.PrepareToPrintImage(FileName, SizeString, False, 0) <> -1 Then
                    ImageTimeOut = 0
                    WatchForImageTimeOut = True
                End If
            End If
        Case "ID_HoldOn", "ID_HoldOff"
            With DB.rsPrintQue
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If PrintQueGrid.ActiveRow.Selected = True Then
                            .Fields("HoldEnable").Value = IIf(Tool.ID = "ID_HoldOn", -1, 0)
                        End If
                        .MoveNext
                    Loop
                End If
            End With
            PrintQueGrid.Refresh ssRefetchAndFireInitializeRow
        Case "ID_PrintOn", "ID_PrintOff"
            With DB.rsPrintQue
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If PrintQueGrid.ActiveRow.Selected = True Then
                            .Fields("PrintEnable").Value = IIf(Tool.ID = "ID_PrintOn", -1, 0)
                        End If
                        .MoveNext
                    Loop
                End If
            End With
            PrintQueGrid.Refresh ssRefetchAndFireInitializeRow
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintQue:PrintQueToolbar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UserMode_Click                                        **
'**                                                                        **
'**  Description..:  This routine handles Maintenance mode w/ Security.    **
'**                                                                        **
'****************************************************************************
Private Sub UserMode_Click(index As Integer, Value As Integer)
    Dim TabNum As Integer
    If index = 0 Then
        '--- Maintenance Mode
        If Value = ssCBUnchecked Then
        '    If TabControl.Tabs(2).Visible = True Then
        '        For TabNum = 2 To TabControl.Tabs.Count
        '            TabControl.Tabs(TabNum).Enabled = False
        '            TabControl.Tabs(TabNum).Visible = False
        '        Next
        '    End If
        Else
            Dim GoodPassword As Boolean
            GoodPassword = False
            If DB.AdminPassword <> "" Then
                PasswordForm.Show vbModal
                If Trim(UCase(DB.AdminPassword)) = Trim(UCase(PasswordForm.Password.Text)) Then
                    GoodPassword = True
                Else
                    MsgBox "Access denied.", vbInformation + vbApplicationModal + vbOKOnly, "Wrong password"
                    
                    'UserMode(0).Value = ssCBUnchecked
                    
                End If
            Else
                GoodPassword = True
            End If
            If GoodPassword = True Then
                'For TabNum = 2 To TabControl.Tabs.Count
                '    TabControl.Tabs(TabNum).Enabled = True
                '    TabControl.Tabs(TabNum).Visible = True
                'Next
            End If
        End If
    Else
        '--- Q.C. Mode
        If Value = ssCBUnchecked Then
            PrinterConsole.RunQueTimer True
            'QueButton(0).Enabled = True
            'QueButton(1).Enabled = True
        Else
            PrinterConsole.RunQueTimer True
            'PrintQue.Visible = False
            'QueButton(0).Enabled = False                'No Que printing in Q.C. Mode
            'QueButton(1).Enabled = False
        End If
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ProcessQue                                            **
'**                                                                        **
'**  Description..:  This is the one and only print que processing routine!**
'**                                                                        **
'****************************************************************************
Private Sub ProcessQue()
    If ProcessingPrintQue = True Then                          'Avoid recurrsion - it would be fatal to call this twice!
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Qty As Integer, ExpNum As Integer, ImageFileName As String, SizeString As String, BookmarkRecord As Variant
    Dim Text1 As String, Text2 As String, PunchCode As Byte
    Dim mPC As New PerformanceCounter
    PrinterConsole.RunQueTimer False                    'Turn Off Queue Processor - queues do not process while printing!
    DB.RemovePrintedImages                              'Make sure printed images are not left in the print queue
    AppLog DebugMsg, "ProcessQue,=========================================================================="
    AppLog DebugMsg, "ProcessQue,Starting Print Que Processing...."
    
    DisableToolBarForPrinting
    
    SizeSettingsForm.DisableSetupSettings                   'Make sure focus marks, caption text, and pixel ruler are OFF prior to printing batch!
    
    DiagnosticsForm.BitON DiagnosticsForm.RightPanShutterBit
    DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, True
    
    
    '---- This needs to be a configurable option.
    'MotorDiagnostics.AdvancePaper 14                    'Pull paper to ensure no double exposures and working on clean paper!
    
    
    FirstPrint = True
    
    ProcessingPrintQue = True
    With DB.rsPrintQue
        If .RecordCount = 0 Then
            AppLog ErrorMsg, "ProcessQue,No images in print queue to print"
            MsgBox "No images in print queue.", vbApplicationModal + vbOKOnly + vbInformation, "Error"
            ProcessingPrintQue = False
            Exit Sub
        End If
        mPC.StartTimer True
        .MoveFirst
        Do While Not .EOF                               'Loop for every record in the print queue (which may be added to along the way!)
            BookmarkRecord = .Bookmark
            AppLog DebugMsg, "ProcessQue,=== At top of print loop"
            If DiagnosticsForm.CheckSensesInPrintLoop() = False Then
                Exit Do
            End If
            If StopPrinterButtonPressed = True Then     'If the user pressed the STOP button
                AppLog ErrorMsg, "ProcessQue,The Stop Printer button was pressed."
                Exit Do
            End If
            If .Fields("PrintEnable").Value = -1 Then   'If the Print Enable flag is set
                ImageFileName = Trim(.Fields("ImageFileName").Value)
                If ImageFileName <> "" Then
                    If FileSystemHandle.FileExists(ImageFileName) Then
                        AppLog DebugMsg, "ProcessQue,Preparing to print " & ImageFileName
                        .Fields("Status").Value = "Exposing..."
                        
                        SizeString = .Fields("PrintSize").Value             'Validated when Queue loaded
                        Qty = .Fields("PrintQuantity").Value                'Validated when Queue loaded
                        Text1 = .Fields("BackWriterText1").Value
                        Text2 = .Fields("BackWriterText2").Value
                        PunchCode = .Fields("PunchCode").Value
                        
                        For ExpNum = 1 To Qty
                            MainForm.StatusLabel(0).Caption = "Exposure " & ExpNum & " of " & Qty
                            
                            DiagnosticsForm.MakeExposure ImageFileName, SizeString, False, Text1, Text2, PunchCode
                            
                            DoEvents
                            If StopPrinterButtonPressed = True Then       'If the user pressed the STOP button
                                .Fields("Status").Value = "Stopped."
                                GoTo Cleanup
                            End If
                        Next
                        If .Fields("HoldEnable").Value = 0 Then
                            .Fields("Status").Value = "Printed."
                            DB.RemoveImageFromHotFolder .Fields("PrintQue").Value, ImageFileName
                        Else
                            .Fields("Status").Value = "Holding."
                        End If
                    Else
                        '---- Move to File Error
                        .Fields("Status").Value = "File Error."
                    End If
                Else
                    '---- Move to File Error
                    .Fields("Status").Value = "File Error."
                End If
            Else
                
            End If
            .UpdateBatch adAffectCurrent
            PrinterConsole.CheckForNewImages            'Add new images to print que - this may change the current record in rsPrintQueue if more files are added
            .Bookmark = BookmarkRecord                  'Make sure we are on the current record we just printed
            
            If DB.BackWritersInstalled = True Then
                If FirstPrint = False Then
                    AppLog DebugMsg, "ProcessQue,FirstPrint=" & FirstPrint & ",Backwriter Text=" & BackWriterBuffer & ",Enabling=" & True
                    DiagnosticsForm.MarkText = BackWriterBuffer
                    DiagnosticsForm.MarkTextRequired = True
                Else
                    AppLog DebugMsg, "ProcessQue,FirstPrint=" & FirstPrint & ",Backwriter Text=" & BackWriterBuffer & ",Enabling=" & False
                    DiagnosticsForm.MarkTextRequired = False
                End If
            End If
            
            BackWriterBuffer = Text1
            AppLog DebugMsg, "ProcessQue,Looping for next image..."
            
            If Not .EOF Then
                .MoveNext
            Else
                'At the last exposure!
                DiagnosticsForm.MarkText = BackWriterBuffer         'Store the final mark text to the buffer to be written on tear out or next advance.
                
            End If
                            
            FirstPrint = False
            Me.Refresh
        Loop
    End With
    AppLog DebugMsg, "ProcessQue,=========================================================================="
Cleanup:
    '--- Clean up
    
    'MotorDiagnostics.AdvancePaper 10                     'Pull paper to ensure no double exposures and working on clean paper!
    
    AppLog DebugMsg, "ProcessQue,Closing Pan Shutter..."
    ProcessingPrintQue = False                                 'Clear recurrsion flag
    DiagnosticsForm.BitOFF DiagnosticsForm.RightPanShutterBit
    DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, False
    MotorDiagnostics.PaperMotorTorqueOFF
    SizeSettingsForm.ClearImage                                          'Put blank image to LCD
    AppLog DebugMsg, "ProcessQue,Completing Print Que Processing..."
    DoEvents
    DB.RemovePrintedImages                              'remove printed images from the queue
    MainForm.StatusLabel(0).Caption = PrinterIdleMessage
    PrinterConsole.RunQueTimer True                           'Start the Queue Processor
    EnableToolBarAfterPrinting
    AppLog DebugMsg, "ProcessQue,Timed," & Format(mPC.StopTimer, "####.####")
    Set mPC = Nothing
    Exit Sub
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
    ErrorForm.ReportError Me.Name & ":ProcessQue", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  QueTimer_Timer                                        **
'**                                                                        **
'**  Description..:  Process Queue Timer.                                  **
'**                                                                        **
'****************************************************************************
Private Sub QueTimer_Timer()
    On Error GoTo ErrorHandler
    If ProcessingQueue = True Then                  'If already processing the queue
        Exit Sub                                    'Get out of here!  Avoid recurrsion
    End If
    CheckForNewImages
    DB.UpdateStatistics
    'PrinterConsole.PrinterStatistics.RefreshStatistics
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintQue:QueTimer_Timer", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  QueTimer_Timer                                        **
'**                                                                        **
'**  Description..:  Process Queue Timer.                                  **
'**                                                                        **
'****************************************************************************
Public Sub RunQueTimer(RunMode As Boolean)
    QueTimer.Enabled = RunMode
End Sub

Public Function GetTimerEnabled() As Boolean
    GetTimerEnabled = QueTimer.Enabled
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckForNewImages                                     **
'**                                                                        **
'**  Description..:  This routine checks hot folders for new images.       **
'**                                                                        **
'****************************************************************************
Public Function CheckForNewImages()
    If ProcessingQueue = True Then Exit Function
    ProcessingQueue = True                          'Set processing flag
    On Error GoTo ErrorHandler                      'Set error handler
    Dim Extension As String, AddedImage As Boolean, QuePath As String
    Dim FolderHandle As Scripting.Folder, TempFileHandle As Scripting.File 'FileSystem Pointers
    AddedImage = False
    With DB.rsHotFolders
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If .Fields("FolderEnabled").Value = True Then
                    QuePath = Trim(.Fields("HotFolderPath").Value) & "\"
                    If FileSystemHandle.FolderExists(QuePath) = True Then
                        Set FolderHandle = FileSystemHandle.GetFolder(QuePath)
                        '---- Only check hot folder if number of files has changed
                        If .Fields("NumberOfFiles").Value <> FolderHandle.Files.Count Then
                            'DiagnosticsForm.log DebugMsg, "CheckForNewImages,Refreshing from HotFolder=[" & QuePath & "] from " & .Fields("NumberOfFiles").Value & " files to " & FolderHandle.Files.Count & " files."
                            
                            
                            
                            '----- RDR Should only loop back from new file count to old file count!!
                            
                            
                            '---- Loop for each image file in the hotfolder
                            For Each TempFileHandle In FolderHandle.Files
                                Extension = UCase(Right(TempFileHandle.Name, 3))
                                If Extension <> "TXT" Then          'Skip Control Files
                                    If AddImageToPrintQue(QuePath, .Fields("FolderType").Value, Trim(QuePath) & Trim(TempFileHandle.Name)) = True Then
                                        AddedImage = True
                                    End If
                                End If
                            Next
                            
                            
                            '----- RDR ERROR BELOW:
                            '-- If ANY of the files error out, they will not be scanned again into the queue!
                            
                            
                            .Fields("NumberOfFiles").Value = FolderHandle.Files.Count
                            .Fields("Status").Value = "Ok"
                            .UpdateBatch adAffectCurrent
                        End If
                    Else
                        ''DiagnosticsForm.log ErrorMsg, "CheckForNewImages,Hot Folder " & QuePath & " does not exist."
                        .Fields("Status").Value = "Error"
                        .UpdateBatch adAffectCurrent
                    End If
                End If
                .MoveNext
            Loop
        End If
        
        'PrinterConsole.QueButton(0).Enabled = IIf(DB.rsHotFolders.RecordCount > 0, True, False)
        
    End With
    If AddedImage = True Then                           'Sort the Print Que by Date/Time
        PrintQueGrid.Bands(0).Columns(5).SortIndicator = ssSortIndicatorAscending
    End If
    
    '---- If there are file errors, make that tab Red
    'If DB.rsFileErrors.RecordCount > 0 Then
    '    PrintQueTab.Tabs(3).BackColorSource = ssUseTab
    '    PrintQueTab.Tabs(3).BackColor = RGB(255, 0, 0)
    'Else
    '    PrintQueTab.Tabs(3).BackColorSource = ssUseControl
    '    PrintQueTab.Tabs(3).BackColor = &H8000000F
    'End If
    
'    PutStat StatTimeRunning, Format(Now - CDate(GetStat(StatTimeLastCleared)), "hh:mm:ss")
    Set TempFileHandle = Nothing
    Set FolderHandle = Nothing
    ProcessingQueue = False
    Exit Function
ErrorHandler:
    ProcessingQueue = False
    ErrorForm.ReportError "PrintQue:CheckForNewImages", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddImageToPrintQue                                    **
'**                                                                        **
'**  Description..:  This routine adds image data to the print queue. It is**
'**                  the only place where images are validated.            **
'**                                                                        **
'****************************************************************************
Public Function AddImageToPrintQue(QuePath As String, QueType As Integer, FileName As String) As Boolean
    On Error GoTo ErrorHandler
    With DB.rsPrintQue
        
        
        '---- RDR - SELECT IN DATABASE Class will be faster than FIND here
        
        
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "ImageFilename='" & FileName & "'"
            If Not .EOF Then
                Exit Function
            End If
        End If
        
        '-----------------------------------------------------------------
        
        
        Dim myFile As Scripting.File, SizeString As String, Qty As Integer, PunchCode As Integer, Text1 As String, Text2 As String, Text3 As String, Text4 As String
        Dim ExtensionPos As Integer, ControlFileName As String, TmpBuf As String, CharsRead As Long, HasControlFile As Boolean
        Set myFile = FileSystemHandle.GetFile(FileName)
        HasControlFile = False
        Qty = 1
        
        If GetImageData(QuePath, FileName) = False Then
            Exit Function
        End If
        
        SizeString = DB.GetPrintSize(PrinterConsole.ImagePreview.InfoHeight, PrinterConsole.ImagePreview.InfoWidth)
        If SizeString = "" Then
            'DiagnosticsForm.log ErrorMsg, "AddImageToPrintQue,Image " & FileName & " does not have a defined print size."
            '---- Add the Image to the FileErrors Table
            With DB.rsFileErrors
                If .RecordCount > 0 Then .MoveFirst
                .Find "ImageFileName='" & FileName & "'"
                If .EOF Then
                    .AddNew
                    .Fields("PrinterName").Value = PrinterName
                    .Fields("PrintQue").Value = QuePath
                    .Fields("ImageFileName").Value = FileName
                    .Fields("ErrorDescription").Value = "Print size not defined (" & PrinterConsole.ImagePreview.InfoHeight & "," & PrinterConsole.ImagePreview.InfoWidth & ") ratio=" & Format(PrinterConsole.ImagePreview.InfoWidth / PrinterConsole.ImagePreview.InfoHeight, "####.###")
                    .Fields("TimeOfError").Value = Now
                    .UpdateBatch adAffectCurrent
                End If
            End With
            Exit Function
        End If
        
        Select Case QueType
            Case 1                  'Render Que
                '---- Fuji code = QY
                If DB.QueGetQty = True Then                     'Quantity
                    If DB.FujiFileNameEncoding = True Then
                        Qty = Val(Trim(GetFujiTextFromFileName("QY", FileName)))
                    Else
                        Qty = GetQuantityFromFileName(FileName)
                    End If
                    If Qty < 1 Or Qty > 200 Then Qty = 1
                Else
                    Qty = 1
                End If
                
                If DB.QueGetPunch = True Then                   'Punch Code
                    PunchCode = GetPunchCodeFromFileName(FileName)
                    If PunchCode <= 0 Or PunchCode > 255 Then PunchCode = 1
                Else
                    PunchCode = 1
                End If
                
                If DB.BackWriterOverride = True Then
                    Text1 = DB.BW_Text1                         'Fixed text from backwriter settings
                    Text2 = DB.BW_Text2
                Else
                    If DB.FujiFileNameEncoding = True Then
                        Text1 = Trim(GetFujiTextFromFileName("PC", FileName)) & " " & Trim(GetFujiTextFromFileName("CN", FileName))
                        Text2 = Trim(GetFujiTextFromFileName("PC", FileName)) & " " & Trim(GetFujiTextFromFileName("CN", FileName))
                    Else
                        Text1 = Mid(FileName, InStrRev(FileName, "\", Len(FileName)) + 1, Len(FileName))
                        Text2 = Mid(FileName, InStrRev(FileName, "\", Len(FileName)) + 1, Len(FileName))
                    End If
                End If
                
                
            Case 2                  'Control File
                ExtensionPos = InStr(1, FileName, ".", vbTextCompare)
                If ExtensionPos > 0 Then
                    ControlFileName = Left(FileName, ExtensionPos) & "TXT"
                    If FileSystemHandle.FileExists(ControlFileName) = True Then
                        'DiagnosticsForm.log DebugMsg, "Reading control file..." & ControlFileName
                        TmpBuf = String(181, 0)
                        CharsRead = GetPrivateProfileString(ControlFileSection, ControlFileQuantity, "1", TmpBuf, 80, ControlFileName)
                        If CharsRead > 0 Then Qty = CInt(Left(TmpBuf, CharsRead))
                        CharsRead = GetPrivateProfileString(ControlFileSection, ControlFilePunchCode, "1", TmpBuf, 80, ControlFileName)
                        If CharsRead > 0 Then PunchCode = CInt(Left(TmpBuf, CharsRead))
                        GetPrivateProfileString ControlFileSection, ControlFileProdCode, "0", TmpBuf, 80, ControlFileName
                        GetPrivateProfileString ControlFileSection, ControlFileBackPrint1, "", TmpBuf, 180, ControlFileName
                        If CharsRead > 0 Then Text1 = CInt(Left(TmpBuf, CharsRead))
                        GetPrivateProfileString ControlFileSection, ControlFileBackPrint2, "", TmpBuf, 180, ControlFileName
                        If CharsRead > 0 Then Text2 = CInt(Left(TmpBuf, CharsRead))
                        GetPrivateProfileString ControlFileSection, ControlFileBackPrint3, "", TmpBuf, 180, ControlFileName
                        If CharsRead > 0 Then Text3 = CInt(Left(TmpBuf, CharsRead))
                        GetPrivateProfileString ControlFileSection, ControlFileBackPrint4, "", TmpBuf, 180, ControlFileName
                        If CharsRead > 0 Then Text4 = CInt(Left(TmpBuf, CharsRead))
                        GetPrivateProfileString ControlFileSection, ControlFileImageName, "", TmpBuf, 180, ControlFileName
                        HasControlFile = True
                    End If
                Else
                    AppLog ErrorMsg, "Error reading control file for image " & FileName
                    Qty = 1
                End If
        End Select
        '---- Add the image data to the ActivePrintQueue Table
        .AddNew
        .Fields("PrinterName").Value = PrinterName
        .Fields("ImageFileName").Value = FileName
        .Fields("PrintQue").Value = QuePath
        .Fields("PrintQueType").Value = QueType
        .Fields("PrintSize").Value = SizeString
        .Fields("PrintQuantity").Value = Qty                          'Print Quantity from File Name or Control File
        .Fields("PunchCode").Value = PunchCode                        'Punch Code from File Name or Control File
        .Fields("BackWriterText1").Value = Text1                      'Text from File Name or Control File
        .Fields("BackWriterText2").Value = Text2                      'Text from File Name or Control File
        .Fields("BackWriterText3").Value = Text3                      'Text from File Name or Control File
        .Fields("BackWriterText4").Value = Text4                      'Text from File Name or Control File
        .Fields("ImageResolution").Value = PrinterConsole.ImagePreview.InfoXRes
        .Fields("PixelsX").Value = CLng(PrinterConsole.ImagePreview.InfoHeight)
        .Fields("PixelsY").Value = CLng(PrinterConsole.ImagePreview.InfoWidth)
        .Fields("ColorDepth").Value = PrinterConsole.ImagePreview.InfoBits
        .Fields("Compression").Value = PrinterConsole.ImagePreview.InfoCompress
        .Fields("DiskSize").Value = PrinterConsole.ImagePreview.InfoSizeDisk
        .Fields("ControlFile").Value = IIf(HasControlFile = True, -1, 0)
        .Fields("PrintEnable").Value = -1
        .Fields("HoldEnable").Value = 0
        .Fields("Status").Value = "Waiting."
        .Fields("FileDate").Value = myFile.DateLastModified
        .UpdateBatch adAffectCurrent
        .Resync adAffectAllChapters
        PrinterConsole.ImagePreview.Bitmap = 0
        
        
        AppLog DebugMsg, "Added " & FileName & " to print que using: Qty=" & Str(Qty) & " Punch=" & Str(PunchCode) & " Size=" & SizeString & " Txt1=" & Text1
        
    End With
    Exit Function
ErrorHandler:
    ErrorForm.ReportError "PrintQue:AddImageToPrintQueue", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetImageData                                          **
'**                                                                        **
'**  Description..:  This routine verifies the file is ready/gets file info**
'**                                                                        **
'****************************************************************************
Private Function GetImageData(QuePath As String, FileName As String) As Boolean
    On Error Resume Next                    'Required to continue this routine on error
    Dim ReadTry As Integer, ReturnVal As Integer
    GetImageData = False
    ReadTry = 1
    Do While ReadTry <= 30                  'If it takes longer than 3 seconds, give it up!
        ReturnVal = PrinterConsole.ImagePreview.GetFileInfo(FileName, 0, FILEINFO_TOTALPAGES)
        If ReturnVal <> 0 Then
            '---- Problem reading bitmap information, file is most likely still being copied or not an image file
            Sleep 200                       'File is not ready, check again in 100ms
        Else
            'Check for Valid Image Format
            If PrinterConsole.ImagePreview.InfoHeight = 0 Or PrinterConsole.ImagePreview.InfoWidth = 0 Then
                Sleep 100
            Else
                If PrinterConsole.ImagePreview.InfoHeight > 2400 Or PrinterConsole.ImagePreview.InfoWidth > 3400 Then
                    With DB.rsFileErrors
                        If .RecordCount > 0 Then .MoveFirst
                        .Find "ImageFileName='" & FileName & "'"
                        If .EOF Then
                            .AddNew
                            .Fields("PrinterName").Value = PrinterName
                            .Fields("PrintQue").Value = QuePath
                            .Fields("ImageFileName").Value = FileName
                            .Fields("ErrorDescription").Value = "Image too large for LCD (" & PrinterConsole.ImagePreview.InfoHeight & "," & PrinterConsole.ImagePreview.InfoWidth & ")."
                            .Fields("TimeOfError").Value = Now
                            .UpdateBatch adAffectCurrent
                        End If
                    End With
                    Exit Function
                End If
                GetImageData = True             'File is ready, and file information is loaded
                Exit Function                   'Get out of here
            End If
        End If
        ReadTry = ReadTry + 1
    Loop
    '---- There is a problem with the file - may not be an image!  Add record to FileErrors Table
    With DB.rsFileErrors
        If .RecordCount > 0 Then .MoveFirst
        .Find "ImageFileName='" & FileName & "'"
        If .EOF Then
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("PrintQue").Value = QuePath
            .Fields("ImageFileName").Value = FileName
            .Fields("ErrorDescription").Value = "Error#" & ReturnVal & " " & DB.GetLeadError(ReturnVal)
            .Fields("TimeOfError").Value = Now
            .UpdateBatch adAffectCurrent
        End If
    End With
    FileErrorsForm.FileErrorGrid.Refresh ssRefetchAndFireInitializeRow
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetQuantityFromFileName                               **
'**                                                                        **
'**  Description..:  This routine returns Print Quantity from File Name.   **
'**                                                                        **
'****************************************************************************
Public Function GetQuantityFromFileName(SourceName As String) As Integer
    '---- This routine retrieves the punch code from the file name (passed by Kodak DP2 - _ppppp_xxxxx_qty.jpg, where qty is quantity)
    On Error GoTo ErrorHandler
    Dim Qty As String
    Qty = GetDelimitedTextFromFileName(SourceName, 1)
    If Trim(Qty) = "" Then
        GetQuantityFromFileName = 0
    Else
        GetQuantityFromFileName = CInt(Qty)
    End If
    Exit Function
ErrorHandler:
    GetQuantityFromFileName = 1
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetPunchCodeFromFileName                              **
'**                                                                        **
'**  Description..:  This routine returns Punch Code from File Name.       **
'**                                                                        **
'****************************************************************************
Public Function GetPunchCodeFromFileName(SourceName As String) As Byte
    '---- This routine retrieves the punch code from the file name (passed by Kodak DP2 - _ppppp_xxxxx_qty.jpg, where pppp is the punch code)
    On Error GoTo ErrorHandler
    Dim PunchCode As String
    PunchCode = GetDelimitedTextFromFileName(SourceName, 2)
    If Trim(PunchCode) = "" Then
        GetPunchCodeFromFileName = 0
    Else
        If PunchCode < 256 Then
            GetPunchCodeFromFileName = CByte(PunchCode)
        Else
            GetPunchCodeFromFileName = 0
        End If
    End If
    Exit Function
ErrorHandler:
    MsgBox "Error " & Err.Number & " in GetPunchCodeFromFileName."
    GetPunchCodeFromFileName = 1
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetDelimitedTextFromFileName                          **
'**                                                                        **
'**  Description..:  This routine returns text from File Name.             **
'**                                                                        **
'****************************************************************************
Public Function GetDelimitedTextFromFileName(SourceText As String, PositionNo As Integer) As String
    On Error GoTo ErrorHandler
    Dim Msgs() As String, Tmp As String, ExtensionPos As Integer
    
    Msgs = Split(Trim(SourceText), "_")
    If UBound(Msgs) >= 1 Then
        Tmp = Msgs(UBound(Msgs) - (PositionNo - 1))
        ExtensionPos = InStr(1, Tmp, ".", vbTextCompare)
        If ExtensionPos > 0 Then
            GetDelimitedTextFromFileName = Left(Tmp, ExtensionPos - 1)
        Else
            GetDelimitedTextFromFileName = Tmp
        End If
    Else
        GetDelimitedTextFromFileName = ""
    End If
    Exit Function
ErrorHandler:
    MsgBox "Error " & Err.Number & " in GetDelimitedTextFromFileName."
    GetDelimitedTextFromFileName = ""
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetFujiTextFromFileName                               **
'**                                                                        **
'**  Description..:  This routine returns text from a Fuji Style File Name.**
'**                                                                        **
'****************************************************************************
Public Function GetFujiTextFromFileName(ByVal Delimiter As String, ByVal SourceText As String) As String
    On Error GoTo ErrorHandler
    Dim Msgs() As String, Tmp As String, pos As Integer
    SourceText = Left(SourceText, Len(SourceText) - 4)          'get rid of extension
    Msgs = Split(Trim(SourceText), "_")
    If UBound(Msgs) >= 1 Then
        '---- Find the Delimiter in the array
        For pos = 0 To UBound(Msgs())
            If Msgs(pos) = Delimiter Then
                GetFujiTextFromFileName = Msgs(pos + 1)
            End If
        Next
    Else
        GetFujiTextFromFileName = ""
    End If
    Exit Function
ErrorHandler:
    GetFujiTextFromFileName = ""
End Function

Public Sub DisableToolBarForPrinting()
    '---- Disable buttons to prevent bad user interaction during the cycling
    With Me.MainToolBar
        .Tools("ID_Add").Enabled = False
        .Tools("ID_Erase").Enabled = False
        .Tools("ID_PrintOn").Enabled = False
        .Tools("ID_PrintOff").Enabled = False
        .Tools("ID_HoldOn").Enabled = False
        .Tools("ID_HoldOff").Enabled = False
        .Tools("ID_AdvancePaper").Enabled = False
        .Tools("ID_RemovePaper").Enabled = False
        .Tools("ID_StartServer").Enabled = False
        .Tools("ID_Preview").Enabled = False
        .Tools("ID_Refresh").Enabled = False
    End With
End Sub

Public Sub EnableToolBarAfterPrinting()
    '---- Disable buttons to prevent bad user interaction during the cycling
    With Me.MainToolBar
        .Tools("ID_Add").Enabled = True
        .Tools("ID_Erase").Enabled = True
        .Tools("ID_PrintOn").Enabled = True
        .Tools("ID_PrintOff").Enabled = True
        .Tools("ID_HoldOn").Enabled = True
        .Tools("ID_HoldOff").Enabled = True
        .Tools("ID_AdvancePaper").Enabled = True
        .Tools("ID_RemovePaper").Enabled = True
        .Tools("ID_StartServer").Enabled = True
        .Tools("ID_Preview").Enabled = True
        .Tools("ID_Refresh").Enabled = True
        .Refresh
    End With
End Sub

