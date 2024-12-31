Attribute VB_Name = "DVP2_Main"
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2                                                     **
'**                                                                        **
'** Module.....: Main                                                      **
'**                                                                        **
'** Description: This module is called on startup.                         **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.71 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit                                     'Require explicit variable declaration
Global Const DefaultInterval As Integer = 5000      'Default Queue Interval
Global Const PrinterIdleMessage As String = "Printer Idle."
Global Const PrinterNotReadyMessage As String = "Printer not ready."
Global Const StartPrintingMessage As String = "Start Printing"
Global Const StopPrintingMessage As String = "Stop Printing"
Global Const StartServerMessage As String = "Start Server"
Global Const StopServerMessage As String = "Stop Server"
Public PrinterName As String                        'The Primary Key into the Settings Database
Public AppPath As String                            'The Path the application started in
Public InitializationFileName As String             'The name of the initialization file (.ini file)
Public DemoMode As Boolean                          'Set to TRUE when printer is in DEMO mode (no dongle)
Public CurrentPrintFile As String                   'The Name of the Current Print File (Image File on LCD)
Public StartupDrive As String                       'This is the disk drive the application started on
Public DatabasePath As String                       'The path to the application database
Public PrinterQuePath As String                     'The path to the print queue database
Public FastSettingsFolder As String                 'The path to the Fast Settings Folder (RAMDISK)
Public SettingsFolder As String                     'This is the settings folder (same as apppath)
Public RamDiskConnected As Boolean                  'Set TRUE when RAM Disk is accessible
Public OffsetFilePath As String                     'This is the path to the offset files
Public LutFilePath As String                        'This is the path to the LUT files
Public MaxLutBlock As Integer                       'The maximum # of LUT Blocks to read (germans=48,picto=72)
Public DB As New DataBaseInterface                  'Interface to DVP2 Database
Public ImageTimeOut As Long                         'The amount of time the current image is displayed (for clearing LCD)
Public WatchForImageTimeOut As Boolean              'Flag indicating non-clear image is on LCD

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Main                                                  **
'**                                                                        **
'**  Description..:  This is the main application function - called by Win.**
'**                                                                        **
'****************************************************************************
Public Sub Main()
    On Error GoTo ErrorHandler
    
    If App.PrevInstance = True Then                     'If the application is already running
        MsgBox "DVP2 is already running.", vbSystemModal + vbCritical + vbOKOnly, "Error"
        End                                             'Prevent the program from running again
    End If
    AppPath = Trim(App.Path)                            'VB Application Path
    
    Load frmSplash
    frmSplash.Show vbModeless
    
    frmSplash.SetStatus "Initializing log file..."
    InitLogFile                                         'Initialize the log file
    
    frmSplash.SetStatus "Reading initialization file..."
    ReadIniFile                                         'Read DVP2 Initialization File
    
    frmSplash.SetStatus "Loading error handler..."
    Load ErrorForm                                      'Load the application error handler form
    
    frmSplash.SetStatus "Checking for USB Key..."
    Load UsbKeyDiagnostics
    DemoMode = IIf(UsbKeyDiagnostics.GetSecurity(&H1F5), False, True) 'Check for dongle
    
    frmSplash.SetStatus "Loading passwords..."
    Load PasswordForm                                   'Load the password form
    
    frmSplash.SetStatus "Opening database..."
    DB.OpenDatabase DatabasePath, PrinterQuePath        'Connect to the Database
    DB.GetPrinterSettings                               'Retrieve all global settings from the database
    DB.GetPrinterStatistics                             'Retrieve printer statistics
    
    frmSplash.SetStatus "Configuring Settings Folders..."
    CopySettingsFolder                                  'Copy LUT & Offset files to RAMDISK or Tmp Directory
    
    frmSplash.SetStatus "Loading Hot Folders..."
    Load HotFolderSettingsForm
    HotFolderSettingsForm.Setup
    
    frmSplash.SetStatus "Loading Print Queue History..."
    Load PrintQueHistoryForm
    PrintQueHistoryForm.Setup
    
    frmSplash.SetStatus "Loading Q.C. Settings.."
    Load PrinterQcForm
    
    frmSplash.SetStatus "Loading File Errors.."
    Load FileErrorsForm
    FileErrorsForm.Setup
    
    frmSplash.SetStatus "Loading Statistics.."
    Load PrinterStatisticsForm
    PrinterStatisticsForm.Setup
    
    frmSplash.SetStatus "Loading Settings.."
    Load SettingsControlForm
    SettingsControlForm.Setup
    
    frmSplash.SetStatus "Loading color control..."
    Load ColorControlForm
    ColorControlForm.Setup                                  'Setup the Q.C. Color Control
    
    frmSplash.SetStatus "Loading Emulsion control..."
    Load EmulsionForm
    EmulsionForm.Setup                                      'Setup the Emulsion Control
    
    frmSplash.SetStatus "Loading LUT control..."
    Load LutControlForm
    LutControlForm.Setup                                    'Setup the Q.C. LUT Control
    
    frmSplash.SetStatus "Loading Offset control..."
    Load OffsetControlForm
    OffsetControlForm.Setup                                 'Setup the Q.C. Offset Control
    
    frmSplash.SetStatus "Communicating with panel..."
    Load DiagnosticsForm                                    'Load the diagnostics form - this is the main application form
    DiagnosticsForm.InitializeHardware                      'Initialize Hardware (Digital I/O, Communications Ports, LCD)
    
    frmSplash.SetStatus "Loading Punch Settings..."
    Load PunchDiagnostics
    PunchDiagnostics.Setup
    
    
    frmSplash.SetStatus "Communicating with motors..."
    Load MotorDiagnostics
    MotorDiagnostics.Setup
    
    frmSplash.SetStatus "Communicating with backwriters..."
    Load BackWriterDiagnostics
    BackWriterDiagnostics.Setup
    
    frmSplash.SetStatus "Loading size settings..."
    Load SizeSettingsForm
    SizeSettingsForm.Setup                                  'Sstup the Printer Size Settings Control
    SizeSettingsForm.ClearImage
    
    frmSplash.SetStatus "Loading printer console..."
    Load PrinterConsole
    PrinterConsole.Setup                                    'Setup the Print Queue Control
    
    ImageTimeOut = 0
    
    frmSplash.SetStatus "Configuring printer..."
    Load MainForm
    frmSplash.Hide
    MainForm.Show                                           'Show the main form
    
    Exit Sub
ErrorHandler:
    MsgBox "Error starting DVP2.  Please contact technical support.", vbSystemModal + vbCritical + vbOKOnly, "Error"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ReadIniFile                                           **
'**                                                                        **
'**  Description..:  This routine Creates/Reads the DVP2.INI File.         **
'**                                                                        **
'****************************************************************************
Private Sub ReadIniFile()
    On Error GoTo ErrorHandler
    Dim TmpBuf As String, CharsRead As Integer, fh As TextStream
    AppLog InfoMsg, "ReadIniFile,Starting application in " & AppPath
    ChDir AppPath                                           'Make sure Application Path is current directory
    StartupDrive = IIf(Mid(AppPath, 2, 1) = ":", Left(AppPath, 2), "C:")
    InitializationFileName = AppPath & "\DVP2.ini"
    If FileSystemHandle.FileExists(InitializationFileName) = False Then
        '--- The INI File is gone - create a standard one.
        MsgBox "Missing " & InitializationFileName & " creating a standard one."
        Set fh = FileSystemHandle.CreateTextFile(InitializationFileName, True, False)
        fh.WriteLine "[Main]"
        fh.WriteLine "DatabasePath=C:\DVP2_R2\Settings.mdb"
        fh.WriteLine "PrintQuePath=C:\DVP2_R2\PrintQue.mdb"
        fh.WriteLine "SettingsPath=C:\DVP2_R2\"
        fh.WriteLine "PrinterName=DVP2_0001"
        fh.Close
        Set fh = Nothing
    End If
    '---- Database Path
    TmpBuf = String(181, 0)
    CharsRead = GetPrivateProfileString("Main", "DatabasePath", "C:\DVP2_R3\Settings.mdb", TmpBuf, 180, InitializationFileName)
    If CharsRead <> 0 Then DatabasePath = Left(TmpBuf, CharsRead)
    '---- PrintQue Path
    TmpBuf = String(181, 0)
    CharsRead = GetPrivateProfileString("Main", "PrintQuePath", "C:\DVP2_R3\PrintQue.mdb", TmpBuf, 180, InitializationFileName)
    If CharsRead <> 0 Then PrinterQuePath = Left(TmpBuf, CharsRead)
    '---- Printer Name
    TmpBuf = String(181, 0)
    CharsRead = GetPrivateProfileString("Main", "PrinterName", "DVP2_0001", TmpBuf, 180, InitializationFileName)
    If CharsRead <> 0 Then PrinterName = Left(TmpBuf, CharsRead)
    '---- Settings Path
    TmpBuf = String(181, 0)
    CharsRead = GetPrivateProfileString("Main", "SettingsPath", "C:\DVP2_R3\", TmpBuf, 180, InitializationFileName)
    If CharsRead <> 0 Then SettingsFolder = Left(TmpBuf, CharsRead)
    TmpBuf = ""
    LutFilePath = SettingsFolder
    OffsetFilePath = SettingsFolder
    SettingsFolder = SettingsFolder & PrinterName & "\"
    AppLog InfoMsg, "ReadIniFile,Setting startup drive to " & StartupDrive
    AppLog InfoMsg, "ReadIniFile,Setting INI file to " & InitializationFileName
    AppLog InfoMsg, "ReadIniFile,Setting Database Path to " & DatabasePath
    AppLog InfoMsg, "ReadIniFile,Setting Print Que Path to " & PrinterQuePath
    AppLog InfoMsg, "ReadIniFile,Setting Printer Name to " & PrinterName
    AppLog InfoMsg, "ReadIniFile,Setting Settings Path to " & SettingsFolder
    Exit Sub
ErrorHandler:
    MsgBox "Error reading initialization file.", vbApplicationModal + vbCritical + vbOKOnly, "Critical Error"
    End
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CopySettingsFolder                                    **
'**                                                                        **
'**  Description..:  This routine copies LUT & Offset Files to Temp Folder.**
'**  If the printer controller has a RAM DISK drive installed, then this   **
'**  procedure will copy the necessary LUT and OFFSET files to it to speed **
'**  up image processing roughly 50-75%.                                   **
'**                                                                        **
'****************************************************************************
Private Function CopySettingsFolder() As Boolean
    On Error GoTo ErrorHandler
    
    Dim mPC As PerformanceCounter                   'Declare a performance counter
    Set mPC = New PerformanceCounter                'Initialize the performance counter
    mPC.StartTimer True                             'Start a high precision timer
    
    Dim drv As Drive, Tmp As Boolean                'Declare  FileSystem and temp variables
    
    With FileSystemHandle                           'Using the global file system handle object
        
        '---- If the RAMDISK exists, then check to see if ready and copy files
        If .DriveExists(DB.RamDiskPath) Then         'Check to see if the drive exists
            Set drv = .GetDrive(DB.RamDiskPath)      'Get Disk Drive Information (must exist for this call to work)
            If drv.IsReady = True Then
                RamDiskConnected = True
                FastSettingsFolder = DB.RamDiskPath & "\Settings\"
                CurrentPrintFile = DB.RamDiskPath & "\PRINTFILE.BMP"
                AppLog InfoMsg, "CopySettingsFolder,Using RAM disk for temporary files,Folder=" & FastSettingsFolder
                '---- Initialize RAM-DISK and Copy LUT/Offset Files
                If .FolderExists(FastSettingsFolder) = False Then
                    AppLog InfoMsg, "CopySettingsFolder,Creating settings folder " & FastSettingsFolder
                    .CreateFolder FastSettingsFolder
                    .CreateFolder FastSettingsFolder & "LUT\"
                    .CreateFolder FastSettingsFolder & "Offset\"
                Else
                    AppLog InfoMsg, "CopySettingsFolder,Using settings folder " & FastSettingsFolder
                End If
                AppLog InfoMsg, "CopySettingsFolder,Copying LUT Files from " & SettingsFolder & " to " & FastSettingsFolder
                .CopyFile SettingsFolder & "LUT\*.lut", FastSettingsFolder & "LUT\"
                AppLog InfoMsg, "CopySettingsFolder,Copying Offset Files from " & SettingsFolder & " to " & FastSettingsFolder
                .CopyFile SettingsFolder & "Offset\*.frm", FastSettingsFolder & "Offset\"
            Else
                RamDiskConnected = False
                FastSettingsFolder = SettingsFolder
                CurrentPrintFile = SettingsFolder & "PRINTFILE.BMP"
                AppLog InfoMsg, "CopySettingsFolder,Using Hard Disk for temporary files, RAMDISK is not formatted,Folder=" & FastSettingsFolder
            End If
        Else
            RamDiskConnected = False
            FastSettingsFolder = SettingsFolder
            CurrentPrintFile = SettingsFolder & "PRINTFILE.BMP"
            AppLog InfoMsg, "CopySettingsFolder,Using Hard Disk for temporary files, RAMDISK is not available,Folder=" & FastSettingsFolder
        End If
    End With
    AppLog InfoMsg, "CopySettingsFolder,Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    Set mPC = Nothing
    Exit Function
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
    ErrorForm.ReportError "DVP2_Main:CopySettingsFolder", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

