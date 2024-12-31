Attribute VB_Name = "DVP2QC_Main"
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2 Quality Control                                     **
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
Public PrinterName As String                        'The Primary Key into the Settings Database
Public AppPath As String
Public DemoMode As Boolean                          'Set to TRUE when printer is in DEMO mode (no dongle)
Public DebugMode As Boolean                         'Application debug flag
Public CurrentPrintFile As String
Public DB As New DataBaseInterface
Public StartupDrive As String                       'This is the disk drive the application started on
Public InitializationFileName As String
Public OffsetFilePath As String                     'This is the path to the offset files
Public LutFilePath As String                        'This is the path to the LUT files
Public DatabasePath As String
Public PrinterQuePath As String
Public PrintQueFile As String
Public SettingsFolder As String                     'This is the settings folder (same as apppath)


Public Sub Main()
    If App.PrevInstance = True Then                         'If the application is already running
        MsgBox "DVP2 QC is already running.", vbSystemModal + vbCritical + vbOKOnly, "Error"
        End
    End If
    ReadIniFile                                         'Read DVP2 Initialization File
    InitLogFile                                         'Initialize the log file
    
    DebugMode = True
    
    AppLog InfoMsg, "Main,Setting startup drive to " & StartupDrive
    AppLog InfoMsg, "Main,Setting INI file to " & InitializationFileName
    AppLog InfoMsg, "Main,Setting Database Path to " & DatabasePath
    AppLog InfoMsg, "Main,Setting Print Que Path to " & PrinterQuePath
    AppLog InfoMsg, "Main,Setting Printer Name to " & PrinterName
    AppLog InfoMsg, "Main,Setting Settings Path to " & SettingsFolder
    AppLog InfoMsg, "Main,Setting Offset File Path to " & OffsetFilePath
    AppLog InfoMsg, "Main,Setting LUT File Path to " & LutFilePath
    
    Load ErrorForm                                      'Load the application error handler form
    
    '---- Validate Paths
    With FileSystemHandle
        If .FolderExists(OffsetFilePath) = False Then
            .CreateFolder OffsetFilePath
        End If
        If .FolderExists(LutFilePath) = False Then
            .CreateFolder LutFilePath
        End If
    End With
    
    DB.OpenDatabase DatabasePath, ""                    'Connect to the Database
    Load ColorControlForm
    Load LutControlForm
    Load OffsetControlForm
    Load SettingsControlForm
    Load PrinterControlForm
    Load MainForm
    MainForm.Show
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
    AppPath = Trim(App.Path)                        'VB Application Path
    ChDir AppPath                                   'Make sure Application Path is current directory
    StartupDrive = IIf(Mid(AppPath, 2, 1) = ":", Left(AppPath, 2), "C:")
    InitializationFileName = AppPath & "\DVP2_QC.ini"
    If FileSystemHandle.FileExists(InitializationFileName) = False Then
        '--- The INI File is gone - create a standard one.
        MsgBox "Missing " & InitializationFileName & " creating a standard one."
        Set fh = FileSystemHandle.CreateTextFile(InitializationFileName, True, False)
        fh.WriteLine "[Main]"
        fh.WriteLine "DatabasePath=C:\DVP2_R2\Database\Settings.mdb"
        fh.WriteLine "PrintQuePath=C:\DVP2_R2\PrintQue.mdb"
        fh.WriteLine "OffsetFilePath=C:\DVP2_R2\DVP2 Printers\"
        fh.WriteLine "LutFilePath=C:\DVP2_R2\DVP2 Printers\"
        fh.WriteLine "SettingsPath=C:\DVP2_R2\DVP2 Printers\Default\"
        fh.Close
        Set fh = Nothing
    End If
    '---- Database Path
    TmpBuf = String(128, 0)
    CharsRead = GetPrivateProfileString("Main", "DatabasePath", "C:\DVP2_R2\Settings.mdb", TmpBuf, 80, InitializationFileName)
    If CharsRead <> 0 Then DatabasePath = Left(TmpBuf, CharsRead)
    '---- PrintQue Path
    TmpBuf = String(128, 0)
    CharsRead = GetPrivateProfileString("Main", "PrintQuePath", "C:\DVP2_R2\PrintQue.mdb", TmpBuf, 80, InitializationFileName)
    If CharsRead <> 0 Then PrinterQuePath = Left(TmpBuf, CharsRead)
    '---- Settings Path
    TmpBuf = String(128, 0)
    CharsRead = GetPrivateProfileString("Main", "SettingsPath", "C:\DVP2_R2\", TmpBuf, 80, InitializationFileName)
    If CharsRead <> 0 Then SettingsFolder = Left(TmpBuf, CharsRead)
    '---- Offset File Path
    TmpBuf = String(128, 0)
    CharsRead = GetPrivateProfileString("Main", "OffsetFilePath", "C:\DVP2_R2\CalFiles\", TmpBuf, 80, InitializationFileName)
    If CharsRead <> 0 Then OffsetFilePath = Left(TmpBuf, CharsRead)
    '---- LUT File Path
    TmpBuf = String(128, 0)
    CharsRead = GetPrivateProfileString("Main", "LutFilePath", "C:\DVP2_R2\CalFiles\", TmpBuf, 80, InitializationFileName)
    If CharsRead <> 0 Then LutFilePath = Left(TmpBuf, CharsRead)
    TmpBuf = ""
    
    Exit Sub
ErrorHandler:
    MsgBox "Error reading initialization file.", vbApplicationModal + vbCritical + vbOKOnly, "Critical Error"
    End
End Sub
