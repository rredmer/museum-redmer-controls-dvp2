VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBaseInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2 Quality Control                                     **
'**                                                                        **
'** Module.....: DatabaseInterface                                         **
'**                                                                        **
'** Description: Provides database connectivity                            **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.71 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Const CON_SQLCONNECT = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=moose1027;Initial Catalog=DVP2_R2;Data Source=(local)"
Const CON_JETCONNECT = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Private cnn As ADODB.Connection                         'ADO Database Connection
Private cat As ADOX.Catalog

Public rsPrinterList As ADODB.Recordset
Public rsOptions As ADODB.Recordset
Public rsSettings As ADODB.Recordset
Public rsExposureTime As ADODB.Recordset
Public rsExposureTimeHistory As ADODB.Recordset
Public rsPrintSizes As ADODB.Recordset
Public rsHotFolders As ADODB.Recordset                  'Hot Folder Recordset
Public rsInputs As ADODB.Recordset
Public rsOutputs As ADODB.Recordset
Public rsPrinterStatistics As ADODB.Recordset           'Printer Statistics Recordset
Public rsLutDensiValues As ADODB.Recordset
Public rsOffsets As ADODB.Recordset
Public rsLuts As ADODB.Recordset
Public rsLabAim As ADODB.Recordset
Public rsDailyCalibration As ADODB.Recordset
Public rsRingArounds As ADODB.Recordset
Public rsRingAroundsAutoScan As ADODB.Recordset
Public rsBackWriterSettings As ADODB.Recordset
Public rsEmulsions As ADODB.Recordset
Public rsEmulsionData As ADODB.Recordset
Private rsLeadErrorCodes As ADODB.Recordset             'Lead Error Descriptions
Private QueCnn As ADODB.Connection                      'Connection to print que database
Public rsPrintQue As ADODB.Recordset                    'Print Que Recordset
Public rsPrintHistory As ADODB.Recordset                'Print History Recordset
Public rsFileErrors As ADODB.Recordset                  'File Errors

'---- Printer Settings
Private Const RecCalibrationImagePath As Integer = 1
Private Const RecLogFilePath As Integer = 2
Private Const RecRamDiskPath As Integer = 3
Private Const RecAdminPassword As Integer = 4
Private Const RecMotorPPISetting As Integer = 5
Private Const RecPaperAdvanceLength As Integer = 6
Private Const RecPaperAdvanceTearOffLength As Integer = 7
Private Const RecQualityControlPassword As Integer = 8
Private Const RecBW_Text1 As Integer = 10
Private Const RecBW_Text2 As Integer = 11
Private Const RecBW_Char_Delay As Integer = 12
Private Const RecBW_Dot_OnTime As Integer = 13
Private Const RecBW_Dot_OffTime As Integer = 14
Private Const RecBW_MaxChars As Integer = 15
Private Const RecBW_AdvDelay As Integer = 16
Private Const RecBW_RibbonSpeed As Integer = 17
Private Const RecBW_FontSize As Integer = 18
Private Const RecBW_MotorWait As Integer = 19
Private Const RecMaskHomePrintCount As Integer = 30
Private Const RecMaskStepsOnClick As Integer = 31
Private Const RecMaskLeftLimit As Integer = 32
Private Const RecMaskRightLimit As Integer = 33
Private Const RecMaskHoldingTorque As Integer = 34
Private Const RecPunchSolenoidTime As Integer = 40
Private Const RecPunchEngageTime As Integer = 41
Private Const RecPunchDisengageTime As Integer = 42
Private Const RecPunchPkgExtendTime As Integer = 43
Private Const RecDigitalCutSize As Integer = 50
Private Const RecDigitalCut8x10Offset As Integer = 51
Private Const RecDigitalCut11x14Offset As Integer = 52
Private Const RecPurgeLogDays As Integer = 60
Private Const RecPurgeExposureHistoryDays As Integer = 61
Private Const RecPurgeQueHistoryDays As Integer = 62

Public CalibrationImagePath As String
Public LogFilePath As String
Public RamDiskPath As String
Public AdminPassword As String
Public MotorPPISetting As Integer
Public PaperAdvanceLength As Integer
Public PaperAdvanceTearOffLength As Integer
Public QualityControlPassword As String
Public BW_Text1 As String
Public BW_Text2 As String
Public BW_Char_Delay As Integer
Public BW_Dot_OnTime As Integer
Public BW_Dot_OffTime As Integer
Public BW_MaxChars As Integer
Public BW_AdvDelay As Integer
Public BW_RibbonSpeed As Integer
Public BW_FontSize As Integer
Public BW_MotorWait As String
Public MaskHomePrintCount As Integer
Public MaskStepsOnClick As Integer
Public MaskLeftLimit As Long
Public MaskRightLimit As Long
Public MaskHoldingTorque As Integer
Public PunchSolenoidTime As Integer
Public PunchEngageTime As Integer
Public PunchDisengageTime As Integer
Public PunchPkgExtendTime As Integer
Public DigitalCutSize As Integer
Public DigitalCut8x10Offset As Integer
Public DigitalCut11x14Offset As Integer
Public PurgeLogDays As Integer
Public PurgeExposureHistoryDays As Integer
Public PurgeQueHistoryDays As Integer

'----- Printer Options
Private Const RecDebugMode As Integer = 1                                        'Debug flag, Checkbox control set to 1 for debug-level message logging
Private Const RecNordConversion As Integer = 2
Private Const RecDoorSwitchInstalled As Integer = 3
Private Const RecBackWritersInstalled As Integer = 4
Private Const RecDualBackWritersInstalled As Integer = 5
Private Const RecPunchEnable As Integer = 6
Private Const RecPackagePunchEnable As Integer = 7
Private Const RecRS422MotorConfiguration As Integer = 8
Private Const RecPlatenCylinderEnable As Integer = 9
Private Const RecQueGetQty As Integer = 10
Private Const RecQueGetPunch As Integer = 11
Private Const RecQueGetText As Integer = 12
Private Const RecQueSort As Integer = 13
Private Const RecFujiFileNameEncoding As Integer = 14
Private Const RecBackWriterOverride As Integer = 15
Private Const RecApplyGrayBorder As Integer = 20
Private Const RecApplyMullerSohnLUT As Integer = 21
Private Const RecApplyPictoLUT As Integer = 22
Private Const RecRenderCutCode As Integer = 30
Private Const RecEnableQC As Integer = 40
Private Const RecQueHistory As Integer = 41

Public DebugMode As Boolean
Public StepperMaskInstalled As Boolean
Public DoorSwitchInstalled  As Boolean
Public BackWritersInstalled  As Boolean
Public DualBackWritersInstalled  As Boolean
Public PunchEnable As Boolean
Public PackagePunchEnable As Boolean
Public RS422MotorConfiguration  As Boolean
Public PlatenCylinderEnable As Boolean
Public QueGetQty As Boolean
Public QueGetPunch As Boolean
Public QueGetText As Boolean
Public QueSort As Boolean
Public ApplyGrayBorder As Boolean
Public ApplyMullerSohnLUT As Boolean
Public ApplyPictoLUT As Boolean
Public RenderCutCode As Boolean
Public EnableQC As Boolean
Public QueHistory As Boolean
Public FujiFileNameEncoding As Boolean
Public BackWriterOverride As Boolean

'---- Printer Statistics
Private Const RecStatTotalImages As Integer = 1
Private Const RecStatTotalExposures As Integer = 2
Private Const RecStatExposuresPerSecond As Integer = 3
Private Const RecStatAverageExposureTime As Integer = 4
Private Const RecStatAverageServerTime As Integer = 5
Private Const RecStatAverageAdvanceTime As Integer = 6
Private Const RecStatTimeRunning As Integer = 7
Private Const RecStatPaperUsed As Integer = 8

Public StatTotalImages As Double
Public StatTotalExposures As Double
Public StatExposuresPerSecond As Double
Public StatAverageExposureTime As Double
Public StatAverageServerTime As Double
Public StatAverageAdvanceTime As Double
Public StatTimeRunning As Double
Public StatPaperUsed As Double

'****************************************************************************
'**                                                                        **
'**  Procedure....:  OpenDatabase                                          **
'**                                                                        **
'**  Description..:  This routine connects to the application database.    **
'**                                                                        **
'****************************************************************************
Public Function OpenDatabase(DatabasePath As String, PrinterQuePath As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim ConnectString As String, SQLcmd As String, mPC As New PerformanceCounter
    mPC.StartTimer True
    
    If PrinterQuePath <> "" Then                                'Printer Que Path only used on Printer, not Q.C.
        Set QueCnn = New ADODB.Connection
        Set rsPrintQue = New ADODB.Recordset                    'Print Que Recordset
        Set rsPrintHistory = New ADODB.Recordset                'Print History Recordset
        Set rsFileErrors = New ADODB.Recordset                  'File Errors
    
        '---- Open the Print Que database
        If FileSystemHandle.FileExists(PrinterQuePath) = False Then
            AppLog ErrorMsg, "OpenDatabase,Path [" & PrinterQuePath & "] to printer queue database not found."
            MsgBox "Path [" & PrinterQuePath & "] to printer queue database not found."
            End
        End If
        CompactDatabase PrinterQuePath
        ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PrinterQuePath & ";Persist Security Info=False"
        AppLog InfoMsg, "OpenDatabase,Opening ADO Connection=" & ConnectString
        QueCnn.CursorLocation = adUseClient
        QueCnn.Open ConnectString
        
        '---- Active Print Que recordset
        Set rsPrintQue = New ADODB.Recordset
        SQLcmd = "SELECT * FROM ActivePrintQue WHERE PrinterName='" & PrinterName & "';"
        AppLog InfoMsg, "OpenDatabase,Opening Printer Queue Recordset as " & SQLcmd
        rsPrintQue.Open SQLcmd, QueCnn, adOpenDynamic, adLockOptimistic, adCmdText
        
        Set rsFileErrors = New ADODB.Recordset
        SQLcmd = "SELECT * FROM FileErrors WHERE PrinterName='" & PrinterName & "';"
        AppLog InfoMsg, "OpenDatabase,Opening File Error Recordset as " & SQLcmd
        rsFileErrors.Open SQLcmd, QueCnn, adOpenDynamic, adLockOptimistic, adCmdText
        
        Set rsPrintHistory = New ADODB.Recordset
        SQLcmd = "SELECT * FROM PrintHistory WHERE PrinterName='" & PrinterName & "';"
        AppLog InfoMsg, "OpenDatabase,Opening Print History Recordset as " & SQLcmd
        rsPrintHistory.Open SQLcmd, QueCnn, adOpenDynamic, adLockOptimistic, adCmdText
    End If
    
    '---- Validate the Database Path - we can't run without the database!
    If UCase(Left(DatabasePath, 3)) = "SQL" Then
        '--- Need to validate SQL Connection...
        ConnectString = CON_SQLCONNECT
        AppLog InfoMsg, "OpenDatabase,Connecting to SQL Server using: " & ConnectString
    Else
        If FileSystemHandle.FileExists(DatabasePath) = False Then
            AppLog ErrorMsg, "OpenDatabase,Path [" & DatabasePath & "] to database not found."
            MsgBox "Path [" & DatabasePath & "] to database not found."
            End
        End If
        ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & ";Persist Security Info=true"
        AppLog InfoMsg, "OpenDatabase,Connecting to Micrsoft Access Database using: " & ConnectString
    End If
    Set rsPrinterList = New ADODB.Recordset
    Set rsOptions = New ADODB.Recordset
    Set rsSettings = New ADODB.Recordset
    Set rsPrintSizes = New ADODB.Recordset
    Set rsExposureTime = New ADODB.Recordset
    Set rsExposureTimeHistory = New ADODB.Recordset
    Set rsHotFolders = New ADODB.Recordset
    Set rsOutputs = New ADODB.Recordset
    Set rsInputs = New ADODB.Recordset
    Set rsPrinterStatistics = New ADODB.Recordset
    Set rsLutDensiValues = New ADODB.Recordset
    Set rsLabAim = New ADODB.Recordset
    Set rsDailyCalibration = New ADODB.Recordset
    Set rsRingArounds = New ADODB.Recordset
    Set rsRingAroundsAutoScan = New ADODB.Recordset
    Set rsOffsets = New ADODB.Recordset
    Set rsLuts = New ADODB.Recordset
    Set rsBackWriterSettings = New ADODB.Recordset
    Set rsEmulsions = New ADODB.Recordset
    Set rsEmulsionData = New ADODB.Recordset
    
    AppLog InfoMsg, "OpenDatabase,Opening ADO Connection=" & ConnectString
    Set cnn = New ADODB.Connection
    cnn.CursorLocation = adUseClient
    cnn.Open ConnectString
    
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = cnn
    
    ValidateTables
    
    SQLcmd = "SELECT * FROM PrinterList ORDER BY PrinterName"
    AppLog InfoMsg, "OpenDatabase,Opening Printer List Recordset as " & SQLcmd
    rsPrinterList.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    If rsPrinterList.RecordCount > 0 Then
        If PrinterName = "" Then            'Printer Name is blank for Q.C., defined in Printer App by INI File
            rsPrinterList.MoveFirst
            PrinterName = rsPrinterList.Fields("PrinterName").Value
        End If
        GetPrinterRecordsets PrinterName
    Else
        MsgBox "No printers defined.", vbApplicationModal + vbCritical + vbOKOnly, "ERROR"
        End
    End If
    
    '---- LeadTools Error Description Recordset
    Set rsLeadErrorCodes = New ADODB.Recordset
    SQLcmd = "SELECT * FROM LeadErrorCodes"
    AppLog InfoMsg, "OpenDatabase,Opening Lead Tools Error Recordset as " & SQLcmd
    rsLeadErrorCodes.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    
    '---- Lab AIM
    Set rsLabAim = New ADODB.Recordset
    SQLcmd = "SELECT * FROM LabAim"
    AppLog InfoMsg, "OpenDatabase,Opening Lab Aim Recordset as " & SQLcmd
    rsLabAim.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    
    AppLog InfoMsg, "OpenDatabase,Timed," & Format(mPC.StopTimer, "####.####") & " seconds."
    Set FileSystemHandle = Nothing
    Set mPC = Nothing
    Exit Function
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
    AppLog ErrorMsg, "OpenDatabase, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Class_Terminate                                       **
'**                                                                        **
'**  Description..:  This routine closes recordsets & the databases.       **
'**                                                                        **
'****************************************************************************
Public Sub CloseDataBase()
    On Error GoTo ErrorHandler
    
    rsOptions.Close
    rsSettings.Close
    rsPrintSizes.Close
    rsExposureTime.Close
    rsInputs.Close
    rsOutputs.Close
    rsHotFolders.Close
    rsLeadErrorCodes.Close
    rsEmulsions.Close
    rsEmulsionData.Close
    cnn.Close
    Set cat = Nothing
    Set rsHotFolders = Nothing
    Set rsInputs = Nothing
    Set rsOutputs = Nothing
    Set rsOptions = Nothing
    Set rsSettings = Nothing
    Set rsPrintSizes = Nothing
    Set rsExposureTime = Nothing
    Set rsLeadErrorCodes = Nothing
    Set rsEmulsions = Nothing
    Set rsEmulsionData = Nothing
    Set cnn = Nothing
    
    If QueCnn Is Nothing Then
    Else
        If QueCnn.State = adStateOpen Then
            rsFileErrors.Close
            rsPrintQue.Close
            QueCnn.Close
            Set rsPrintQue = Nothing
            Set QueCnn = Nothing
        End If
    End If
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "Class_Terminate, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
    Resume Next
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetPrinterRecordsets                                  **
'**                                                                        **
'**  Description..:  This routine opens all recordsets used on the printer.**
'**                                                                        **
'****************************************************************************
Public Function GetPrinterRecordsets(PrinterName As String)
    On Error GoTo ErrorHandler
    
    With DB.rsPrinterList
        If .RecordCount > 0 Then .MoveFirst
        .Find "PrinterName='" & PrinterName & "'"
        If .EOF Then
            MsgBox "Internal Error.  Printer record not found."
            End
        End If
    End With
    
    GetOptionsRecordset
    GetSettingsRecordset
    GetPrintSizesRecordset
    GetInputRecordset
    GetOutputRecordset
    GetExposureTimeRecordset
    GetExposureTimeHistoryRecordset
    GetHotFolderRecordset
    GetPrinterStatisticsRecordset
    GetDailyCalibrationRecordset
    GetRingAroundRecordset
    GetRingAroundAutoScanRecordset
    GetOffsetRecordset
    GetLutRecordset
    GetBackWriterRecordset
    
    GetEmulsionsRecordset
    GetEmulsionDataRecordset
    
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetPrinterRecordsets, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetOptionsRecordset                                   **
'**                                                                        **
'**  Description..:  This routine returns the Option Recordset.            **
'**                                                                        **
'****************************************************************************
Private Sub GetOptionsRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsOptions
        SQLcmd = "SELECT * FROM PrinterOptions WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,OptionNumber"
        AppLog InfoMsg, "OpenDatabase,Opening Printer Option Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    End With
    ValidateOptions
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetOptionsRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetSettingsRecordset                                  **
'**                                                                        **
'**  Description..:  This routine returns the Settings Recordset.          **
'**                                                                        **
'****************************************************************************
Private Sub GetSettingsRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsSettings
        SQLcmd = "SELECT * FROM PrinterSettings WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,SettingNumber"
        AppLog InfoMsg, "OpenDatabase,Opening Printer Settings Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    End With
    ValidateSettings
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetSettingsRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetPrintSizesRecordset                                **
'**                                                                        **
'**  Description..:  This routine returns the Print Sizes Recordset.       **
'**                                                                        **
'****************************************************************************
Private Sub GetPrintSizesRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsPrintSizes
        SQLcmd = "SELECT * FROM PrintSizes WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,PrintSize;"
        AppLog InfoMsg, "OpenDatabase,Opening Print Size Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            '--- Copy Default Options
            Dim rsDefaults As ADODB.Recordset
            Set rsDefaults = New ADODB.Recordset
            SQLcmd = "SELECT * FROM PrintSizes WHERE PrinterName='Default' ORDER BY PrinterName,PrintSize;"
            AppLog InfoMsg, "OpenDatabase,Opening Default Print Size Setting Recordset as " & SQLcmd
            rsDefaults.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
            If rsDefaults.RecordCount = 0 Then
                AppLog ErrorMsg, "Default Print Sizes not found."
            Else
                rsDefaults.MoveFirst
                Do While Not rsDefaults.EOF
                    .AddNew
                    .Fields("PrinterName").Value = PrinterName
                    .Fields("PrintSize").Value = rsDefaults.Fields("PrintSize").Value
                    .Fields("Enable").Value = rsDefaults.Fields("Enable").Value
                    .Fields("AspectRatio").Value = rsDefaults.Fields("AspectRatio").Value
                    .Fields("Caption").Value = rsDefaults.Fields("Caption").Value
                    .Fields("RotateTable").Value = rsDefaults.Fields("RotateTable").Value
                    .Fields("Lens11x14").Value = rsDefaults.Fields("Lens11x14").Value
                    .Fields("VideoPosition").Value = rsDefaults.Fields("VideoPosition").Value
                    .Fields("LeftMaskPosition").Value = rsDefaults.Fields("LeftMaskPosition").Value
                    .Fields("RightMaskPosition").Value = rsDefaults.Fields("RightMaskPosition").Value
                    .Fields("LeftFlapSmall").Value = rsDefaults.Fields("LeftFlapSmall").Value
                    .Fields("LeftFlapLarge").Value = rsDefaults.Fields("LeftFlapLarge").Value
                    .Fields("RightFlap").Value = rsDefaults.Fields("RightFlap").Value
                    .Fields("PaperPreFeed").Value = rsDefaults.Fields("PaperPreFeed").Value
                    .Fields("PaperFeed").Value = rsDefaults.Fields("PaperFeed").Value
                    .Fields("PaperPunchFeed").Value = rsDefaults.Fields("PaperPunchFeed").Value
                    .UpdateBatch adAffectCurrent
                    rsDefaults.MoveNext
                Loop
            End If
            rsDefaults.Close
            Set rsDefaults = Nothing
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetPrintSizesRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetInputRecordset                                     **
'**                                                                        **
'**  Description..:  This routine returns the Digital Input Recordset.     **
'**                                                                        **
'****************************************************************************
Public Sub GetInputRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsInputs
        SQLcmd = "SELECT * FROM DigitalInputs WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,InputNumber"
        AppLog InfoMsg, "OpenDatabase,Opening Input Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            '--- Copy Default Options
            Dim rsDefaults As ADODB.Recordset
            Set rsDefaults = New ADODB.Recordset
            SQLcmd = "SELECT * FROM DigitalInputs WHERE PrinterName='Default' ORDER BY PrinterName,InputNumber"
            AppLog InfoMsg, "OpenDatabase,Opening Default Printer Inputs Recordset as " & SQLcmd
            rsDefaults.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
            If rsDefaults.RecordCount = 0 Then
                AppLog ErrorMsg, "Default Printer Inputs not found."
            Else
                rsDefaults.MoveFirst
                Do While Not rsDefaults.EOF
                    .AddNew
                    .Fields("PrinterName").Value = PrinterName
                    .Fields("InputNumber").Value = rsDefaults.Fields("InputNumber").Value
                    .Fields("Description").Value = rsDefaults.Fields("Description").Value
                    .Fields("InputEnabled").Value = rsDefaults.Fields("InputEnabled").Value
                    .Fields("TimeOut").Value = rsDefaults.Fields("TimeOut").Value
                    .UpdateBatch adAffectCurrent
                    rsDefaults.MoveNext
                Loop
            End If
            rsDefaults.Close
            Set rsDefaults = Nothing
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetInputRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetOutputRecordset                                    **
'**                                                                        **
'**  Description..:  This routine returns the Digital Output Recordset.    **
'**                                                                        **
'****************************************************************************
Public Sub GetOutputRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsOutputs
        SQLcmd = "SELECT * FROM DigitalOutputs WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,OutputNumber"
        AppLog InfoMsg, "OpenDatabase,Opening Output Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            '--- Copy Default Options
            Dim rsDefaults As ADODB.Recordset
            Set rsDefaults = New ADODB.Recordset
            SQLcmd = "SELECT * FROM DigitalOutputs WHERE PrinterName='Default' ORDER BY PrinterName,OutputNumber"
            AppLog InfoMsg, "OpenDatabase,Opening Default Printer Outputs Recordset as " & SQLcmd
            rsDefaults.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
            If rsDefaults.RecordCount = 0 Then
                AppLog ErrorMsg, "Default Printer Outputs not found."
            Else
                rsDefaults.MoveFirst
                Do While Not rsDefaults.EOF
                    .AddNew
                    .Fields("PrinterName").Value = PrinterName
                    .Fields("OutputNumber").Value = rsDefaults.Fields("OutputNumber").Value
                    .Fields("Description").Value = rsDefaults.Fields("Description").Value
                    .UpdateBatch adAffectCurrent
                    rsDefaults.MoveNext
                Loop
            End If
            rsDefaults.Close
            Set rsDefaults = Nothing
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetOutputRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub
    
'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetExposureTimeRecordset                              **
'**                                                                        **
'**  Description..:  This routine returns the Exposure Time Recordset.     **
'**                                                                        **
'****************************************************************************
Public Sub GetExposureTimeRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsExposureTime
        SQLcmd = "SELECT * FROM ExposureTimes WHERE PrinterName='" & PrinterName & "';"
        AppLog InfoMsg, "OpenDatabase,Opening Exposure Time Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("Enabled").Value = True
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetExposureTimeRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetExposureTimeHistoryRecordset                       **
'**                                                                        **
'**  Description..:  This routine returns the Exposure Time Recordset.     **
'**                                                                        **
'****************************************************************************
Public Sub GetExposureTimeHistoryRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    SQLcmd = "SELECT * FROM ExposureTimeHistory WHERE PrinterName='" & PrinterName & "';"
    AppLog InfoMsg, "OpenDatabase,Opening Exposure Time History Recordset as " & SQLcmd
    If rsExposureTimeHistory.State = adStateOpen Then
        rsExposureTimeHistory.Close
    End If
    rsExposureTimeHistory.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetExposureTimeHistoryRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetHotFolderRecordset                                 **
'**                                                                        **
'**  Description..:  This routine returns the Hot Folder Recordset.        **
'**                                                                        **
'****************************************************************************
Public Sub GetHotFolderRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    SQLcmd = "SELECT PrinterName,HotFolderPath,FolderType,FolderEnabled,Status,NumberOfFiles FROM HotFolders WHERE PrinterName='" & PrinterName & "'"
    AppLog InfoMsg, "OpenDatabase,Opening Hot Folder Recordset as " & SQLcmd
    If rsHotFolders.State = adStateOpen Then
        rsHotFolders.Close
    End If
    rsHotFolders.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetHotFolderRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetPrinterStatisticsRecordset                         **
'**                                                                        **
'**  Description..:  This routine returns the Statistics Recordset.        **
'**                                                                        **
'****************************************************************************
Public Sub GetPrinterStatisticsRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    SQLcmd = "SELECT * FROM PrinterStatistics WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,StatisticNumber;"
    AppLog InfoMsg, "OpenDatabase,Opening Printer Statistics Recordset as " & SQLcmd
    If rsPrinterStatistics.State = adStateOpen Then
        rsPrinterStatistics.Close
    End If
    rsPrinterStatistics.Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    ValidateStatistics
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetPrinterStatisticsRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetDailyCalibrationRecordset                          **
'**                                                                        **
'**  Description..:  This routine returns the Color Calibration Recordset. **
'**                                                                        **
'****************************************************************************
Public Sub GetDailyCalibrationRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsDailyCalibration
        SQLcmd = "SELECT * FROM DailyCalibration WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,BlockNum;"
        AppLog InfoMsg, "OpenDatabase,Opening Daily Calibration Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 1
            .Fields("ExposureName").Value = "Densitometer"
            .UpdateBatch adAffectCurrent
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 2
            .Fields("ExposureName").Value = "Lamphouse"
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetDailyCalibrationRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetRingAroundRecordset                                **
'**                                                                        **
'**  Description..:  This routine returns the RingAround Recordset.        **
'**                                                                        **
'****************************************************************************
Public Sub GetRingAroundRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsRingArounds
        SQLcmd = "SELECT * FROM RingArounds WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,BlockNum;"
        AppLog InfoMsg, "OpenDatabase,Opening Ring Around Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 1
            .Fields("ExposureName").Value = "Normal"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 2
            .Fields("ExposureName").Value = "Red"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 3
            .Fields("ExposureName").Value = "Green"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 4
            .Fields("ExposureName").Value = "Blue"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 5
            .Fields("ExposureName").Value = "Cyan"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 6
            .Fields("ExposureName").Value = "Magenta"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 7
            .Fields("ExposureName").Value = "Yellow"
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetRingAroundRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetRingAroundAutoScanRecordset                        **
'**                                                                        **
'**  Description..:  This routine returns the RingAround Scan Recordset.   **
'**                                                                        **
'****************************************************************************
Public Sub GetRingAroundAutoScanRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsRingAroundsAutoScan
        SQLcmd = "SELECT * FROM RingAroundsAutoScan WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,BlockNum;"
        AppLog InfoMsg, "OpenDatabase,Opening Ring Around Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 1
            .Fields("ExposureName").Value = "Normal"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 2
            .Fields("ExposureName").Value = "Red"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 3
            .Fields("ExposureName").Value = "Green"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 4
            .Fields("ExposureName").Value = "Blue"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 5
            .Fields("ExposureName").Value = "Cyan"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 6
            .Fields("ExposureName").Value = "Magenta"
            .UpdateBatch adAffectCurrent
            
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("BlockNum").Value = 7
            .Fields("ExposureName").Value = "Yellow"
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetRingAroundAutoScanRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetOffsetRecordset                                    **
'**                                                                        **
'**  Description..:  This routine returns the Offset Recordset.            **
'**                                                                        **
'****************************************************************************
Public Sub GetOffsetRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsOffsets
        SQLcmd = "SELECT * FROM OffsetCalibrations WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,OffsetNum"
        AppLog InfoMsg, "OpenDatabase,Opening Offset Calibration Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then                    'There are no offsets defined
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("OffsetNum").Value = 1          'This is the clear offset
            .Fields("OffsetDate").Value = Now
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetOffsetRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetLutRecordset                                       **
'**                                                                        **
'**  Description..:  This routine returns the LUT Recordset.               **
'**                                                                        **
'****************************************************************************
Public Sub GetLutRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsLuts
        SQLcmd = "SELECT * FROM LutCalibrations WHERE PrinterName='" & PrinterName & "' ORDER BY PrinterName,LutNum"
        AppLog InfoMsg, "OpenDatabase,Opening Lut Calibration Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then                 'There are no offsets defined
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("LutNum").Value = 1          'This is the clear Lut
            .Fields("LutDate").Value = Now
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetLutRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetBackWriterRecordset                                **
'**                                                                        **
'**  Description..:  This routine returns the BackWriter Settings Recordset**
'**                                                                        **
'****************************************************************************
Public Sub GetBackWriterRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsBackWriterSettings
        SQLcmd = "SELECT * FROM PrinterSettings WHERE PrinterName='" & PrinterName & "' AND SettingNumber>=10 AND SettingNumber<20 ORDER BY PrinterName,SettingNumber"
        AppLog InfoMsg, "OpenDatabase,Opening BackWriter Settings Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetBackWriterRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetEmulsionsRecordset                                 **
'**                                                                        **
'**  Description..:  This routine returns the Emulsion Table Recordset.    **
'**                                                                        **
'****************************************************************************
Public Sub GetEmulsionsRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsEmulsions
        SQLcmd = "SELECT * FROM PaperEmulsions;"
        AppLog InfoMsg, "OpenDatabase,Opening Emulsion Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            .AddNew
            .Fields("EmulsionNumber").Value = 1
            .Fields("EmulsionBrand").Value = "Kodak"
            .Fields("EmulsionName").Value = "Supra"
            .Fields("EmulsionCode").Value = "Supra"
            .Fields("DateModified").Value = Now
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
            With rsEmulsionData
                SQLcmd = "SELECT * FROM PaperEmulsionData WHERE EmulsionNumber=" & rsEmulsions.Fields("EmulsionNumber").Value & " ORDER BY EmulsionNumber,DensityNumber;"
                AppLog InfoMsg, "OpenDatabase,Opening Emulsion Data Recordset as " & SQLcmd
                If .State = adStateOpen Then
                    .Close
                End If
                .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
            End With
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 1, 0, 0, 0, 2.4, 2.4, 2.25
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 2, 8, 8, 8, 2.249948, 2.249948, 2.136864
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 3, 16, 16, 16, 2.101988, 2.101988, 2.024259
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 4, 25, 25, 25, 1.958211, 1.958211, 1.912715
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 5, 33, 33, 33, 1.82071, 1.82071, 1.802762
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 6, 41, 41, 41, 1.691575, 1.691575, 1.69493
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 7, 49, 49, 49, 1.5729, 1.5729, 1.58975
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 8, 58, 58, 58, 1.466774, 1.466774, 1.487753
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 9, 66, 66, 66, 1.373384, 1.373384, 1.390165
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 10, 74, 74, 74, 1.285961, 1.285961, 1.298813
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 11, 82, 82, 82, 1.203789, 1.203789, 1.212948
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 12, 90, 90, 90, 1.126272, 1.126272, 1.131947
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 13, 99, 99, 99, 1.052912, 1.052912, 1.05529
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 14, 107, 107, 107, 0.983284, 0.983284, 0.982533
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 15, 115, 115, 115, 0.917028, 0.917028, 0.913299
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 16, 123, 123, 123, 0.853831, 0.853831, 0.847262
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 17, 132, 132, 132, 0.793425, 0.793425, 0.784141
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 18, 140, 140, 140, 0.735572, 0.735572, 0.723688
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 19, 148, 148, 148, 0.680066, 0.680066, 0.665687
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 20, 156, 156, 156, 0.626724, 0.626724, 0.609947
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 21, 165, 165, 165, 0.575382, 0.575382, 0.556298
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 22, 173, 173, 173, 0.525898, 0.525898, 0.50459
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 23, 181, 181, 181, 0.47814, 0.47814, 0.454686
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 24, 189, 189, 189, 0.431993, 0.431993, 0.406464
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 25, 197, 197, 197, 0.387351, 0.387351, 0.359816
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 26, 206, 206, 206, 0.344119, 0.344119, 0.314641
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 27, 214, 214, 214, 0.302211, 0.302211, 0.27085
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 28, 222, 222, 222, 0.261549, 0.261549, 0.22836
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 29, 230, 230, 230, 0.22206, 0.22206, 0.187096
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 30, 239, 239, 239, 0.183678, 0.183678, 0.146989
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 31, 247, 247, 247, 0.146343, 0.146343, 0.107977
            AddEmulsionData rsEmulsions.Fields("EmulsionNumber").Value, 32, 255, 255, 255, 0.11, 0.11, 0.07
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetEmulsionsRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

Public Sub GetEmulsionDataRecordset()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String
    With rsEmulsionData
        SQLcmd = "SELECT * FROM PaperEmulsionData WHERE EmulsionNumber=" & rsEmulsions.Fields("EmulsionNumber").Value & " ORDER BY EmulsionNumber,DensityNumber;"
        AppLog InfoMsg, "OpenDatabase,Opening Emulsion Data Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetEmulsionDataRecordset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub


Private Sub AddEmulsionData(EmulsionNum As Integer, DensityNum As Integer, RedValue As Integer, GreenValue As Integer, BlueValue As Integer, RedDensity As Double, GreenDensity As Double, BlueDensity As Double)
    On Error GoTo ErrorHandler
    With rsEmulsionData
        .AddNew
        .Fields("EmulsionNumber").Value = EmulsionNum
        .Fields("DensityNumber").Value = DensityNum
        .Fields("RedValue").Value = RedValue
        .Fields("GreenValue").Value = GreenValue
        .Fields("BlueValue").Value = BlueValue
        .Fields("RedDensity").Value = RedDensity
        .Fields("GreenDensity").Value = GreenDensity
        .Fields("BlueDensity").Value = BlueDensity
        .UpdateBatch adAffectCurrent
        .Resync adAffectAllChapters
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "AddEmulsionData, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddNewLUT                                             **
'**                                                                        **
'**  Description..:  This routine Adds a new LUT.                          **
'**                                                                        **
'****************************************************************************
Public Sub AddNewLUT()
    On Error GoTo ErrorHandler
    With rsLuts
        .AddNew
        .Fields("PrinterName").Value = PrinterName
        .Fields("LutNum").Value = GetNextLutNumber
        .Fields("LutDate").Value = Now
        .UpdateBatch adAffectCurrent
        .Resync adAffectAllChapters
    End With
    UpdateLutDensiValues
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "AddNewLUT, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateLutDensiValues                                  **
'**                                                                        **
'**  Description..:  This routine updates densitometer values.             **
'**                                                                        **
'****************************************************************************
Public Sub UpdateLutDensiValues()
    On Error GoTo ErrorHandler
    Dim SQLcmd As String, BlockNum As Integer
    With rsLutDensiValues
        SQLcmd = "SELECT * FROM LutDensitometerValues WHERE PrinterName='" & PrinterName & "' AND LutNum=" & rsLuts.Fields("LutNum").Value & " ORDER BY PrinterName,BlockNum;"
        AppLog InfoMsg, "OpenDatabase,Opening LUT Densitometer Values Recordset as " & SQLcmd
        If .State = adStateOpen Then
            .Close
        End If
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If rsLutDensiValues.RecordCount = 0 Then
            For BlockNum = 1 To MaxLutBlock
                .AddNew
                .Fields("PrinterName").Value = PrinterName
                .Fields("LutNum").Value = rsLuts.Fields("LutNum").Value
                .Fields("BlockNum").Value = BlockNum
                .Fields("DensRed").Value = 0
                .Fields("DensGreen").Value = 0
                .Fields("DensBlue").Value = 0
                .UpdateBatch adAffectCurrent
            Next
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "UpdateLutDensiValues, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetNextLutNumber                                      **
'**                                                                        **
'**  Description..:  This routine returns the next seq. LUT Number.        **
'**                                                                        **
'****************************************************************************
Private Function GetNextLutNumber() As Long
    On Error GoTo ErrorHandler
    Dim rsNum As ADODB.Recordset, SQLcmd As String
    Set rsNum = New ADODB.Recordset
    SQLcmd = "SELECT MAX(LutNum) AS NextNum FROM LutCalibrations WHERE PrinterName='" & PrinterName & "';"
    AppLog InfoMsg, "OpenDatabase,Getting next lut number as " & SQLcmd
    With rsNum
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            GetNextLutNumber = 1
        Else
            GetNextLutNumber = .Fields("NextNum").Value + 1
        End If
    End With
    rsNum.Close
    Set rsNum = Nothing
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetNextLutNumber, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SetCurrentLUT                                         **
'**                                                                        **
'**  Description..:  This routine sets the current LUT.                    **
'**                                                                        **
'****************************************************************************
Public Function SetCurrentLUT() As Boolean
    On Error GoTo ErrorHandler
    Dim Rec As Variant
    With rsLuts
        If .RecordCount > 0 Then
            Rec = .Bookmark
            .MoveFirst
            Do While Not .EOF
                .Fields("Selected").Value = 0
                .UpdateBatch adAffectCurrent
                .MoveNext
            Loop
            .MoveFirst
            .Bookmark = Rec
            .Fields("Selected").Value = -1
            .UpdateBatch adAffectCurrent
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "SetCurrentLUT, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddNewOffset                                          **
'**                                                                        **
'**  Description..:  This routine adds a new offset record.                **
'**                                                                        **
'****************************************************************************
Public Sub AddNewOffset()
    On Error GoTo ErrorHandler
    With DB.rsOffsets
        .AddNew
        .Fields("PrinterName").Value = PrinterName
        .Fields("OffsetNum").Value = GetNextOffsetNumber
        .Fields("OffsetDate").Value = Now
        .Fields("Selected").Value = 0
        .UpdateBatch adAffectCurrent
        .Resync adAffectAllChapters
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "AddNewOffset, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddNewSize                                            **
'**                                                                        **
'**  Description..:  This routine adds a new print size record.            **
'**                                                                        **
'****************************************************************************
Public Sub AddNewSize()
    On Error GoTo ErrorHandler
    With DB.rsPrintSizes
        .AddNew
        .Fields("PrinterName").Value = PrinterName
        .Fields("PrintSize").Value = "New"
        .Fields("Enable").Value = True
        .Fields("AspectRatio").Value = 1#
        .UpdateBatch adAffectCurrent
        .Resync adAffectAllChapters
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "AddNewSize, Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetNextOffsetNumber                                   **
'**                                                                        **
'**  Description..:  This routine returns the next seq offset number.      **
'**                                                                        **
'****************************************************************************
Private Function GetNextOffsetNumber() As Long
    On Error GoTo ErrorHandler
    Dim rsNum As ADODB.Recordset, SQLcmd As String
    Set rsNum = New ADODB.Recordset
    SQLcmd = "SELECT MAX(OffsetNum) AS NextNum FROM OffsetCalibrations WHERE PrinterName='" & PrinterName & "';"
    AppLog InfoMsg, "OpenDatabase,Getting next lut number as " & SQLcmd
    With rsNum
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            GetNextOffsetNumber = 1
        Else
            GetNextOffsetNumber = .Fields("NextNum").Value + 1
        End If
    End With
    rsNum.Close
    Set rsNum = Nothing
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetNextOffsetNumber,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetNextPrinterID                                      **
'**                                                                        **
'**  Description..:  This routine returns the next seq printer ID.         **
'**                                                                        **
'****************************************************************************
Public Function GetNextPrinterID() As String
    '---- Unique id starts in char 6
    On Error GoTo ErrorHandler
    Dim rsNum As ADODB.Recordset, SQLcmd As String
    Set rsNum = New ADODB.Recordset
    SQLcmd = "SELECT PrinterName AS NextNum FROM PrinterList WHERE PrinterName<>'Default' ORDER BY PrinterName"
    AppLog InfoMsg, "OpenDatabase,Getting next printer id " & SQLcmd
    With rsNum
        .Open SQLcmd, cnn, adOpenDynamic, adLockOptimistic, adCmdText
        If .RecordCount = 0 Then
            GetNextPrinterID = "DVP2_0001"
        Else
            .MoveLast
            GetNextPrinterID = "DVP2_" & Format(Val(Trim(Mid(.Fields("NextNum").Value, 6, 4))) + 1, "0000")
        End If
    End With
    rsNum.Close
    Set rsNum = Nothing
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetNextPrinterID,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SetCurrentOffset                                      **
'**                                                                        **
'**  Description..:  This routine sets the current printer offset.         **
'**                                                                        **
'****************************************************************************
Public Function SetCurrentOffset() As Boolean
    On Error GoTo ErrorHandler
    Dim Rec As Variant
    With rsOffsets
        If .RecordCount > 0 Then
            Rec = .Bookmark
            .MoveFirst
            Do While Not .EOF
                .Fields("Selected").Value = 0
                .MoveNext
            Loop
            .MoveFirst
            .Bookmark = Rec
            .Fields("Selected").Value = -1
            .UpdateBatch adAffectCurrent
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "SetCurrentOffset,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetServerSetting                                      **
'**                                                                        **
'**  Description..:  This routine returns settings from the database.      **
'**                                                                        **
'****************************************************************************
Public Function GetServerSetting(SettingNum As Integer) As String
    On Error GoTo ErrorHandler
    Dim SettingValue As String, SettingName As String
    With rsSettings
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "SettingNumber=" & SettingNum, 0, adSearchForward, 0
            If Not .EOF Then
                SettingValue = Trim(.Fields("SettingValue").Value & "")
                SettingName = Trim(.Fields("SettingName").Value & "")
                AppLog InfoMsg, "GetServerSetting," & SettingNum & "," & SettingName & "," & SettingValue
                GetServerSetting = SettingValue
            Else
                AppLog ErrorMsg, "GetServerSetting," & SettingNum & ",not defined."
                GetServerSetting = ""
            End If
        Else
            AppLog ErrorMsg, "GetServerSetting,No Printer Settings defined."
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetServerSetting,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetServerOption                                       **
'**                                                                        **
'**  Description..:  This routine returns options from the database.       **
'**                                                                        **
'****************************************************************************
Public Function GetServerOption(OptionNum As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim OptionValue As Boolean, OptionName As String
    With rsOptions
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "OptionNumber=" & OptionNum, 0, adSearchForward, 0
            If Not .EOF Then
                OptionValue = .Fields("OptionValue").Value
                OptionName = .Fields("OptionName").Value
                AppLog InfoMsg, "GetServerOption," & OptionNum & "," & OptionName & "," & OptionValue
                GetServerOption = OptionValue
            Else
                AppLog ErrorMsg, "GetServerOption," & OptionNum & ",not defined."
                GetServerOption = ""
            End If
        Else
            AppLog ErrorMsg, "GetServerOption,No Printer Options defined."
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetServerOption,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CompactDatabase                                       **
'**                                                                        **
'**  Description..:  This routine compacts the print queue database.       **
'**                                                                        **
'****************************************************************************
Private Function CompactDatabase(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim ConnectString As String, TargetString As String, mPC As PerformanceCounter
    Set mPC = New PerformanceCounter
    mPC.StartTimer True
    AppLog InfoMsg, "CompactDatabase,Compacting " & FileName
    ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";Persist Security Info=False"
    TargetString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ".tmp"
    Dim JRO As JRO.JetEngine
    Set JRO = New JRO.JetEngine
    JRO.CompactDatabase ConnectString, TargetString
    If FileSystemHandle.FileExists(FileName & ".bak") Then
        Kill FileName & ".bak"
    End If
    If FileSystemHandle.FileExists(FileName) Then
        Name FileName As FileName & ".bak"
    End If
    If FileSystemHandle.FileExists(FileName & ".tmp") Then
        Name FileName & ".tmp" As FileName
    Else
        AppLog ErrorMsg, "CompactDatabase,Error compacting print queue database - .tmp file not found."
    End If
    AppLog InfoMsg, "CompactDatabase,Timed," & Format(mPC.StopTimer, "####.####")
    Exit Function
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
    AppLog ErrorMsg, "CompactDatabase,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetPrintSize                                          **
'**                                                                        **
'**  Description..:  This routine returns print size for given aspect ratio**
'**                                                                        **
'****************************************************************************
Public Function GetPrintSize(hRes As Long, vRes As Long) As String
    On Error GoTo ErrorHandler
    Dim RatioGrid As String, SizeRow As Long, AspectRatio As String
    GetPrintSize = ""
    If hRes = 0 Or vRes = 0 Then
        Exit Function
    End If
    AspectRatio = Format(vRes / hRes, "#.###")
    With rsPrintSizes
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "AspectRatio=" & AspectRatio, 0, adSearchForward, 0
            If Not .EOF Then
                GetPrintSize = .Fields("PrintSize").Value
            End If
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetPrintSize,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetLeadError                                          **
'**                                                                        **
'**  Description..:  This routine returns LeadTools Error Description.     **
'**                                                                        **
'****************************************************************************
Public Function GetLeadError(ErrorNo As Integer) As String
    On Error GoTo ErrorHandler
    With rsLeadErrorCodes
        .MoveFirst
        .Find "ErrorNumber=" & ErrorNo
        If Not .EOF Then
            GetLeadError = Trim(.Fields("ErrorDescription").Value)
        Else
            GetLeadError = "Unkown Error"
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetLeadError,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  RemovePrintedImages                                   **
'**                                                                        **
'**  Description..:  This routine removes printed images from the print que**
'**                                                                        **
'****************************************************************************
Public Sub RemovePrintedImages()
    On Error GoTo ErrorHandler
    Dim Done As Boolean, HasImage As Boolean, HoldTheFile As Boolean
    Done = False
    With rsPrintQue
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If .Fields("HoldEnable").Value = 0 And .Fields("Status").Value = "Printed." Then
                    .Delete adAffectCurrent
                End If
                .MoveNext
            Loop
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "RemovePrintedImages,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  RemoveImageFromHotFolder                              **
'**                                                                        **
'**  Description..:  This routine deletes images from the HotFolder.       **
'**                                                                        **
'****************************************************************************
Public Sub RemoveImageFromHotFolder(HotFolderName As String, FileName As String)
    On Error GoTo ErrorHandler
    With rsHotFolders
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "HotFolderPath='" & Left(HotFolderName, InStrRev(HotFolderName, "\") - 1) & "'"
            If Not .EOF Then
                If FileSystemHandle.FileExists(FileName) = True Then
                    Kill FileName
                    .Fields("NumberOfFiles").Value = .Fields("NumberOfFiles").Value - 1
                    .UpdateBatch adAffectCurrent
                End If
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "RemoveImageFromHotFolder,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ValidateSettings                                      **
'**                                                                        **
'**  Description..:  This routine validates printer settings in database.  **
'**                                                                        **
'****************************************************************************
Public Sub ValidateSettings()
    CheckSetting RecCalibrationImagePath, "Calibration File Path", "C:\DVP2_R2\Calibration Images"
    CheckSetting RecLogFilePath, "Log File Path", "C:\DVP2_R2\Log Files"
    CheckSetting RecRamDiskPath, "RAM Disk Path", "S:"
    CheckSetting RecAdminPassword, "Administrative Password", " "
    CheckSetting RecMotorPPISetting, "Paper Advance Motor PPI", "236"
    CheckSetting RecPaperAdvanceLength, "Paper Advance Manual Length", "10"
    CheckSetting RecPaperAdvanceTearOffLength, "Paper Advance TearOff Length", "30"
    CheckSetting RecQualityControlPassword, "Quality Control Password", " "
    CheckSetting RecBW_Text1, "BackWriter Text 1", " "
    CheckSetting RecBW_Text2, "BackWriter Text 2", " "
    CheckSetting RecBW_Char_Delay, "Character Delay", "1200"
    CheckSetting RecBW_Dot_OnTime, "Dot On Time", "500"
    CheckSetting RecBW_Dot_OffTime, "Dot Off Time", "500"
    CheckSetting RecBW_MaxChars, "Maximum # Characters", "100"
    CheckSetting RecBW_AdvDelay, "Delay from start of advance", "100"
    CheckSetting RecBW_RibbonSpeed, "Ribbon Motor Speed", "200"
    CheckSetting RecBW_FontSize, "Character Height (dots)", "9"
    CheckSetting RecBW_MotorWait, "Wait for stepper motor input", "Yes"
    CheckSetting RecMaskHomePrintCount, "Stepper Mask Home Every # Prints", "0"
    CheckSetting RecMaskStepsOnClick, "Stepper Mask Steps on Button Click", "500"
    CheckSetting RecMaskLeftLimit, "Stepper Mask Left Limit", "-225000"
    CheckSetting RecMaskRightLimit, "Stepper Mask Right Limit", "74000"
    CheckSetting RecMaskHoldingTorque, "Stepper Mask Holding Torque (%, Max=60)", "50"
    CheckSetting RecPunchSolenoidTime, "Punch Solenoid Enable Time (ms)", "100"
    CheckSetting RecPunchEngageTime, "Punch Engage Time (ms)", "100"
    CheckSetting RecPunchDisengageTime, "Punch Disengage Time (ms)", "100"
    CheckSetting RecPunchPkgExtendTime, "Punch Pkg Extend Time (ms)", "100"
    CheckSetting RecDigitalCutSize, "Digital Cut Code Size (Pixels)", "16"
    CheckSetting RecDigitalCut8x10Offset, "Digital Cut Code 8x10 Pixels From Left Edge", "0"
    CheckSetting RecDigitalCut11x14Offset, "Digital Cut Code 11x14 Pixels From Top Edge", "0"
    CheckSetting RecPurgeLogDays, "Purge log files after # days", "30"
    CheckSetting RecPurgeExposureHistoryDays, "Purge exposure time history after # days", "30"
    CheckSetting RecPurgeQueHistoryDays, "Purge print queue history after # days", "30"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckSetting                                          **
'**                                                                        **
'**  Description..:  This routine validates printer settings in database.  **
'**                                                                        **
'****************************************************************************
Public Sub CheckSetting(SettingNum As Integer, SettingName As String, DefaultValue As String)
    On Error GoTo ErrorHandler
    With rsSettings
        If .RecordCount > 0 Then .MoveFirst
        .Find "SettingNumber=" & SettingNum
        If .EOF Then
            AppLog InfoMsg, "CheckSetting,Added Setting " & SettingNum & " Name=" & SettingName
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("SettingNumber").Value = SettingNum
            .Fields("SettingName").Value = SettingName
            .Fields("SettingValue").Value = DefaultValue
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "CheckSetting,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ValidateOptions                                       **
'**                                                                        **
'**  Description..:  This routine validates printer options in database.   **
'**                                                                        **
'****************************************************************************
Public Sub ValidateOptions()
    CheckOption RecDebugMode, "Debug Mode", True
    CheckOption RecNordConversion, "Stepper Mask Installed", True
    CheckOption RecDoorSwitchInstalled, "Daylight Load", False
    CheckOption RecBackWritersInstalled, "Single Backwriter Installed", False
    CheckOption RecDualBackWritersInstalled, "Dual Backwriter Installed", False
    CheckOption RecPunchEnable, "Single Punch Installed", False
    CheckOption RecPackagePunchEnable, "Package Punch Installed", False
    CheckOption RecRS422MotorConfiguration, "RS-422 Motor Configuration", True
    CheckOption RecPlatenCylinderEnable, "Platen Cylinder Installed", False
    CheckOption RecQueGetQty, "Render Queue - Get Quantity From File Name", False
    CheckOption RecQueGetPunch, "Render Queue - Get Punch Code From File Name", False
    CheckOption RecQueGetText, "Render Queue - Get BackWriter Text from File Name", False
    CheckOption RecQueSort, "Render Queue - Sort By File Name", False
    CheckOption RecFujiFileNameEncoding, "Render Queue - Fuji File Naming", False
    CheckOption RecBackWriterOverride, "Use Fixed BackWriter Text From Maintenance Screen", False
    CheckOption RecApplyGrayBorder, "Apply Gray Borders on LCD Image", True
    CheckOption RecApplyMullerSohnLUT, "Apply MullerSOHN LUT", True
    CheckOption RecApplyPictoLUT, "Apply Pictographics LUT", False
    CheckOption RecRenderCutCode, "Render Digital Cut Code", False
    CheckOption RecEnableQC, "Enable Q.C. in Printer Console", True
    CheckOption RecQueHistory, "Enable Print Queue History", True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckOption                                           **
'**                                                                        **
'**  Description..:  This routine validates printer options in database.   **
'**                                                                        **
'****************************************************************************
Public Sub CheckOption(OptionNum As Integer, OptionName As String, DefaultValue As Boolean)
    On Error GoTo ErrorHandler
    With rsOptions
        If .RecordCount > 0 Then .MoveFirst
        .Find "OptionNumber=" & OptionNum
        If .EOF Then
            AppLog InfoMsg, "CheckOption,Added Option " & OptionNum & " Name=" & OptionName
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("OptionNumber").Value = OptionNum
            .Fields("OptionName").Value = OptionName
            .Fields("OptionValue").Value = DefaultValue
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "CheckOption,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ValidateStatistics                                    **
'**                                                                        **
'**  Description..:  This routine validates printer stats in database.     **
'**                                                                        **
'****************************************************************************
Public Sub ValidateStatistics()
    CheckStatistic RecStatTotalImages, "Total Images"
    CheckStatistic RecStatTotalExposures, "Total Exposures"
    CheckStatistic RecStatExposuresPerSecond, "Exposures/Second"
    CheckStatistic RecStatAverageExposureTime, "Average Exposure Time (Sec.)"
    CheckStatistic RecStatAverageServerTime, "Average Server Time (Sec.)"
    CheckStatistic RecStatAverageAdvanceTime, "Average Advance Time (Sec.)"
    CheckStatistic RecStatTimeRunning, "Time Running (Hrs.)"
    CheckStatistic RecStatPaperUsed, "Paper Used (ft.)"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckStatistic                                        **
'**                                                                        **
'**  Description..:  This routine validates printer stats in database.     **
'**                                                                        **
'****************************************************************************
Public Sub CheckStatistic(StatisticNum As Integer, StatisticName As String)
    On Error GoTo ErrorHandler
    With rsPrinterStatistics
        If .RecordCount > 0 Then .MoveFirst
        .Find "StatisticNumber=" & StatisticNum
        If .EOF Then
            AppLog InfoMsg, "CheckStatistic,Added Statistic " & StatisticNum & " Name=" & StatisticName
            .AddNew
            .Fields("PrinterName").Value = PrinterName
            .Fields("StatisticNumber").Value = StatisticNum
            .Fields("StatisticName").Value = StatisticName
            .Fields("StatisticValue").Value = 0
            .Fields("StatisticEnable").Value = True
            .Fields("DateModified").Value = Now
            .UpdateBatch adAffectCurrent
            .Resync adAffectAllChapters
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "CheckStatistic,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub


Public Sub UpdateStatistics()
    UpdateStatistic RecStatTotalImages, StatTotalImages
    UpdateStatistic RecStatTotalExposures, StatTotalExposures
    UpdateStatistic RecStatExposuresPerSecond, StatExposuresPerSecond
    UpdateStatistic RecStatAverageExposureTime, StatAverageExposureTime
    UpdateStatistic RecStatAverageServerTime, StatAverageServerTime
    UpdateStatistic RecStatAverageAdvanceTime, StatAverageAdvanceTime
    UpdateStatistic RecStatTimeRunning, StatTimeRunning
    UpdateStatistic RecStatPaperUsed, StatPaperUsed
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateStatistic                                       **
'**                                                                        **
'**  Description..:  This routine stores printer stats in database.        **
'**                                                                        **
'****************************************************************************
Public Sub UpdateStatistic(StatisticNum As Integer, StatValue As Double)
    On Error GoTo ErrorHandler
    With rsPrinterStatistics
        If .RecordCount > 0 Then .MoveFirst
        .Find "StatisticNumber=" & StatisticNum
        If Not .EOF Then
            'AppLog InfoMsg, "UpdateStatistic,Updated " & StatisticNum & " Name=" & .Fields("StatisticName").Value
            .Fields("StatisticValue").Value = StatValue
            .Fields("DateModified").Value = Now
            .UpdateBatch adAffectCurrent
        Else
            AppLog ErrorMsg, "UpdateStatistic,Error updating " & StatisticNum
        End If
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "UpdateStatistic,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetStatistics                                         **
'**                                                                        **
'**  Description..:  This routine retrieves printer statistics.            **
'**                                                                        **
'****************************************************************************
Public Function GetStatistic(StatisticNum As Integer) As Double
    On Error GoTo ErrorHandler
    With rsPrinterStatistics
        If .RecordCount > 0 Then .MoveFirst
        .Find "StatisticNumber=" & StatisticNum
        If Not .EOF Then
            AppLog InfoMsg, "GetStatistic,Retrieved " & StatisticNum & " Name=" & .Fields("StatisticName").Value
            GetStatistic = .Fields("StatisticValue").Value
        Else
            AppLog ErrorMsg, "GetStatistic,Error retrieving " & StatisticNum
        End If
    End With
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "GetStatistic,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function


'****************************************************************************
'**                                                                        **
'**  Procedure....:  GetPrinterSetting                                     **
'**                                                                        **
'**  Description..:  This routine returns settings from the database.      **
'**                                                                        **
'****************************************************************************
Public Sub GetPrinterSettings()
    On Error GoTo ErrorHandler
    
    '---- Get Options
    DebugMode = GetServerOption(RecDebugMode)
    StepperMaskInstalled = GetServerOption(RecNordConversion)
    DoorSwitchInstalled = GetServerOption(RecDoorSwitchInstalled)
    BackWritersInstalled = GetServerOption(RecBackWritersInstalled) Or GetServerOption(RecDualBackWritersInstalled)
    PunchEnable = GetServerOption(RecPunchEnable)
    PackagePunchEnable = GetServerOption(RecPackagePunchEnable)
    PlatenCylinderEnable = GetServerOption(RecPlatenCylinderEnable)
    RS422MotorConfiguration = GetServerOption(RecRS422MotorConfiguration)
    PlatenCylinderEnable = GetServerOption(RecPlatenCylinderEnable)
    QueGetQty = GetServerOption(RecQueGetQty)
    QueGetPunch = GetServerOption(RecQueGetPunch)
    QueGetText = GetServerOption(RecQueGetText)
    QueSort = GetServerOption(RecQueSort)
    FujiFileNameEncoding = GetServerOption(RecFujiFileNameEncoding)
    BackWriterOverride = GetServerOption(RecBackWriterOverride)
    ApplyGrayBorder = GetServerOption(RecApplyGrayBorder)
    ApplyMullerSohnLUT = GetServerOption(RecApplyMullerSohnLUT)
    ApplyPictoLUT = GetServerOption(RecApplyPictoLUT)
    If ApplyPictoLUT = True Then
        MaxLutBlock = 72
    Else
        MaxLutBlock = 48
    End If
    RenderCutCode = GetServerOption(RecRenderCutCode)
    EnableQC = GetServerOption(RecEnableQC)
    QueHistory = GetServerOption(RecQueHistory)
    
    '---- Get Settings
    CalibrationImagePath = GetServerSetting(RecCalibrationImagePath)
    LogFilePath = GetServerSetting(RecLogFilePath)
    RamDiskPath = GetServerSetting(RecRamDiskPath)
    AdminPassword = GetServerSetting(RecAdminPassword)
    MotorPPISetting = Val(Trim(GetServerSetting(RecMotorPPISetting)))
    PaperAdvanceLength = Val(Trim(GetServerSetting(RecPaperAdvanceLength)))
    PaperAdvanceTearOffLength = Val(Trim(GetServerSetting(RecPaperAdvanceTearOffLength)))
    QualityControlPassword = GetServerSetting(RecQualityControlPassword)
    BW_Text1 = GetServerSetting(RecBW_Text1)
    BW_Text2 = GetServerSetting(RecBW_Text2)
    BW_Char_Delay = Val(Trim(GetServerSetting(RecBW_Char_Delay)))
    BW_Dot_OnTime = Val(Trim(GetServerSetting(RecBW_Dot_OnTime)))
    BW_Dot_OffTime = Val(Trim(GetServerSetting(RecBW_Dot_OffTime)))
    BW_MaxChars = Val(Trim(GetServerSetting(RecBW_MaxChars)))
    BW_AdvDelay = Val(Trim(GetServerSetting(RecBW_AdvDelay)))
    BW_RibbonSpeed = Val(Trim(GetServerSetting(RecBW_RibbonSpeed)))
    BW_FontSize = Val(Trim(GetServerSetting(RecBW_FontSize)))
    BW_MotorWait = GetServerSetting(RecBW_MotorWait)
    
    MaskHomePrintCount = Val(Trim(GetServerSetting(RecMaskHomePrintCount)))
    MaskStepsOnClick = Val(Trim(GetServerSetting(RecMaskStepsOnClick)))
    MaskLeftLimit = CLng(Trim(GetServerSetting(RecMaskLeftLimit)))
    MaskRightLimit = CLng(Trim(GetServerSetting(RecMaskRightLimit)))
    MaskHoldingTorque = Val(Trim(GetServerSetting(RecMaskHoldingTorque)))
    PunchSolenoidTime = Val(Trim(GetServerSetting(RecPunchSolenoidTime)))
    PunchEngageTime = Val(Trim(GetServerSetting(RecPunchEngageTime)))
    PunchDisengageTime = Val(Trim(GetServerSetting(RecPunchDisengageTime)))
    PunchPkgExtendTime = Val(Trim(GetServerSetting(RecPunchPkgExtendTime)))
    DigitalCutSize = Val(Trim(GetServerSetting(RecDigitalCutSize)))
    DigitalCut8x10Offset = Val(Trim(GetServerSetting(RecDigitalCut8x10Offset)))
    DigitalCut11x14Offset = Val(Trim(GetServerSetting(RecDigitalCut11x14Offset)))
    PurgeLogDays = Val(Trim(GetServerSetting(RecPurgeLogDays)))
    PurgeExposureHistoryDays = Val(Trim(GetServerSetting(RecPurgeExposureHistoryDays)))
    PurgeQueHistoryDays = Val(Trim(GetServerSetting(RecPurgeQueHistoryDays)))
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "GetPrinterSettings,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

Public Sub GetPrinterStatistics()
    '---- Get Statistics
     StatTotalImages = GetStatistic(RecStatTotalImages)
     StatTotalExposures = GetStatistic(RecStatTotalExposures)
     StatExposuresPerSecond = GetStatistic(RecStatExposuresPerSecond)
     StatAverageExposureTime = GetStatistic(RecStatAverageExposureTime)
     StatAverageServerTime = GetStatistic(RecStatAverageServerTime)
     StatAverageAdvanceTime = GetStatistic(RecStatAverageAdvanceTime)
     StatTimeRunning = GetStatistic(RecStatTimeRunning)
     StatPaperUsed = GetStatistic(RecStatPaperUsed)
End Sub


Private Sub ValidateTables()
    On Error GoTo ErrorHandler
    Dim tbl As ADOX.Table
    Dim idx As ADOX.index
    Dim kyForeign As ADOX.Key
    
    '---- Check for existance of paper emulsion table
    If TableExists("PaperEmulsions") = False Then
        Set tbl = New ADOX.Table
        Set idx = New ADOX.index
        With tbl
            .Name = "PaperEmulsions"
            .Columns.Append "EmulsionNumber", adInteger
            .Columns.Append "EmulsionBrand", adVarWChar, 20
            .Columns.Append "EmulsionName", adVarWChar, 50
            .Columns.Append "EmulsionCode", adVarWChar, 50
            .Columns.Append "DateModified", adDate
        End With
        cat.Tables.Append tbl
        idx.Name = "PrimaryKey"
        idx.Columns.Append "EmulsionNumber"
        idx.PrimaryKey = True
        tbl.Indexes.Append idx
        Set tbl = Nothing
        Set idx = Nothing
    End If
    
    If TableExists("PaperEmulsionData") = False Then
        Set tbl = New ADOX.Table
        Set idx = New ADOX.index
        Set kyForeign = New ADOX.Key
        With tbl
            .Name = "PaperEmulsionData"
            .Columns.Append "EmulsionNumber", adInteger
            .Columns.Append "DensityNumber", adInteger
            .Columns.Append "RedValue", adInteger
            .Columns.Append "GreenValue", adInteger
            .Columns.Append "BlueValue", adInteger
            .Columns.Append "RedDensity", adDouble
            .Columns.Append "GreenDensity", adDouble
            .Columns.Append "BlueDensity", adDouble
        End With
        cat.Tables.Append tbl
        idx.Name = "PrimaryKey"
        idx.Columns.Append "EmulsionNumber"
        idx.Columns.Append "DensityNumber"
        idx.PrimaryKey = True
        tbl.Indexes.Append idx
        
        kyForeign.Name = "EmulsionKey"
        kyForeign.Type = adKeyForeign
        kyForeign.RelatedTable = "PaperEmulsions"
        kyForeign.Columns.Append "EmulsionNumber"
        kyForeign.Columns("EmulsionNumber").RelatedColumn = "EmulsionNumber"
        kyForeign.UpdateRule = adRICascade
        
        cat.Tables("PaperEmulsionData").Keys.Append kyForeign
        
        Set tbl = Nothing
        Set idx = Nothing
        Set kyForeign = Nothing
    End If
    Set tbl = Nothing
    Set idx = Nothing
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "ValidateTables,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

Private Function TableExists(TableName As String) As Boolean
    On Error GoTo ErrorHandler
    TableExists = False
    Dim TblNum As Integer
    For TblNum = 0 To cat.Tables.Count - 1
        If cat.Tables(TblNum).Name = TableName Then
            TableExists = True
            Exit Function
        End If
    Next
    Exit Function
ErrorHandler:
    AppLog ErrorMsg, "TableExists,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Function



