VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00130003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltocx13n.ocx"
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form OffsetControlForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12180
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter OffsetSplitter 
      Height          =   12180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14850
      _ExtentX        =   26194
      _ExtentY        =   21484
      _Version        =   262144
      AutoSize        =   1
      SplitterBarJoinStyle=   0
      PaneTree        =   "OffsetControlForm.frx":0000
      Begin ActiveTabs.SSActiveTabs OffsetTab 
         Height          =   12120
         Left            =   4785
         TabIndex        =   2
         Top             =   30
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   21378
         _Version        =   262144
         TabCount        =   2
         BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TagVariant      =   ""
         Tabs            =   "OffsetControlForm.frx":0052
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
            Height          =   11700
            Left            =   30
            TabIndex        =   3
            Top             =   390
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   20638
            _Version        =   262144
            TabGuid         =   "OffsetControlForm.frx":00DF
            Begin LEADLib.LEAD CalcImage 
               Height          =   10695
               Left            =   30
               TabIndex        =   4
               Top             =   30
               Width           =   10545
               _Version        =   65540
               _ExtentX        =   18600
               _ExtentY        =   18865
               _StockProps     =   229
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   1
               ScaleHeight     =   711
               ScaleWidth      =   701
               DataField       =   ""
               BitmapDataPath  =   ""
               AnnDataPath     =   ""
               PanWinTitle     =   "PanWindow"
               CLeadCtrl       =   0
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
            Height          =   11700
            Left            =   30
            TabIndex        =   5
            Top             =   390
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   20638
            _Version        =   262144
            TabGuid         =   "OffsetControlForm.frx":0107
            Begin LEADLib.LEAD OffsetImage 
               Height          =   10755
               Left            =   30
               TabIndex        =   6
               Top             =   60
               Width           =   10575
               _Version        =   65540
               _ExtentX        =   18653
               _ExtentY        =   18971
               _StockProps     =   229
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   1
               ScaleHeight     =   715
               ScaleWidth      =   703
               DataField       =   ""
               BitmapDataPath  =   ""
               AnnDataPath     =   ""
               PanWinTitle     =   "PanWindow"
               CLeadCtrl       =   0
            End
            Begin MSComctlLib.ProgressBar OffsetProgress 
               Height          =   315
               Left            =   30
               TabIndex        =   7
               Top             =   10890
               Visible         =   0   'False
               Width           =   3885
               _ExtentX        =   6853
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   1
            End
         End
      End
      Begin UltraGrid.SSUltraGrid OffsetGrid 
         Height          =   12120
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   21378
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   68157460
         BorderStyle     =   5
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Override        =   "OffsetControlForm.frx":012F
         Caption         =   "LCD Offset Calibrations"
      End
   End
   Begin ActiveToolBars.SSActiveToolBars ToolBarMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "OffsetControlForm.frx":01AD
      ToolBars        =   "OffsetControlForm.frx":41E1
   End
End
Attribute VB_Name = "OffsetControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2 Quality Control                                     **
'**                                                                        **
'** Module.....: OffsetControl                                             **
'**                                                                        **
'** Description: This module provides LCD Offset Calibration.              **
'**                                                                        **
'** History....:                                                           **
'**    10/02/03 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Private Const CalcName As String = "\FRMtoBMP.exe"
Private Const OffsetFolder As String = "\Offset"
Private Const OffsetHistoryFolder As String = "\History"

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Public Function Setup()
    On Error GoTo ErrorHandler
    '---- Make sure path structure for offsets exists
    Dim TargetPath As String
    With FileSystemHandle
        TargetPath = OffsetFilePath & Trim(PrinterName) & "\"
        If .FolderExists(TargetPath) = False Then
            .CreateFolder TargetPath
        End If
        TargetPath = OffsetFilePath & Trim(PrinterName) & OffsetFolder & "\"
        If .FolderExists(TargetPath) = False Then
            .CreateFolder TargetPath
            TargetPath = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\"
            .CreateFolder TargetPath
            ClearOffset True
        End If
    End With
    
    With OffsetGrid
        Set .DataSource = DB.rsOffsets
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Activation = ssActivationActivateOnly
        .Bands(0).Columns(1).Header.Caption = "#"
        .Bands(0).Columns(1).Width = 600
        .Bands(0).Columns(2).Activation = ssActivationActivateOnly
        .Bands(0).Columns(2).Header.Caption = "Date Scanned"
        .Bands(0).Columns(2).Width = 2000
        .Bands(0).Columns(3).Activation = ssActivationActivateOnly
        .Bands(0).Columns(3).Header.Caption = "Selected"
        .Bands(0).Columns(3).Width = 900
    End With
    DB.rsOffsets.MoveLast
    DisplayCalcFile
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

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
        OffsetSplitter.Width = Me.Width - 100
        OffsetTab.Width = OffsetSplitter.Panes(1).Width - 100
        OffsetGrid.Width = OffsetSplitter.Panes(0).Width - 100
        OffsetImage.Width = OffsetTab.Width - 100
        CalcImage.Width = OffsetTab.Width - 100
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ToolBarMain_ToolClick                                 **
'**                                                                        **
'**  Description..:  This routine handles the offset command buttons.      **
'**                                                                        **
'****************************************************************************
Private Sub ToolBarMain_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Dim OffsetNum As Integer
    OffsetNum = DB.rsOffsets.Fields("OffsetNum").Value
    Select Case Tool.ID
        Case "ID_Add"
            MsgBox "Feature coming soon..."
        Case "ID_Erase"
            If OffsetNum = 1 Then
                MsgBox "Cannot erase the clear offset.", vbApplicationModal + vbInformation + vbOKOnly, "NOTICE"
            Else
                If MsgBox("Erase Offset#" & DB.rsOffsets.Fields("OffsetNum").Value & "?", vbApplicationModal + vbQuestion + vbYesNoCancel + vbDefaultButton2, "Are you sure?") = vbYes Then
                    '---- Kill the offset Binary Files
                    If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "r.frm") = True Then
                        Kill OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset" & Trim(Str(OffsetNum)) & "r.frm"
                    End If
                    If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "g.frm") = True Then
                        Kill OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset" & Trim(Str(OffsetNum)) & "g.frm"
                    End If
                    If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "b.frm") = True Then
                        Kill OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset" & Trim(Str(OffsetNum)) & "b.frm"
                    End If
                    If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Calc" & Trim(Str(OffsetNum)) & ".bmp") = True Then
                        Kill OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset_calc" & Trim(Str(OffsetNum)) & ".bmp"
                    End If
                    If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Calc" & Trim(Str(OffsetNum)) & ".txt") = True Then
                        Kill OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset_calc" & Trim(Str(OffsetNum)) & ".txt"
                    End If
                    If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Scan" & Trim(Str(OffsetNum)) & ".bmp") = True Then
                        Kill OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset_scan" & Trim(Str(OffsetNum)) & ".bmp"
                    End If
                    DB.rsOffsets.Delete adAffectCurrent
                    DB.rsOffsets.MoveFirst
                End If
            End If
        Case "ID_Select"
            DB.SetCurrentOffset
            If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "r.frm") = True Then
                If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "g.frm") = True Then
                    If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "b.frm") = True Then
                        FileSystemHandle.CopyFile OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset" & Trim(Str(OffsetNum)) & "r.frm", OffsetFilePath & Trim(PrinterName) & OffsetFolder & "\offsetr.frm"
                        FileSystemHandle.CopyFile OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset" & Trim(Str(OffsetNum)) & "g.frm", OffsetFilePath & Trim(PrinterName) & OffsetFolder & "\offsetg.frm"
                        FileSystemHandle.CopyFile OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\offset" & Trim(Str(OffsetNum)) & "b.frm", OffsetFilePath & Trim(PrinterName) & OffsetFolder & "\offsetb.frm"
                    End If
                End If
            End If
        Case "ID_Scan"
            ScanOffset
        Case "ID_Recalculate"
            DoCalcOffset
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ToolBarMain_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ClearOffset                                           **
'**                                                                        **
'**  Description..:  This routine copies offset files to backup dir.       **
'**                                                                        **
'****************************************************************************
Public Function ClearOffset(ClearFlag As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim TargetPath As String
    With FileSystemHandle
        TargetPath = OffsetFilePath & Trim(PrinterName) & OffsetFolder
        If .FileExists(TargetPath & "\Offsetr.lut") = False Or ClearFlag = True Then
            '--- Copy CLEAR Offset to folder
            .CopyFile SettingsFolder & OffsetFolder & "\ClearOffsetr.frm", TargetPath & "\Offsetr.frm", True
            .CopyFile SettingsFolder & OffsetFolder & "\ClearOffsetg.frm", TargetPath & "\Offsetg.frm", True
            .CopyFile SettingsFolder & OffsetFolder & "\ClearOffsetb.frm", TargetPath & "\Offsetb.frm", True
            
            .CopyFile SettingsFolder & OffsetFolder & "\ClearOffsetr.frm", TargetPath & OffsetHistoryFolder & "\Offset1r.frm", True
            .CopyFile SettingsFolder & OffsetFolder & "\ClearOffsetg.frm", TargetPath & OffsetHistoryFolder & "\Offset1g.frm", True
            .CopyFile SettingsFolder & OffsetFolder & "\ClearOffsetb.frm", TargetPath & OffsetHistoryFolder & "\Offset1b.frm", True
            Sleep 500
            MakeCalcFile
        End If
    End With
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ClearOffset", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ScanOffset                                            **
'**                                                                        **
'**  Description..:  This routine handles scanning of printed offset.      **
'**                                                                        **
'****************************************************************************
Public Sub ScanOffset()
    On Error GoTo ErrorHandler
    Dim RetVal As Integer
    OffsetImage.EnableMethodErrors = False
    OffsetImage.TwainSelect hWnd
    RetVal = OffsetImage.TwainAcquire(hWnd)
    If RetVal <> 0 Then
        MsgBox "Error."
    Else
        DB.AddNewOffset
        DoCalcOffset
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ScanOffset", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  DoCalcOffset                                          **
'**                                                                        **
'**  Description..:  This routine calculates a new offset.                 **
'**                                                                        **
'****************************************************************************
Private Sub DoCalcOffset()
    On Error GoTo ErrorHandler
    
    Dim OffsetNum As Integer, ScanFileName As String
    Dim RedOffsetFile As String, GreenOffsetFile As String, BlueOffsetFile As String
        
    OffsetNum = DB.rsOffsets("OffsetNum").Value
    
    ScanFileName = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Scan" & Trim(Str(OffsetNum)) & ".bmp"
    RedOffsetFile = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "r.frm"
    GreenOffsetFile = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "g.frm"
    BlueOffsetFile = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & "b.frm"
    
    '---- not sure if calc function uses the last offset file.... msg displays from DLL if files exist - does not show error, only reads "Add"
    FileSystemHandle.CopyFile OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum - 1)) & "r.frm", RedOffsetFile
    FileSystemHandle.CopyFile OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum - 1)) & "g.frm", GreenOffsetFile
    FileSystemHandle.CopyFile OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum - 1)) & "b.frm", BlueOffsetFile
    
    OffsetImage.PaintSizeMode = PAINTSIZEMODE_FITSIDES
    OffsetImage.ForceRepaint
    OffsetImage.Save ScanFileName, FILE_BMP, 24, 0, 0
    OffsetProgress.Visible = True
    OffsetProgress.Value = 0
    SetScanFile MakeCstring(ScanFileName)
    OffsetProgress.Value = 10
    
    SetOffsetFile MakeCstring(BlueOffsetFile)
    CalcOffset 0
    OffsetProgress.Value = 30
    SetOffsetFile MakeCstring(GreenOffsetFile)
    CalcOffset 1
    OffsetProgress.Value = 40
    SetOffsetFile MakeCstring(RedOffsetFile)
    CalcOffset 2
    OffsetProgress.Value = 50
    OffsetProgress.Value = 100
    
    MakeCalcFile
    
    DisplayCalcFile
    OffsetProgress.Visible = False
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":DoCalcOffset", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MakeCalcFile                                          **
'**                                                                        **
'**  Description..:  This routine calls the Offset validation program.     **
'**                                                                        **
'****************************************************************************
Private Function MakeCalcFile() As Boolean
    On Error GoTo ErrorHandler
    
    '---- Call FRMtoBMP.EXE to create caulated bitmap
    Dim CallCmd As String, OffsetNum As Integer, OffsetFileName As String, OffsetTxtFileName As String
    OffsetNum = DB.rsOffsets.Fields("OffsetNum").Value
    OffsetFileName = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Calc" & Trim(Str(OffsetNum)) & ".bmp"
    OffsetTxtFileName = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Calc" & Trim(Str(OffsetNum)) & ".txt"
    CallCmd = AppPath & CalcName & " " & OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum))
    
    If FileSystemHandle.FileExists(OffsetFileName) = True Then
        Kill OffsetFileName
    End If
    
    If FileSystemHandle.FileExists(OffsetTxtFileName) = True Then
        Kill OffsetTxtFileName
    End If
    
    
    If FileSystemHandle.FileExists(AppPath & CalcName) = True Then
        
        If Shell(CallCmd, vbMinimizedNoFocus) = 0 Then
            MsgBox "Error executing " & CallCmd
        Else
            AppLog InfoMsg, "MakeCalcFile," & CallCmd
            
            
            If FileSystemHandle.FileExists(OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & ".bmp") = True Then
                Name OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & ".bmp" As OffsetFileName
                Name OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset" & Trim(Str(OffsetNum)) & ".txt" As OffsetTxtFileName
            End If
        End If
        Sleep 1000
    Else
        MsgBox "Offset calculation program not found."
    End If
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":MakeCalcFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  DisplayCalcFile                                       **
'**                                                                        **
'**  Description..:  This routine displays the Calculated Offset File.     **
'**                                                                        **
'****************************************************************************
Private Function DisplayCalcFile() As Boolean
    On Error GoTo ErrorHandler
    '---- Display Calcultated Offset File
    Dim OffsetNum As Integer, OffsetFileName As String, OffsetCalcFileName As String
    OffsetNum = DB.rsOffsets.Fields("OffsetNum").Value
    OffsetFileName = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Scan" & Trim(Str(OffsetNum)) & ".bmp"
    OffsetCalcFileName = OffsetFilePath & Trim(PrinterName) & OffsetFolder & OffsetHistoryFolder & "\Offset_Calc" & Trim(Str(OffsetNum)) & ".bmp"
    If FileSystemHandle.FileExists(OffsetCalcFileName) = True Then
        CalcImage.Load OffsetCalcFileName, 24, 0, 1
    Else
        MakeCalcFile
        If FileSystemHandle.FileExists(OffsetCalcFileName) = True Then
            CalcImage.Load OffsetCalcFileName, 24, 0, 1
        Else
            CalcImage.Bitmap = 0
        End If
    End If
    If FileSystemHandle.FileExists(OffsetFileName) = True Then
        OffsetImage.Load OffsetFileName, 24, 0, 1
    Else
        OffsetImage.Bitmap = 0
    End If
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":DisplayCalcFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  OffsetGrid_AfterSelectChange                          **
'**                                                                        **
'**  Description..:  This routine updates the offset graphics.             **
'**                                                                        **
'****************************************************************************
Private Sub OffsetGrid_AfterSelectChange(ByVal SelectChange As UltraGrid.Constants_SelectChange)
    On Error GoTo ErrorHandler
    If SelectChange = ssSelectChangeRow Then
        DisplayCalcFile
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":OffsetGrid_AfterSelectChange", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
