VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form LutControlForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12135
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   30
      Top             =   11610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveToolBars.SSActiveToolBars MainToolBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   2
      ToolsCount      =   10
      Tools           =   "LutControlForm.frx":0000
      ToolBars        =   "LutControlForm.frx":749F
   End
   Begin SSSplitter.SSSplitter LutSplitter 
      Height          =   12135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   21405
      _Version        =   262144
      AutoSize        =   1
      SplitterBarJoinStyle=   0
      PaneTree        =   "LutControlForm.frx":762A
      Begin Threed.SSPanel SSPanel1 
         Height          =   3645
         Left            =   30
         TabIndex        =   4
         Top             =   8460
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   6429
         _Version        =   262144
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.Frame CurrentLUTFrame 
            Caption         =   "Current LUT Setting"
            Height          =   795
            Left            =   60
            TabIndex        =   7
            Top             =   540
            Width           =   4275
            Begin VB.Label LUTLabel 
               Caption         =   "LUTLabel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   4095
            End
         End
         Begin VB.TextBox PictoLutRef 
            Height          =   345
            Left            =   2460
            TabIndex        =   6
            Text            =   "2"
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Calculate Picto LUT Using Lut#"
            Height          =   375
            Left            =   150
            TabIndex        =   5
            Top             =   180
            Width           =   2475
         End
      End
      Begin UltraGrid.SSUltraGrid LutDensGrid 
         Height          =   4425
         Left            =   4785
         TabIndex        =   1
         Top             =   30
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   7805
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Densitometer Readings"
      End
      Begin ChartfxLibCtl.ChartFX LutChart 
         Height          =   7560
         Left            =   4785
         TabIndex        =   2
         Top             =   4545
         Width           =   9960
         _cx             =   17568
         _cy             =   13335
         Build           =   20
         TypeMask        =   109576195
         MarkerShape     =   0
         Axis(2).Min     =   0
         Axis(2).Max     =   100
         _Data_          =   "LutControlForm.frx":76BC
      End
      Begin UltraGrid.SSUltraGrid LutGrid 
         Height          =   8340
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   14711
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
         Override        =   "LutControlForm.frx":7714
         Caption         =   "LCD LUT Calibrations"
      End
   End
End
Attribute VB_Name = "LutControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2 Quality Control                                     **
'**                                                                        **
'** Module.....: LUTControl                                                **
'**                                                                        **
'** Description: This module provides basic LUT preview and editing.       **
'**                                                                        **
'** History....:                                                           **
'**    10/02/03 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Public LutFileHdr As String                             'This is the header for the binary LUT File
Public DensitometerLutPass As Integer                   'Densitometer strip # (3 strips read)
Private Const LutFolder As String = "\LUT"
Private Const LutHistoryFolder As String = "\History"

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes controls & validates folders.**
'**                                                                        **
'****************************************************************************
Public Function Setup()
    On Error GoTo ErrorHandler
    Dim TargetPath As String
    With FileSystemHandle
        TargetPath = LutFilePath & Trim(PrinterName)
        If .FolderExists(TargetPath) = False Then
            .CreateFolder TargetPath & "\"
        End If
        TargetPath = LutFilePath & Trim(PrinterName) & LutFolder
        If .FolderExists(TargetPath) = False Then                   'LUT Folder does not exist
            .CreateFolder TargetPath & "\"                          'Create it
            TargetPath = LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder
            .CreateFolder TargetPath & "\"                          'Create the LUT history folder
            ClearLUT True                                           'Copy the CLEAR LUT to current
        End If
    End With
   
    With LutGrid
        Set .DataSource = DB.rsLuts
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Activation = ssActivationActivateOnly
        .Bands(0).Columns(1).Header.Caption = "Lut#"
        .Bands(0).Columns(1).Width = 600
        .Bands(0).Columns(2).Activation = ssActivationActivateOnly
        .Bands(0).Columns(2).Header.Caption = "Date Created"
        .Bands(0).Columns(2).Width = 2000
        .Bands(0).Columns(3).Activation = ssActivationActivateOnly
        .Bands(0).Columns(3).Header.Caption = "Selected"
        .Bands(0).Columns(3).Width = 600
    End With
    
    DB.rsLuts.MoveLast
    DB.UpdateLutDensiValues
    
    UpdateLutDensiGrid
    ClearLUT False
    LoadLutTable DB.rsLuts("LutNum").Value
    DensitometerLutPass = 0
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateLutDensiGrid                                    **
'**                                                                        **
'**  Description..:  This routine updates densi grid on lut change.        **
'**                                                                        **
'****************************************************************************
Public Sub UpdateLutDensiGrid()
    On Error GoTo ErrorHandler
    With LutDensGrid
        Set .DataSource = DB.rsLutDensiValues
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Hidden = True
        .Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
        .Bands(0).Columns(2).Width = 650
        .Bands(0).Columns(3).Width = 900
        .Bands(0).Columns(4).Width = 900
        .Bands(0).Columns(5).Width = 900
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":UpdateLutDensiGrid", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

Private Sub Form_Activate()
    If DB.ApplyMullerSohnLUT = True Then
        Me.LUTLabel.Caption = "48-Step MuellerSOHN"
    Else
        Me.LUTLabel.Caption = "72-Step PictoGraphics"
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Resize                                           **
'**                                                                        **
'**  Description..:  This routine resizes controls on form resize.         **
'**                                                                        **
'****************************************************************************
Private Sub Form_Resize()
    '---- Simply resize controls on form resize
    On Error GoTo ErrorHandler
    If Me.Width - 100 > 0 Then
        LutSplitter.Width = Me.Width - 100
        LutDensGrid.Width = LutSplitter.Panes(1).Width - 100
        LutChart.Width = LutSplitter.Panes(2).Width - 100
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  LutButton_Click                                       **
'**                                                                        **
'**  Description..:  This routine handles the LUT Command Buttons.         **
'**                                                                        **
'****************************************************************************
Private Sub MainToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Dim LutNum As Integer, TargetName As String
    LutNum = DB.rsLuts("LutNum").Value
    Select Case Tool.ID
        Case "ID_Add"
            If MsgBox("Add a new LUT?", vbApplicationModal + vbQuestion + vbDefaultButton2 + vbYesNo, "WARNING!") = vbYes Then
                DB.AddNewLUT
                UpdateLutDensiGrid                                      'Update the LUT Densi grid display
                LoadLutTable DB.rsLuts("LutNum").Value
            End If
        Case "ID_Erase"
            If LutNum = 1 Then
                MsgBox "Cannot erase the clear LUT.", vbApplicationModal + vbInformation + vbOKOnly, "NOTICE"
            Else
                If MsgBox("Are you sure you wish to erase LUT #" & LutNum & "?", vbApplicationModal + vbQuestion + vbDefaultButton2 + vbYesNo, "WARNING!") = vbYes Then
                    '---- Kill the LUT Binary Files
                    If FileSystemHandle.FileExists(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "r.lut") = True Then
                        Kill LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "r.lut"
                    End If
                    If FileSystemHandle.FileExists(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "g.lut") = True Then
                        Kill LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "g.lut"
                    End If
                    If FileSystemHandle.FileExists(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "b.lut") = True Then
                        Kill LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "b.lut"
                    End If
                    '---- Delete the LUT Record (associated records will be deleted automatically by the database)
                    DB.rsLuts.Delete adAffectCurrent
                    DB.rsLuts.MoveFirst
                    DB.UpdateLutDensiValues                             'Update the LUT Select Statement
                    UpdateLutDensiGrid                                  'Update the LUT Densi grid display
                    LoadLutTable DB.rsLuts("LutNum").Value
                End If
            End If
        Case "ID_Select"
            If MsgBox("Make LUT #" & LutNum & " the current printer LUT?", vbApplicationModal + vbQuestion + vbDefaultButton2 + vbYesNo, "Select LUT") = vbYes Then
                DB.SetCurrentLUT
                '---- Copy LUT To Current
                If FileSystemHandle.FileExists(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "r.lut") = True Then
                    If FileSystemHandle.FileExists(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "g.lut") = True Then
                        If FileSystemHandle.FileExists(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "b.lut") = True Then
                            FileSystemHandle.CopyFile LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "r.lut", LutFilePath & Trim(PrinterName) & LutFolder & "\lutr.lut"
                            FileSystemHandle.CopyFile LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "g.lut", LutFilePath & Trim(PrinterName) & LutFolder & "\lutg.lut"
                            FileSystemHandle.CopyFile LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "b.lut", LutFilePath & Trim(PrinterName) & LutFolder & "\lutb.lut"
                        End If
                    End If
                End If
            End If
        Case "ID_Recalculate"
            If MsgBox("Recalculate LUT #" & LutNum & "?", vbApplicationModal + vbQuestion + vbDefaultButton2 + vbYesNo, "Recalculate LUT") = vbYes Then
                If DB.ApplyPictoLUT = False Then
                    SaveScannedLUT
                Else
                    CreatePictoLUT
                End If
            End If
        Case "ID_LoadCalculatedLUT"
            ImportCalculatedLUT
        Case "ID_LoadDensitometerValues"
            ImportDensitometerValues
        Case "ID_ExportCalculatedLUT"
            ExportCalculatedLUT
        Case "ID_ExportDensitometerValues"
            ExportDensiLUT
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":MainToolBar_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  LoadLutTable                                          **
'**                                                                        **
'**  Description..:  This routine loads binary LUT tables into spreadsheet.**
'**                                                                        **
'****************************************************************************
Public Function LoadLutTable(LutNum As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim RedFileHandle As Scripting.TextStream, GrnFileHandle As Scripting.TextStream, BluFileHandle As Scripting.TextStream    'Pointer to LUT File
    Dim RowNum As Long, value1 As Byte, value2 As Byte, iValue As Long              ', LutNum As Integer
    
    '---- Old Code ---> LutNum = DB.rsLuts("LutNum").Value
    
    If FileSystemHandle.FileExists(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "r.lut") = False Then
        
        '--- RDR debugging muellersohn lut - following could be the problem
        'SaveScannedLUT
        
    End If
    
    With LutChart
        .MenuBar = True
        .ToolBar = True
        .ShowTips = True
        .Gallery = CURVE
        .ClearData CD_DATA
        .MarkerShape = MK_NONE
        .LineWidth = 1
        .MaxValues = 256
        .Axis(AXIS_Y).Grid = True
        .Axis(AXIS_Y).AutoScale = False
        .Axis(AXIS_Y).Decimals = 0
        .Axis(AXIS_Y).Min = 0
        .Axis(AXIS_Y).Max = 4096
        .Axis(AXIS_Y).STEP = 256
        .Axis(AXIS_X).Grid = True
        .Axis(AXIS_X).AutoScale = False
        .Axis(AXIS_X).Min = 0
        .Axis(AXIS_X).Max = 256
        .Axis(AXIS_X).STEP = 8
        .Axis(AXIS_X).Decimals = 0
        .OpenDataEx COD_VALUES, 4, 256
        
        Set RedFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "r.lut", ForReading, False, TristateFalse)
        Set GrnFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "g.lut", ForReading, False, TristateFalse)
        Set BluFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & Trim(Str(LutNum)) & "b.lut", ForReading, False, TristateFalse)
        LutFileHdr = RedFileHandle.Read(16)
        LutFileHdr = GrnFileHandle.Read(16)
        LutFileHdr = BluFileHandle.Read(16)
        For iValue = 0 To .MaxValues - 1
            .ValueEx(0, iValue) = CInt(4096 - (iValue * (4096 / 256)))
            value1 = Asc(RedFileHandle.Read(1))
            value2 = Asc(RedFileHandle.Read(1))
            .ValueEx(1, iValue) = CInt(value1 + (value2 * 256))
            value1 = Asc(GrnFileHandle.Read(1))
            value2 = Asc(GrnFileHandle.Read(1))
            .ValueEx(2, iValue) = CInt(value1 + (value2 * 256))
            value1 = Asc(BluFileHandle.Read(1))
            value2 = Asc(BluFileHandle.Read(1))
            .ValueEx(3, iValue) = CInt(value1 + (value2 * 256))
        Next
        RedFileHandle.Close
        GrnFileHandle.Close
        BluFileHandle.Close
        .CloseData COD_VALUES
        .OpenDataEx COD_COLORS, 4, 256
        .Series(0).color = vbBlack
        .Series(1).color = vbRed
        .Series(2).color = vbGreen
        .Series(3).color = vbBlue
        .CloseData COD_COLORS
        .DataEditor = True
        .DataEditorObj.Docked = TGFP_LEFT
        .DataEditorObj.Moveable = False
        .DataEditorObj.Sizeable = BAS_NORESIZE
    End With
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":LoadLutTable", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ClearLUT                                              **
'**                                                                        **
'**  Description..:  This routine copies the Clear LUT files over current. **
'**                                                                        **
'****************************************************************************
Public Function ClearLUT(ClearFlag As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim TargetPath As String
    TargetPath = LutFilePath & Trim(PrinterName) & LutFolder
    With FileSystemHandle
        If .FileExists(TargetPath & "\lutr.lut") = False Or ClearFlag = True Then
            '--- Copy CLEAR LUT to folder
            .CopyFile SettingsFolder & "LUT\ClearLutr.lut", TargetPath & "\Lutr.lut", True
            .CopyFile SettingsFolder & "LUT\ClearLutg.lut", TargetPath & "\Lutg.lut", True
            .CopyFile SettingsFolder & "LUT\ClearLutb.lut", TargetPath & "\Lutb.lut", True
            '--- Copy CLEAR LUT to history folder
            .CopyFile SettingsFolder & "LUT\ClearLutr.lut", TargetPath & LutHistoryFolder & "\Lut1r.lut", True
            .CopyFile SettingsFolder & "LUT\ClearLutg.lut", TargetPath & LutHistoryFolder & "\Lut1g.lut", True
            .CopyFile SettingsFolder & "LUT\ClearLutb.lut", TargetPath & LutHistoryFolder & "\Lut1b.lut", True
        End If
    End With
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ClearLUT", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SaveScannedLUT                                        **
'**                                                                        **
'**  Description..:  This routine saves densi data to binary LUT.          **
'**                                                                        **
'****************************************************************************
Public Sub SaveScannedLUT()
    On Error GoTo ErrorHandler

    Dim TargetPath As String, BlockNum As Integer, LutNum As Integer
    LutNum = DB.rsLuts("LutNum").Value
    TargetPath = LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder
    
    AppLog DebugMsg, "LutButton_Click,Calculating LUT - Copying Clear LUT Files..."
    If FileSystemHandle.FileExists(SettingsFolder & "LUT\ClearLutr.lut") Then
        FileSystemHandle.CopyFile SettingsFolder & "LUT\ClearLutr.lut", TargetPath & "\Lut" & Trim(Str(LutNum)) & "r.lut", True
    Else
        AppLog ErrorMsg, "LutButton_Click,Clear LUT File [" & SettingsFolder & "LUT\ClearLutr.lut" & "] not found."
        Exit Sub
    End If
    
    If FileSystemHandle.FileExists(SettingsFolder & "LUT\ClearLutg.lut") Then
        FileSystemHandle.CopyFile SettingsFolder & "LUT\ClearLutg.lut", TargetPath & "\Lut" & Trim(Str(LutNum)) & "g.lut", True
    Else
        AppLog ErrorMsg, "LutButton_Click,Clear LUT File [" & SettingsFolder & "LUT\ClearLutg.lut" & "] not found."
        Exit Sub
    End If
    
    If FileSystemHandle.FileExists(SettingsFolder & "LUT\ClearLutb.lut") Then
        FileSystemHandle.CopyFile SettingsFolder & "LUT\ClearLutb.lut", TargetPath & "\Lut" & Trim(Str(LutNum)) & "b.lut", True
    Else
        AppLog ErrorMsg, "LutButton_Click,Clear LUT File [" & SettingsFolder & "LUT\ClearLutb.lut" & "] not found."
        Exit Sub
    End If
    
    DoEvents
    
    With DB.rsLutDensiValues
        AppLog DebugMsg, "LutButton_Click,Calculating Red LUT to " & TargetPath & "\Lut" & Trim(Str(LutNum)) & "r.lut"
        For BlockNum = 1 To 48
            .MoveFirst
            .Find "BlockNum=" & BlockNum
            If Not .EOF Then
                
                SetDensiValue BlockNum - 1, CLng(.Fields("DensRed").Value)
                
            Else
                MsgBox "Error finding block# " & BlockNum
            End If
        Next
        SetLutFile TargetPath & "\Lut" & Trim(Str(LutNum)) & "r.lut"
        DoEvents
        CalcLut         'DLL
        AppLog DebugMsg, "LutButton_Click,Calculating Green LUT to " & TargetPath & "\Lut" & Trim(Str(LutNum)) & "g.lut"
        For BlockNum = 1 To 48
            .MoveFirst
            .Find "BlockNum=" & BlockNum
            If Not .EOF Then
                SetDensiValue BlockNum - 1, CLng(.Fields("DensGreen").Value)
            Else
                MsgBox "Error finding block# " & BlockNum
            End If
        Next
        SetLutFile TargetPath & "\Lut" & Trim(Str(LutNum)) & "g.lut"
        DoEvents
        CalcLut         'DLL
        AppLog DebugMsg, "LutButton_Click,Calculating Blue LUT to " & TargetPath & "\Lut" & Trim(Str(LutNum)) & "b.lut"
        For BlockNum = 1 To 48
            .MoveFirst
            .Find "BlockNum=" & BlockNum
            If Not .EOF Then
                SetDensiValue BlockNum - 1, CLng(.Fields("DensBlue").Value)
            Else
                MsgBox "Error finding block# " & BlockNum
            End If
        Next
        SetLutFile TargetPath & "\Lut" & Trim(Str(LutNum)) & "b.lut"
        DoEvents
        CalcLut         'DLL
    End With
    
    LoadLutTable DB.rsLuts("LutNum").Value
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SaveScannedLUT", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SaveLutTable                                          **
'**                                                                        **
'**  Description..:  This routine saves binary LUT tables to files.        **
'**                                                                        **
'****************************************************************************
Function SaveLutTable() As Boolean
    On Error GoTo ErrorHandler
    Dim RedFileHandle As Scripting.TextStream, GrnFileHandle As Scripting.TextStream, BluFileHandle As Scripting.TextStream  'Pointer to LUT File
    Dim iValue As Long, HexData As String
    With LutChart
        Set RedFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & "\Newlutr.lut", ForWriting, False, TristateFalse)
        Set GrnFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & "\Newlutg.lut", ForWriting, False, TristateFalse)
        Set BluFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & "\Newlutb.lut", ForWriting, False, TristateFalse)
        RedFileHandle.Write Left(LutFileHdr, 16)
        GrnFileHandle.Write Left(LutFileHdr, 16)
        BluFileHandle.Write Left(LutFileHdr, 16)
        For iValue = 0 To .MaxValues - 1
            HexData = MyHex(.ValueEx(1, iValue), 4)
            RedFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
            HexData = MyHex(.ValueEx(2, iValue), 4)
            GrnFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
            HexData = MyHex(.ValueEx(3, iValue), 4)
            BluFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
        Next
        RedFileHandle.Close
        GrnFileHandle.Close
        BluFileHandle.Close
    End With
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SaveLutTable", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  LutGrid_AfterSelectChange                             **
'**                                                                        **
'**  Description..:  This routine updates LUT display on grid row change   **
'**                                                                        **
'****************************************************************************
Private Sub LutGrid_AfterSelectChange(ByVal SelectChange As UltraGrid.Constants_SelectChange)
    On Error GoTo ErrorHandler
    If SelectChange = ssSelectChangeRow Then
        DB.UpdateLutDensiValues                             'Update the LUT Select Statement
        UpdateLutDensiGrid                                  'Update the LUT Densi grid display
        LoadLutTable DB.rsLuts("LutNum").Value
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":LutGrid_AfterSelectChange", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  LutDensGrid_Validate                                  **
'**                                                                        **
'**  Description..:  This routine verifies cancellation of LUT reading.    **
'**                                                                        **
'****************************************************************************
Private Sub LutDensGrid_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    '---- Verify we are scanning a new LUT
    If DensitometerLutPass <> 0 Then
        If MsgBox("Cancel reading current LUT?", vbApplicationModal + vbDefaultButton2 + vbYesNo + vbQuestion, "Are you sure?") = vbNo Then
            Cancel = True
        Else
            DensitometerLutPass = 0
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":LutDensGrid_Validate", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ImportCalculatedLUT                                   **
'**                                                                        **
'**  Description..:  This routine loads delimited text into binary LUT.    **
'**                                                                        **
'****************************************************************************
Private Sub ImportDensitometerValues()
    On Error GoTo ErrorHandler
    Dim RedFileHandle As Scripting.TextStream, GrnFileHandle As Scripting.TextStream, BluFileHandle As Scripting.TextStream  'Pointer to LUT File
    Dim fh As Scripting.TextStream, Buf As String, RGB() As String
    Dim LutRed(4096) As Currency, LutGreen(4096) As Currency, LutBlue(4096) As Currency
    Dim iValue As Long, HexData As String, ImportFile As String
    
    '---- Display File Open Diaglog
    With CommonDialog
        .DefaultExt = ".txt|Comma-Separated-Value File (CSV)"
        .DialogTitle = "Select LUT Text File"
        .ShowOpen
        ImportFile = .FileName
    End With
    
    If ImportFile = "" Then
        Exit Sub
    End If
    If MsgBox("Import " & ImportFile & " into new LUT?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
        Set fh = FileSystemHandle.OpenTextFile(ImportFile, ForReading, False, TristateFalse)
        iValue = 0
        DB.AddNewLUT                                            'Add new LUT record to the database
        With DB.rsLutDensiValues
            If .RecordCount > 0 Then .MoveFirst
            Do While Not fh.AtEndOfStream
                Buf = Trim(fh.ReadLine)
                'Debug.Print Buf
                RGB = Split(Buf, " ")
                Debug.Print RGB(0), RGB(1), RGB(2)
                If UBound(RGB) = 2 Then
                    .Fields("DensRed").Value = RGB(0) * 100
                    .Fields("DensGreen").Value = RGB(1) * 100
                    .Fields("DensBlue").Value = RGB(2) * 100
                    .UpdateBatch adAffectCurrent
                    iValue = iValue + 1
                    If iValue < MaxLutBlock Then
                        .MoveNext
                    End If
                Else
                    AppLog ErrorMsg, "ImportDensitometerValues,Invalid LUT import row [" & Buf & "]"
                End If
            Loop
        End With
        fh.Close                                                'Close the import file
        UpdateLutDensiGrid                                      'Update the LUT Densi grid display
        If DB.ApplyPictoLUT = False Then
            If iValue = MaxLutBlock Then
                SaveScannedLUT
            End If
        Else
            If iValue = MaxLutBlock Then
                CreatePictoLUT
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ImportDensitometerValues", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ImportCalculatedLUT                                   **
'**                                                                        **
'**  Description..:  This routine loads delimited text into binary LUT.    **
'**                                                                        **
'****************************************************************************
Private Sub ImportCalculatedLUT()
    On Error GoTo ErrorHandler
    Dim RedFileHandle As Scripting.TextStream, GrnFileHandle As Scripting.TextStream, BluFileHandle As Scripting.TextStream  'Pointer to LUT File
    Dim fh As Scripting.TextStream, Buf As String, RGB() As String
    Dim LutRed(4096) As Currency, LutGreen(4096) As Currency, LutBlue(4096) As Currency
    Dim iValue As Long, HexData As String, ImportFile As String
    
    '---- Display File Open Diaglog
    With CommonDialog
        .DefaultExt = ".txt|Comma-Separated-Value File (CSV)"
        .DialogTitle = "Select LUT Text File"
        .ShowOpen
        ImportFile = .FileName
    End With
    
    If ImportFile = "" Then
        Exit Sub
    End If
    If MsgBox("Import " & ImportFile & " into new LUT?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
        Set fh = FileSystemHandle.OpenTextFile(ImportFile, ForReading, False, TristateFalse)
        iValue = 0
        
        Do While Not fh.AtEndOfStream
            Buf = fh.ReadLine
            RGB = Split(Buf, ",")
            If UBound(RGB) = 2 Then
                LutRed(iValue) = RGB(0)
                LutGreen(iValue) = RGB(1)
                LutBlue(iValue) = RGB(2)
                iValue = iValue + 1
            Else
                AppLog ErrorMsg, "ImportCalculatedLUT,Invalid LUT import row [" & Buf & "]"
            End If
        Loop
        fh.Close                                                    'Close the import file
        If iValue = 256 Then
            DB.AddNewLUT                                            'Add new LUT record to the database
            UpdateLutDensiGrid                                      'Update the LUT Densi grid display
            Set RedFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & DB.rsLuts.Fields("LutNum").Value & "r.lut", ForWriting, True, TristateFalse)
            Set GrnFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & DB.rsLuts.Fields("LutNum").Value & "g.lut", ForWriting, True, TristateFalse)
            Set BluFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & DB.rsLuts.Fields("LutNum").Value & "b.lut", ForWriting, True, TristateFalse)
            RedFileHandle.Write Left(LutFileHdr, 16)
            GrnFileHandle.Write Left(LutFileHdr, 16)
            BluFileHandle.Write Left(LutFileHdr, 16)
            For iValue = 0 To 256
                HexData = MyHex(LutRed(iValue), 4)
                RedFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
                HexData = MyHex(LutGreen(iValue), 4)
                GrnFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
                HexData = MyHex(LutBlue(iValue), 4)
                BluFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
            Next
            RedFileHandle.Close
            GrnFileHandle.Close
            BluFileHandle.Close
            
            LoadLutTable DB.rsLuts("LutNum").Value
        Else
            AppLog ErrorMsg, "ImportCalculatedLUT,Invalid number of imported LUT rows [" & iValue & "]"
        End If
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ImportCalculatedLUT", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ExportCalculatedLUT                                   **
'**                                                                        **
'**  Description..:  This routine saves binary LUT to delimited text file. **
'**                                                                        **
'****************************************************************************
Private Sub ExportCalculatedLUT()
    On Error GoTo ErrorHandler
    Dim fh As Scripting.TextStream
    Dim iValue As Long, ExportFile As String
    
    '---- Display File Save Diaglog
    With CommonDialog
        .FileName = ""
        .DefaultExt = "txt"
        .DialogTitle = "Select LUT Text File"
        .ShowSave
        ExportFile = .FileName
    End With
    If ExportFile = "" Then
        Exit Sub
    End If
    If MsgBox("Export to " & ExportFile & "?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
        Set fh = FileSystemHandle.OpenTextFile(ExportFile, ForWriting, True, TristateFalse)
        With LutChart
            For iValue = 0 To .MaxValues - 1
                fh.WriteLine .ValueEx(1, iValue) & "," & .ValueEx(2, iValue) & "," & .ValueEx(3, iValue)
            Next
        End With
        fh.Close
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ExportCalculatedLUT", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ExportDensitometerLUT                                 **
'**                                                                        **
'**  Description..:  This routine saves Densi LUT to delimited text file.  **
'**                                                                        **
'****************************************************************************
Private Sub ExportDensiLUT()
    On Error GoTo ErrorHandler
    Dim fh As Scripting.TextStream
    Dim iValue As Long, ExportFile As String
    
    '---- Display File Save Diaglog
    With CommonDialog
        .FileName = ""
        .DefaultExt = "txt"
        .DialogTitle = "Select LUT Text File"
        .ShowSave
        ExportFile = .FileName
    End With
    If ExportFile = "" Then
        Exit Sub
    End If
    If MsgBox("Export to " & ExportFile & "?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
        Set fh = FileSystemHandle.OpenTextFile(ExportFile, ForWriting, True, TristateFalse)
        With DB.rsLutDensiValues
            .MoveFirst
            Do While Not .EOF
                fh.WriteLine .Fields("DensRed").Value & "," & .Fields("DensGreen").Value & "," & .Fields("DensBlue").Value
                .MoveNext
            Loop
        End With
        fh.Close
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ExportCalculatedLUT", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CreatePictoLUT                                        **
'**                                                                        **
'**  Description..:  This routine creates a pictographics LUT.             **
'**                                                                        **
'****************************************************************************
Private Sub CreatePictoLUT()

    On Error GoTo ErrorHandler
    Dim fh As Scripting.TextStream
    Dim iValue As Long, ExportFile As String

    '----- Export the Target Density File
    ExportFile = App.Path & "\PictoIN.den"
    Set fh = FileSystemHandle.OpenTextFile(ExportFile, ForWriting, True, TristateFalse)
    fh.WriteLine "PictoIN.den"
    fh.WriteLine "0"
    fh.WriteLine "0"
    fh.WriteLine "255"
    fh.WriteLine "4500"
    fh.WriteLine "32"
    With DB.rsEmulsionData
        .MoveFirst
        Do While Not .EOF
            fh.WriteLine Format(.Fields("DensityNumber").Value, "#0") & " " & _
                Format(.Fields("RedValue").Value, "##0") & " " & _
                Format(.Fields("GreenValue").Value, "##0") & " " & _
                Format(.Fields("BlueValue").Value, "##0") & " " & _
                Format(.Fields("RedDensity").Value, "0.000000") & " " & _
                Format(.Fields("GreenDensity").Value, "0.000000") & " " & _
                Format(.Fields("BlueDensity").Value, "0.000000")
            .MoveNext
        Loop
    End With
    fh.Close
    
    '----- Export the Input Map File - this is the last good LUT!!  Need a way to determine which one to use.
    LoadLutTable CInt(Me.PictoLutRef.Text)
    
    ExportFile = App.Path & "\PictoIN.map"
    Set fh = FileSystemHandle.OpenTextFile(ExportFile, ForWriting, True, TristateFalse)
    fh.WriteLine "Version = 1"
    fh.WriteLine "Type = LUT"
    fh.WriteLine "Rows = 256"
    fh.WriteLine "Cols = 3"
    fh.WriteLine "UseXValues = 0"
    With LutChart
        For iValue = 0 To .MaxValues - 1
            '---- Per Gerry Note:  fh.WriteLine Format(.ValueEx(1, iValue) / 10000, "0.0000") & ", " & Format(.ValueEx(2, iValue) / 10000, "0.0000") & ", " & Format(.ValueEx(3, iValue) / 10000, "0.0000")
            fh.WriteLine Format(.ValueEx(1, iValue) / 4095, "0.0000") & ", " & Format(.ValueEx(2, iValue) / 4095, "0.0000") & ", " & Format(.ValueEx(3, iValue) / 4095, "0.0000")
        Next
    End With
    fh.Close
    
    '----- Exposrt the Measurement Data File
    ExportFile = App.Path & "\PictoIN.mes"
    Set fh = FileSystemHandle.OpenTextFile(ExportFile, ForWriting, True, TristateFalse)
    fh.WriteLine "Gray Balance Response"
    With DB.rsLutDensiValues
        .MoveFirst
        Do While Not .EOF
            fh.WriteLine "#" & Trim(Str(.Fields("BlockNum").Value)) & " " & Format(.Fields("DensRed").Value / 100, "0.0000") & " " & Format(.Fields("DensGreen").Value / 100, "0.0000") & " " & Format(.Fields("DensBlue").Value / 100, "0.0000")
            .MoveNext
        Loop
    End With
    fh.Close
    
    '----- Call the RC Pictographics Wrapper
    Shell App.Path & "\PictoDVP2.exe " & Trim(App.Path) & "\"
    
    '----- Import the Output Map File
    Sleep 1000
    ImportPictoLUT Trim(App.Path) & "\PictoOUT.map"
    
    MsgBox "Completed"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":CreatePictoLUT", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ImportPictoLUT                                        **
'**                                                                        **
'**  Description..:  This routine imports the pctographics LUT output file.**
'**                                                                        **
'****************************************************************************
Private Sub ImportPictoLUT(ImportFile As String)
    On Error GoTo ErrorHandler
    
    Dim RedFileHandle As Scripting.TextStream, GrnFileHandle As Scripting.TextStream, BluFileHandle As Scripting.TextStream  'Pointer to LUT File
    Dim fh As Scripting.TextStream, Buf As String, RGB() As String
    Dim LutRed(4096) As Currency, LutGreen(4096) As Currency, LutBlue(4096) As Currency
    Dim iValue As Long, HexData As String       ', ImportFile As String
        
    Set fh = FileSystemHandle.OpenTextFile(ImportFile, ForReading, False, TristateFalse)
    Buf = fh.ReadLine           'version
    Buf = fh.ReadLine           'type
    Buf = fh.ReadLine           'rows
    Buf = fh.ReadLine           'cols
    Buf = fh.ReadLine           'xyvalues
    
    iValue = 0
    
    Do While Not fh.AtEndOfStream
        Buf = fh.ReadLine
        RGB = Split(Buf, ",")
        If UBound(RGB) = 2 Then
            LutRed(iValue) = RGB(0) * 4095
            LutGreen(iValue) = RGB(1) * 4095
            LutBlue(iValue) = RGB(2) * 4095
            iValue = iValue + 1
        Else
            AppLog ErrorMsg, "ImportPictoLUT,Invalid LUT import row [" & Buf & "]"
        End If
    Loop
    fh.Close                                                    'Close the import file
    If iValue = 256 Then
        Set RedFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & DB.rsLuts.Fields("LutNum").Value & "r.lut", ForWriting, True, TristateFalse)
        Set GrnFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & DB.rsLuts.Fields("LutNum").Value & "g.lut", ForWriting, True, TristateFalse)
        Set BluFileHandle = FileSystemHandle.OpenTextFile(LutFilePath & Trim(PrinterName) & LutFolder & LutHistoryFolder & "\lut" & DB.rsLuts.Fields("LutNum").Value & "b.lut", ForWriting, True, TristateFalse)
        RedFileHandle.Write Left(LutFileHdr, 16)
        GrnFileHandle.Write Left(LutFileHdr, 16)
        BluFileHandle.Write Left(LutFileHdr, 16)
        For iValue = 0 To 256
            HexData = MyHex(LutRed(iValue), 4)
            RedFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
            HexData = MyHex(LutGreen(iValue), 4)
            GrnFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
            HexData = MyHex(LutBlue(iValue), 4)
            BluFileHandle.Write Chr(Val("&H" & Mid(HexData, 3, 2))) & Chr(Val("&H" & Mid(HexData, 1, 2)))
        Next
        RedFileHandle.Close
        GrnFileHandle.Close
        BluFileHandle.Close
        
        LoadLutTable DB.rsLuts("LutNum").Value
    Else
        AppLog ErrorMsg, "ImportCalculatedLUT,Invalid number of imported LUT rows [" & iValue & "]"
    End If
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ImportPictoLUT", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

