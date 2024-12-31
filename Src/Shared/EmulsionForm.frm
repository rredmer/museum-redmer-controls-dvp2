VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form EmulsionForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12135
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter LutSplitter 
      Height          =   12135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   21405
      _Version        =   262144
      AutoSize        =   1
      SplitterBarJoinStyle=   0
      PaneTree        =   "EmulsionForm.frx":0000
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
         ToolBarsCount   =   1
         ToolsCount      =   4
         Tools           =   "EmulsionForm.frx":0052
         ToolBars        =   "EmulsionForm.frx":336E
      End
      Begin UltraGrid.SSUltraGrid EmulsionGrid 
         Height          =   12075
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   21299
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
         Override        =   "EmulsionForm.frx":3453
         Caption         =   "Paper Emulsions"
      End
      Begin UltraGrid.SSUltraGrid EmulsionDataGrid 
         Height          =   12075
         Left            =   4785
         TabIndex        =   2
         Top             =   30
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   21299
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
         Caption         =   "Emulsion Target Densities"
      End
   End
End
Attribute VB_Name = "EmulsionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2 Quality Control                                     **
'**                                                                        **
'** Module.....: EmulsionForm                                              **
'**                                                                        **
'** Description: This module provides basic Emulsion preview and editing.  **
'**                                                                        **
'** History....:                                                           **
'**    10/02/03 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Public EmulsionFileHdr As String                             'This is the header for the binary Emulsion File

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes controls & validates folders.**
'**                                                                        **
'****************************************************************************
Public Function Setup()
    On Error GoTo ErrorHandler
    
    With EmulsionGrid
        Set .DataSource = DB.rsEmulsions
        .Refresh ssRefetchAndFireInitializeRow
    End With
    
    DB.rsEmulsions.MoveLast
    
    DB.GetEmulsionDataRecordset
    With EmulsionDataGrid
        Set .DataSource = DB.rsEmulsionData
    End With
    
    
    UpdateEmulsionDataGrid
    
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateEmulsionData                                    **
'**                                                                        **
'**  Description..:  This routine updates data grid on Emulsion change.    **
'**                                                                        **
'****************************************************************************
Public Sub UpdateEmulsionDataGrid()
    On Error GoTo ErrorHandler
    With EmulsionDataGrid
        .Refresh ssRefetchAndFireInitializeRow
        
        .Bands(0).Columns(0).Hidden = True
        '.Bands(0).Columns(1).Hidden = True
        '.Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
        '.Bands(0).Columns(2).Width = 650
        '.Bands(0).Columns(3).Width = 900
        '.Bands(0).Columns(4).Width = 900
        '.Bands(0).Columns(5).Width = 900
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":UpdateEmulsionDensiGrid", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

Private Sub EmulsionGrid_AfterSelectChange(ByVal SelectChange As UltraGrid.Constants_SelectChange)
    On Error GoTo ErrorHandler
    If SelectChange = ssSelectChangeRow Then
        DB.GetEmulsionDataRecordset
        UpdateEmulsionDataGrid
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":LutGrid_AfterSelectChange", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
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
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Private Sub MainToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Import"
            ImportTargetFile
    End Select
End Sub



Private Sub ImportTargetFile()
    On Error GoTo ErrorHandler
    Dim fh As Scripting.TextStream, Buf As String, EDATA() As String
    Dim iValue As Long, HexData As String       ', ImportFile As String
    Dim ImportFile As String
       
    '---- Display File Open Diaglog
    With CommonDialog
        .DefaultExt = ".tgt|Target Density File (TGT)"
        .DialogTitle = "Select Target Density File Text File"
        .ShowOpen
        ImportFile = .FileName
    End With
    
    If ImportFile = "" Then
        Exit Sub
    End If
    If MsgBox("Import " & ImportFile & " into new Emulsion Table?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
    
    
       ' With DB.rsEmulsions
       '     .AddNew
       '     .Fields("EmulsionNumber").Value = 1
       '     .Fields("EmulsionBrand").Value = "New"
       '     .Fields("EmulsionName").Value = "New"
       '     .Fields("EmulsionCode").Value = "New"
       '     .Fields("DateModified").Value = Now
       '     .UpdateBatch adAffectAllChapters
       ' End With
        
    
        Set fh = FileSystemHandle.OpenTextFile(ImportFile, ForReading, False, TristateFalse)
        Buf = fh.ReadLine           'version
        Buf = fh.ReadLine           'unknown
        Buf = fh.ReadLine           'range low
        Buf = fh.ReadLine           'range high
        Buf = fh.ReadLine           'scale
        Buf = fh.ReadLine           'num steps
        
        iValue = 0
        DB.rsEmulsionData.MoveFirst
        
        Do While Not fh.AtEndOfStream
            Buf = fh.ReadLine
            EDATA = Split(Buf, " ")
            If UBound(EDATA) = 6 Then
                With DB.rsEmulsionData
                    '.AddNew
                    .Fields("EmulsionNumber").Value = 1
                    .Fields("DensityNumber").Value = iValue + 1
                    .Fields("RedValue").Value = Val(EDATA(1))
                    .Fields("GreenValue").Value = Val(EDATA(2))
                    .Fields("BlueValue").Value = Val(EDATA(3))
                    .Fields("RedDensity").Value = Val(EDATA(4))
                    .Fields("GreenDensity").Value = Val(EDATA(5))
                    .Fields("BlueDensity").Value = Val(EDATA(6))
                    .UpdateBatch adAffectCurrent
                    .MoveNext
                End With
                iValue = iValue + 1
            Else
                AppLog ErrorMsg, "Import TargetFile,Invalid row [" & Buf & "]"
            End If
        Loop
        fh.Close                                                    'Close the import file
    End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ImportTargetFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
