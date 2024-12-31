VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form HotFolderSettingsForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   13665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13665
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars HotFolderToolBars 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   3
      Tools           =   "HotFolderSettingsForm.frx":0000
      ToolBars        =   "HotFolderSettingsForm.frx":264D
   End
   Begin UltraGrid.SSUltraGrid HotFolderSpread 
      Height          =   12825
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   22622
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   71630852
      BorderStyle     =   6
      BorderStyleCaption=   6
      ValueLists      =   "HotFolderSettingsForm.frx":2714
      Bands           =   "HotFolderSettingsForm.frx":27B9
      Appearance      =   "HotFolderSettingsForm.frx":2A6E
      Caption         =   "HotFolderSpread"
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   60
      Top             =   13140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "HotFolderSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP2                                                      **
'**                                                                        **
'** Module.....: HotFolderSettingsForm                                     **
'**                                                                        **
'** Description: This module provides hot folder configuration.            **
'**                                                                        **
'** History....:                                                           **
'**    10/02/03 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Public Sub Setup()
    On Error GoTo ErrorHandler
    With HotFolderSpread
        '--- HotFolderType is a Value-List
        .ValueLists.Clear
        .ValueLists.Add "KeyList"
        .ValueLists.Item("KeyList").DisplayStyle = ssValueListDisplayStyleDisplayText
        .ValueLists.Item("KeyList").ValueListItems.Add 1, "Render"
        .ValueLists.Item("KeyList").ValueListItems.Add 2, "CF (Single)"
        .ValueLists.Item("KeyList").ValueListItems.Add 3, "CF (Multi)"
        Set .DataSource = DB.rsHotFolders
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Header.Caption = "Hot Folder"
        .Bands(0).Columns(1).Width = 7000
        .Bands(0).Columns(2).Header.Caption = "Type"
        .Bands(0).Columns(2).Width = 2000
        .Bands(0).Columns(2).Style = ssStyleDropDownList
        .Bands(0).Columns(2).ValueList = "KeyList"
        .Bands(0).Columns(3).Header.Caption = "Enabled"
        .Bands(0).Columns(3).Width = 1200
        .Bands(0).Columns(4).Activation = ssActivationActivateOnly
        .Bands(0).Columns(4).Header.Caption = "Status"
        .Bands(0).Columns(4).Width = 1200
        .Bands(0).Columns(5).Activation = ssActivationActivateOnly
        .Bands(0).Columns(5).Header.Caption = "# Files"
        .Bands(0).Columns(5).Width = 1200
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  FolderToolBar_ButtonClick                             **
'**                                                                        **
'**  Description..:  Handle hot folder tool bar.                           **
'**                                                                        **
'****************************************************************************
Private Sub HotFolderToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Dim FolderName As String
    With DB.rsHotFolders
        Select Case Tool.ID
            Case "ID_Add"
                CommonDialog.Flags = cdlOFNPathMustExist Or cdlOFNNoChangeDir Or cdlOFNExplorer Or cdlOFNNoValidate
                CommonDialog.FileName = "*.*"
                CommonDialog.CancelError = False
                CommonDialog.Action = 1
                FolderName = CommonDialog.FileName
                If FolderName <> "*.*" Then
                    FolderName = Left(FolderName, InStrRev(FolderName, "\", Len(FolderName), vbTextCompare) - 1)
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "HotFolderPath='" & Trim(UCase(FolderName)) & "'", 0, adSearchForward, 0
                    End If
                    If .EOF Then
                        .AddNew
                        .Fields("PrinterName").Value = PrinterName
                        .Fields("HotFolderPath").Value = Trim(UCase(FolderName))
                        .Fields("FolderType").Value = 1
                        .Fields("FolderEnabled").Value = 1
                        .Fields("Status").Value = "Ok"
                        .Fields("NumberOfFiles").Value = 0
                        .UpdateBatch adAffectCurrent
                    End If
                End If
            Case "ID_Erase"
                If Not .BOF And Not .EOF Then
                    FolderName = .Fields("HotFolderPath").Value
                    If MsgBox("Delete Hot Folder [" & FolderName & "] from List?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") = vbYes Then
                        AppLog InfoMsg, "Removed " & Trim(FolderName) & " from hot folder list."
                        .Delete adAffectCurrent
                        DoEvents
                        If .RecordCount > 0 Then .MoveFirst
                        HotFolderSpread.Refresh ssRefetchAndFireInitializeRow
                    End If
                End If
            Case "ID_Refresh"
                MsgBox "Coming soon..."
        End Select
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":FolderToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
