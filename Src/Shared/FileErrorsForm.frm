VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form FileErrorsForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17295
   LinkTopic       =   "Form1"
   ScaleHeight     =   12390
   ScaleWidth      =   17295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars FileErrorsToolBars 
      Left            =   0
      Top             =   -30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "FileErrorsForm.frx":0000
      ToolBars        =   "FileErrorsForm.frx":0CE7
   End
   Begin UltraGrid.SSUltraGrid FileErrorGrid 
      Height          =   12045
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   17205
      _ExtentX        =   30348
      _ExtentY        =   21246
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   72351748
      BorderStyle     =   5
      Override        =   "FileErrorsForm.frx":0D67
      Appearance      =   "FileErrorsForm.frx":0DE5
      Caption         =   "FileErrorGrid"
   End
End
Attribute VB_Name = "FileErrorsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Setup()
    '---- File Error Recordset
    On Error GoTo ErrorHandler
    With FileErrorGrid
        Set .DataSource = DB.rsFileErrors                               'Set datasource for Grid
        .Refresh ssRefetchAndFireInitializeRow                          'Refresh grid with data
        .Bands(0).Columns(0).Hidden = True                              'This is the PrinterName column
        .Bands(0).Columns(1).Hidden = True                              'This is the Print Que column
        .Bands(0).Columns(2).Width = 8000                               'Image File Name Column
        .Bands(0).Columns(2).Activation = ssActivationActivateNoEdit
        .Bands(0).Columns(3).Activation = ssActivationActivateNoEdit    'Description
        .Bands(0).Columns(3).Width = 6000
        .Bands(0).Columns(4).Activation = ssActivationActivateNoEdit    'Time of error
        .Bands(0).Columns(4).Width = 4000
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrorHandler
    If Me.Width - 100 > 0 Then
        Me.FileErrorGrid.Width = Me.Width - 100
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  FileErrorsToolBars_ToolClick                          **
'**                                                                        **
'**  Description..:  Erase file error messages.                            **
'**                                                                        **
'****************************************************************************
Private Sub FileErrorsToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Dim FileName As String
    Select Case Tool.ID
        Case "ID_Erase"
        
            '--- RDR - attempt to reload error files into print queue here... then kill remaining errors
            With DB.rsFileErrors
                If MsgBox("Remove selected images from file error table?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
                    .MoveFirst
                    Do While Not .EOF
                        If FileErrorGrid.ActiveRow.Selected = True Then
                            FileName = Trim(.Fields("ImageFileName").Value)
                            If FileSystemHandle.FileExists(FileName) Then
                                Kill FileName
                            End If
                            .Delete adAffectCurrent
                        End If
                        .MoveNext
                    Loop
                End If
            End With
            'PrintQueGrid.Refresh ssRefetchAndFireInitializeRow
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":FileErrorsToolBars_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

