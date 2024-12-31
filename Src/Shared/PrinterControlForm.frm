VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form PrinterControlForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11655
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars MainToolBar 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "PrinterControlForm.frx":0000
      ToolBars        =   "PrinterControlForm.frx":0CFE
   End
   Begin VB.TextBox txtPrinterDescription 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   300
      Width           =   8055
   End
   Begin UltraGrid.SSUltraGrid PrinterStatisticsGrid 
      Height          =   4635
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   8176
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108864
      Caption         =   "Printer Statistics"
   End
   Begin VB.Label PrinterDescriptionLabel 
      Caption         =   "Printer Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   330
      Width           =   2025
   End
End
Attribute VB_Name = "PrinterControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '---- Initialize Form Controls
End Sub


Public Sub Setup()
    '---- Initialize form controls when form is loaded re-initialized
    txtPrinterDescription.Text = MainForm.MainExplorer.SelectedNode.Text
    
    With PrinterStatisticsGrid
        Set .DataSource = DB.rsPrinterStatistics
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        
    End With
    
End Sub

Private Sub MainToolBar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Update"
            '---- Update the Description
            If Not DB.rsPrinterList.EOF Then
                DB.rsPrinterList.Fields("Description").Value = txtPrinterDescription.Text
                MainForm.MainExplorer.SelectedNode.Text = txtPrinterDescription.Text
            End If
    End Select
End Sub
