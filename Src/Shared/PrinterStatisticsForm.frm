VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form PrinterStatisticsForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17325
   LinkTopic       =   "Form1"
   ScaleHeight     =   12540
   ScaleWidth      =   17325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars PrinterStatisticsToolBars 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "PrinterStatisticsForm.frx":0000
      ToolBars        =   "PrinterStatisticsForm.frx":0CF3
   End
   Begin UltraGrid.SSUltraGrid PrinterStatistics 
      Bindings        =   "PrinterStatisticsForm.frx":0D7A
      Height          =   10425
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   18389
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   72613908
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "PrinterStatisticsForm.frx":0D9D
      Override        =   "PrinterStatisticsForm.frx":10E3
      Appearance      =   "PrinterStatisticsForm.frx":1161
      Caption         =   "PrinterStatistics"
   End
End
Attribute VB_Name = "PrinterStatisticsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Setup()
    
    Set PrinterStatistics.DataSource = DB.rsPrinterStatistics
    PrinterStatistics.Refresh ssRefetchAndFireInitializeRow

End Sub


Private Sub PrinterStatisticsToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_Reset"
        If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Reset Statistics") = vbYes Then
            With DB
                .StatAverageAdvanceTime = 0
                .StatAverageExposureTime = 0
                .StatAverageServerTime = 0
                .StatExposuresPerSecond = 0
                .StatPaperUsed = 0
                .StatTimeRunning = 0
                .StatTotalExposures = 0
                .StatTotalImages = 0
                .UpdateStatistics
                PrinterStatistics.Refresh ssRefetchAndFireInitializeRow
            End With
        End If
    End Select
End Sub
