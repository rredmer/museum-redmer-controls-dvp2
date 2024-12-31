VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form PrintQueHistoryForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12690
   ScaleWidth      =   17430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin UltraGrid.SSUltraGrid PrintHistoryGrid 
      Height          =   6765
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   17325
      _ExtentX        =   30559
      _ExtentY        =   11933
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   72351748
      BorderStyle     =   5
      Override        =   "PrintQueHistoryForm.frx":0000
      Appearance      =   "PrintQueHistoryForm.frx":007E
      Caption         =   "PrintHistoryGrid"
   End
End
Attribute VB_Name = "PrintQueHistoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Setup()
    Dim Col As Integer
    
    '---- Print History Recordset
    With PrintHistoryGrid
        Set .DataSource = DB.rsPrintHistory                             'Set datasource for Grid
        .Refresh ssRefetchAndFireInitializeRow                          'Refresh grid with data
        .Bands(0).Columns(0).Hidden = True                              'This is the PrinterName column
        .Bands(0).Columns(2).Width = 6000                               'Image File Name Column
        For Col = 0 To .Bands(0).Columns.Count - 1
            .Bands(0).Columns(Col).Activation = ssActivationActivateNoEdit
        Next
    End With
    

End Sub
