VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form DigitalSettingsForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12795
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin UltraGrid.SSUltraGrid DigitalOutputsGrid 
      Height          =   11535
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   20346
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   71565332
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "DigitalSettingsForm.frx":0000
      Appearance      =   "DigitalSettingsForm.frx":023C
      Caption         =   "Digital Outputs"
   End
   Begin UltraGrid.SSUltraGrid DigitalInputsGrid 
      Height          =   11535
      Left            =   6450
      TabIndex        =   1
      Top             =   30
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   20346
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   71565332
      BorderStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "DigitalSettingsForm.frx":0278
      Appearance      =   "DigitalSettingsForm.frx":04B4
      Caption         =   "Digital Inputs"
   End
End
Attribute VB_Name = "DigitalSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Setup()
    
    With DigitalOutputsGrid
        Set .DataSource = DB.rsOutputs
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Activation = ssActivationActivateOnly
        .Bands(0).Columns(1).Header.Caption = "#"
        .Bands(0).Columns(1).Width = 600
        .Bands(0).Columns(2).Header.Caption = "Name"
        .Bands(0).Columns(2).Width = 5000
    End With
    
    With DigitalInputsGrid
        Set .DataSource = DB.rsInputs
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Activation = ssActivationActivateOnly
        .Bands(0).Columns(1).Header.Caption = "#"
        .Bands(0).Columns(1).Width = 600
        .Bands(0).Columns(2).Header.Caption = "Name"
        .Bands(0).Columns(2).Width = 4000
        .Bands(0).Columns(3).Header.Caption = "Enabled"
        .Bands(0).Columns(3).Width = 1200
        .Bands(0).Columns(4).Header.Caption = "TimeOut"
        .Bands(0).Columns(4).Width = 1200
    End With
    
End Sub
