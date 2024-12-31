VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form SettingsControlForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   ScaleHeight     =   12135
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin UltraGrid.SSUltraGrid OptionsGrid 
      Height          =   4095
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   7223
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
      Bands           =   "SettingsControlForm.frx":0000
      Appearance      =   "SettingsControlForm.frx":023C
      Caption         =   "OptionsGrid"
   End
   Begin UltraGrid.SSUltraGrid PrinterSettings 
      Height          =   6645
      Left            =   30
      TabIndex        =   1
      Top             =   4260
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   11721
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   72613908
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "SettingsControlForm.frx":0278
      Override        =   "SettingsControlForm.frx":04B7
      Appearance      =   "SettingsControlForm.frx":04FF
      Caption         =   "PrinterSettings"
   End
End
Attribute VB_Name = "SettingsControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DatabaseFolder As String = "Database"


Public Sub Setup()

    On Error GoTo ErrorHandler
    
    '---- Validate the Printer Settings Files
    Dim TargetPath As String
    With FileSystemHandle
        TargetPath = OffsetFilePath & Trim(PrinterName) & "\"
        If .FolderExists(TargetPath) = False Then
            .CreateFolder TargetPath
        End If
        TargetPath = OffsetFilePath & Trim(PrinterName) & "\Database\"
        If .FolderExists(TargetPath) = False Then
            .CopyFolder SettingsFolder & "Database", OffsetFilePath & Trim(PrinterName) & "\"
        End If
        TargetPath = OffsetFilePath & Trim(PrinterName) & "\Program_Files\"
        If .FolderExists(TargetPath) = False Then
            .CopyFolder SettingsFolder & "Program_Files", OffsetFilePath & Trim(PrinterName) & "\"
        End If
        TargetPath = OffsetFilePath & Trim(PrinterName) & "\Images\"
        If .FolderExists(TargetPath) = False Then
            .CopyFolder SettingsFolder & "Images", OffsetFilePath & Trim(PrinterName) & "\"
        End If
    End With

    With OptionsGrid
        Set .DataSource = DB.rsOptions
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Activation = ssActivationActivateOnly
        .Bands(0).Columns(1).Header.Caption = "#"
        .Bands(0).Columns(1).Width = 600
        .Bands(0).Columns(2).Activation = ssActivationActivateOnly
        .Bands(0).Columns(2).Header.Caption = "Name"
        .Bands(0).Columns(2).Width = 7000
        .Bands(0).Columns(3).Header.Caption = "Value"
        .Bands(0).Columns(3).Width = 1200
    End With
    
    With PrinterSettings
        Set .DataSource = DB.rsSettings
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Activation = ssActivationActivateOnly
        .Bands(0).Columns(1).Header.Caption = "#"
        .Bands(0).Columns(1).Width = 600
        .Bands(0).Columns(2).Activation = ssActivationActivateOnly
        .Bands(0).Columns(2).Header.Caption = "Name"
        .Bands(0).Columns(2).Width = 3000
        .Bands(0).Columns(3).Header.Caption = "Value"
        .Bands(0).Columns(3).Width = 10000
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

