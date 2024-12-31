VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form PrintSizesform 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin UltraGrid.SSUltraGrid PrintSizeGrid 
      Height          =   10695
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   18865
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
      Bands           =   "PrintSizesform.frx":0000
      Appearance      =   "PrintSizesform.frx":023C
      Caption         =   "PrintSizeGrid"
   End
   Begin Threed.SSCommand PrintSizeSettingButton 
      Height          =   765
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   10800
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1349
      _Version        =   262144
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PrintSizesform.frx":0278
      Caption         =   "Add"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin Threed.SSCommand PrintSizeSettingButton 
      Height          =   765
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   10800
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1349
      _Version        =   262144
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "PrintSizesform.frx":059A
      Caption         =   "Erase"
      Alignment       =   8
      PictureAlignment=   6
   End
End
Attribute VB_Name = "PrintSizesform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Setup()

    With PrintSizeGrid
        Set .DataSource = DB.rsPrintSizes
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True             'Hide printer name column
        .Bands(0).Columns(1).Header.Caption = "Print Size"
        .Bands(0).Columns(1).Width = 1200
        .Bands(0).Columns(2).Header.Caption = "Enable"
        .Bands(0).Columns(2).Width = 750
        .Bands(0).Columns(3).Header.Caption = "Aspect"
        .Bands(0).Columns(3).Width = 750
        .Bands(0).Columns(4).Header.Caption = "Caption"
        .Bands(0).Columns(4).Width = 800
        .Bands(0).Columns(5).Header.Caption = "Rotate"
        .Bands(0).Columns(5).Width = 800
        .Bands(0).Columns(6).Header.Caption = "11x14"
        .Bands(0).Columns(6).Width = 600
    
        .Bands(0).Columns(7).Header.Caption = "Video"
        .Bands(0).Columns(7).Width = 700
        .Bands(0).Columns(8).Header.Caption = "Left Mask"
        .Bands(0).Columns(8).Width = 1100
        .Bands(0).Columns(9).Header.Caption = "Right Mask"
        .Bands(0).Columns(9).Width = 1100
        .Bands(0).Columns(10).Header.Caption = "8x10 Flap"
        .Bands(0).Columns(10).Width = 1000
        .Bands(0).Columns(11).Header.Caption = "7x10 Flap"
        .Bands(0).Columns(11).Width = 1000
        .Bands(0).Columns(12).Header.Caption = "Right Flap"
        .Bands(0).Columns(12).Width = 1000
        .Bands(0).Columns(13).Header.Caption = "PreFeed"
        .Bands(0).Columns(13).Width = 1000
        .Bands(0).Columns(14).Header.Caption = "Feed"
        .Bands(0).Columns(14).Width = 800
        .Bands(0).Columns(15).Header.Caption = "Punch Feed"
        .Bands(0).Columns(15).Width = 1200
    End With


End Sub


Private Sub PrintSizeSettingButton_Click(index As Integer)
    Dim SizeString As String
    With DB.rsPrintSizes
        Select Case index
            Case 0                                      'Add
                DB.AddNewSize
            Case 1                                      'Erase
                If Not .BOF And Not .EOF Then
                    SizeString = .Fields("PrintSize").Value
                    If MsgBox("Delete Print Size [" & SizeString & "] from List?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") = vbYes Then
                        AppLog InfoMsg, "PrintSizeSettingButton_Click,Removed " & Trim(SizeString) & " from print size list."
                        .Delete adAffectCurrent
                    End If
                End If
        End Select
    End With
End Sub


