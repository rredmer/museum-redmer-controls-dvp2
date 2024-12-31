VERSION 5.00
Object = "{D8F5B61D-9152-4399-BF30-A1E4F3F072F6}#4.0#0"; "IGTabs40.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form ColorControlForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12120
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars ColorControlToolBars 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "ColorControlForm.frx":0000
      ToolBars        =   "ColorControlForm.frx":0D35
   End
   Begin ActiveTabs.SSActiveTabs ColorTab 
      Height          =   12105
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   21352
      _Version        =   262144
      TabCount        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHotTracking {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "ColorControlForm.frx":0DE9
      Begin ActiveTabs.SSActiveTabPanel RingAroundPanel 
         Height          =   11685
         Left            =   30
         TabIndex        =   1
         Top             =   390
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   20611
         _Version        =   262144
         TabGuid         =   "ColorControlForm.frx":0E7F
         Begin VB.TextBox DT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8250
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "1.32"
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox OFT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9780
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "0.12057"
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox PatchNum 
            Height          =   285
            Left            =   11550
            TabIndex        =   2
            Text            =   "7"
            Top             =   300
            Width           =   765
         End
         Begin UltraGrid.SSUltraGrid LabAimGrid 
            Height          =   945
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   1667
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   1
            LayoutFlags     =   72351744
            Override        =   "ColorControlForm.frx":0EA7
            Appearance      =   "ColorControlForm.frx":0EFD
            Caption         =   "Lab AIM"
         End
         Begin Threed.SSCheck IsLog 
            Height          =   315
            Left            =   8250
            TabIndex        =   6
            Top             =   660
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Use log for densitometer values"
         End
         Begin UltraGrid.SSUltraGrid ColorCircleAutoScan 
            Height          =   2775
            Left            =   60
            TabIndex        =   7
            Top             =   1050
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   4895
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   1
            LayoutFlags     =   72351744
            Override        =   "ColorControlForm.frx":0F39
            Appearance      =   "ColorControlForm.frx":0F8F
            Caption         =   "Color Circle Densitometer Readings"
         End
         Begin VB.Label Label2 
            Caption         =   "Output Factor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9660
            TabIndex        =   9
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Time Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8220
            TabIndex        =   8
            Top             =   60
            Width           =   1215
         End
      End
      Begin ActiveTabs.SSActiveTabPanel DailyTabPanel 
         Height          =   11685
         Left            =   30
         TabIndex        =   10
         Top             =   390
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   20611
         _Version        =   262144
         TabGuid         =   "ColorControlForm.frx":0FCB
         Begin UltraGrid.SSUltraGrid DailyCalibrationGrid 
            Height          =   1095
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   1931
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   1
            LayoutFlags     =   68161540
            BorderStyle     =   5
            ScrollBars      =   0
            RowConnectorStyle=   1
            ViewStyle       =   0
            BorderStyleCaption=   5
            AlphaBlendEnabled=   0   'False
            RowScrollRegions=   "ColorControlForm.frx":0FF3
            Override        =   "ColorControlForm.frx":1023
            Caption         =   "Daily Calibration Values"
         End
         Begin UltraGrid.SSUltraGrid CurrentExposureTime 
            Height          =   855
            Left            =   120
            TabIndex        =   12
            Top             =   1380
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   1508
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   1
            LayoutFlags     =   72351748
            BorderStyle     =   5
            ScrollBars      =   0
            ViewStyle       =   0
            BorderStyleCaption=   5
            AlphaBlendEnabled=   0   'False
            Override        =   "ColorControlForm.frx":1079
            CaptionAppearance=   "ColorControlForm.frx":10CF
            Caption         =   "Current Exposure Time"
         End
         Begin UltraGrid.SSUltraGrid ExposureTimeHistory 
            Height          =   8985
            Left            =   120
            TabIndex        =   13
            Top             =   2310
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   15849
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   1
            LayoutFlags     =   72613908
            BorderStyle     =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Bands           =   "ColorControlForm.frx":110B
            Override        =   "ColorControlForm.frx":1AAF
            Appearance      =   "ColorControlForm.frx":1B05
            Caption         =   "Exposure Time History"
         End
      End
   End
End
Attribute VB_Name = "ColorControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP-2 Quality Control                                     **
'**                                                                        **
'** Module.....: ColorControl                                              **
'**                                                                        **
'** Description: This module provides Color Calibration.                   **
'**                                                                        **
'** History....:                                                           **
'**    10/02/03 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Public ColorRingPass As Integer
Private Dif(2) As Double
Private U(2, 2), O(2, 2), d(2, 2) As Double

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Public Sub Setup()
    On Error GoTo ErrorHandler
    Dim Col As Integer
    UpdateExposureTimeGrids
    With LabAimGrid
        Set .DataSource = DB.rsLabAim
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Width = 900
        .Bands(0).Columns(0).Header.Appearance.BackColor = vbRed
        .Bands(0).Columns(0).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(0).Header.Caption = "Red"
        .Bands(0).Columns(1).Width = 900
        .Bands(0).Columns(1).Header.Appearance.BackColor = vbGreen
        .Bands(0).Columns(1).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(1).Header.Caption = "Green"
        .Bands(0).Columns(2).Width = 900
        .Bands(0).Columns(2).Header.Appearance.BackColor = vbBlue
        .Bands(0).Columns(2).Header.Caption = "Blue"
        .Bands(0).Columns(2).Header.Appearance.ForeColor = vbWhite
    End With
    With DailyCalibrationGrid
        Set .DataSource = DB.rsDailyCalibration
        .Refresh ssRefetchAndFireInitializeRow
        .BorderStyle = ssBorderStyleSolidLine
        .BorderStyleCaption = ssBorderStyleNone
        .Override.BorderStyleRow = ssBorderStyleNone
        .Override.BorderStyleCell = ssBorderStyleDefault
        .Bands(0).Columns(0).Hidden = True          'Printer Name
        .Bands(0).Columns(1).Hidden = True          'Block #
        .Bands(0).Columns(2).Activation = ssActivationDisabled  'ExposureName
        .Bands(0).Columns(2).CellAppearance.BorderColor = vbBlack
        .Bands(0).Columns(2).CellAppearance.BorderAlpha = ssAlphaOpaque
        .Bands(0).Columns(2).Header.Caption = "Setting"
        .Bands(0).Columns(2).Header.Appearance.BackColor = vbWhite
        .Bands(0).Columns(2).Header.Appearance.ForeColor = vbBlack
        .Bands(0).Columns(2).Header.Appearance.BorderColor = vbBlack
        .Bands(0).Columns(2).Width = 1100
        .Bands(0).Columns(3).Header.Appearance.BackColor = vbRed
        .Bands(0).Columns(3).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(3).Header.Caption = "Red"
        .Bands(0).Columns(3).Width = 900
        .Bands(0).Columns(4).Header.Appearance.BackColor = vbGreen
        .Bands(0).Columns(4).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(4).Header.Caption = "Green"
        .Bands(0).Columns(4).Width = 900
        .Bands(0).Columns(5).Header.Appearance.BackColor = vbBlue
        .Bands(0).Columns(5).Header.Caption = "Blue"
        .Bands(0).Columns(5).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(5).Width = 900
    End With
    UpdateColorCircleAutoScanGrid
    ColorRingPass = 0
    IsLog.Value = ssCBChecked
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ColorControlToolBars_ToolClick                        **
'**                                                                        **
'**  Description..:  This routine handles the toolbar.                     **
'**                                                                        **
'****************************************************************************
Private Sub ColorControlToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Select Case Tool.ID
        Case "ID_CalculateExposureTime"
            CalculateNewTime
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":ColorControlToolBars_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Resize                                           **
'**                                                                        **
'**  Description..:  This routine resizes controls on form resize.         **
'**                                                                        **
'****************************************************************************
Private Sub Form_Resize()
    On Error GoTo ErrorHandler
    If Me.Width - 300 > 0 Then
        ColorTab.Width = Me.Width - 100
        ExposureTimeHistory.Width = ColorTab.Width - 300
        ColorCircleAutoScan.Width = ColorTab.Width - 300
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateExposureTimeGrids                               **
'**                                                                        **
'**  Description..:  This routine initializes the exposure time grid       **
'**                                                                        **
'****************************************************************************
Private Sub UpdateExposureTimeGrids()
    On Error GoTo ErrorHandler
    '---- Current Exposure Time Grid
    With CurrentExposureTime
        Set .DataSource = DB.rsExposureTime
        .Refresh ssRefetchAndFireInitializeRow
        .BorderStyle = ssBorderStyleSolidLine
        .BorderStyleCaption = ssBorderStyleNone
        .Override.BorderStyleRow = ssBorderStyleNone
        .Override.BorderStyleCell = ssBorderStyleDefault
        .Bands(0).Columns(0).Hidden = True
        .Bands(0).Columns(1).Hidden = True
        .Bands(0).Columns(2).Header.Appearance.BackColor = vbRed
        .Bands(0).Columns(2).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(2).Width = 900
        .Bands(0).Columns(3).Header.Appearance.BackColor = vbGreen
        .Bands(0).Columns(3).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(3).Width = 900
        .Bands(0).Columns(4).Header.Appearance.BackColor = vbBlue
        .Bands(0).Columns(4).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(4).Width = 900
        .Bands(0).Columns(5).Header.Appearance.BackColor = vbRed
        .Bands(0).Columns(5).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(5).Width = 1100
        .Bands(0).Columns(6).Header.Appearance.BackColor = vbGreen
        .Bands(0).Columns(6).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(6).Width = 1100
        .Bands(0).Columns(7).Header.Appearance.BackColor = vbBlue
        .Bands(0).Columns(7).Header.Appearance.ForeColor = vbWhite
        .Bands(0).Columns(7).Width = 1100
        .Bands(0).Columns(8).Header.Appearance.BackColor = vbWhite
        .Bands(0).Columns(8).Header.Caption = "Date Modified"
        .Bands(0).Columns(8).Header.Appearance.ForeColor = vbBlack
        .Bands(0).Columns(8).Width = 1675
        .Bands(0).Header.Appearance.BorderAlpha = ssAlphaTransparent
    End With

    With ExposureTimeHistory
        Set .DataSource = DB.rsExposureTimeHistory
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":UpdateExposureTimeGrids", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub
'****************************************************************************
'**                                                                        **
'**  Procedure....:  SetColorTab                                           **
'**                                                                        **
'**  Description..:  This routine displays the color ring tab.             **
'**                                                                        **
'****************************************************************************
Public Function SetColorTab(index As Integer)
    On Error GoTo ErrorHandler
    ColorTab.Tabs(index).Selected = True
    ColorCircleAutoScan.SetFocus
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SetColorTab", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Function

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateColorCircleAutoScanGrid                         **
'**                                                                        **
'**  Description..:  This routine updates scan grid.                       **
'**                                                                        **
'****************************************************************************
Public Sub UpdateColorCircleAutoScanGrid()
    On Error GoTo ErrorHandler
    Dim Col As Integer
    With ColorCircleAutoScan
        Set .DataSource = DB.rsRingAroundsAutoScan
        .Refresh ssRefetchAndFireInitializeRow
        .Bands(0).Columns(0).Hidden = True          'Printer Name
        .Bands(0).Columns(1).Hidden = True          'Block #
        .Bands(0).Columns(2).Activation = ssActivationActivateOnly 'ExposureName
        .Bands(0).Columns(2).Header.Caption = "Exposure"
        .Bands(0).Columns(2).Width = 900
        For Col = 3 To .Bands(0).Columns.Count - 1
            If Left(.Bands(0).Columns(Col).Header.Caption, 4) = "ExpR" Then
                .Bands(0).Columns(Col).Header.Appearance.BackColor = vbRed
                .Bands(0).Columns(Col).Header.Appearance.ForeColor = vbWhite
            End If
            If Left(.Bands(0).Columns(Col).Header.Caption, 4) = "ExpG" Then
                .Bands(0).Columns(Col).Header.Appearance.BackColor = vbGreen
                .Bands(0).Columns(Col).Header.Appearance.ForeColor = vbWhite
            End If
            If Left(.Bands(0).Columns(Col).Header.Caption, 4) = "ExpB" Then
                .Bands(0).Columns(Col).Header.Appearance.BackColor = vbBlue
                .Bands(0).Columns(Col).Header.Appearance.ForeColor = vbWhite
            End If
            .Bands(0).Columns(Col).Width = 900
        Next
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":UpdateColorCircleAutoScanGrid", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CalculateNewTime                                      **
'**                                                                        **
'**  Description..:  This routine performs exposure time calculation.      **
'**                                                                        **
'****************************************************************************
Private Sub CalculateNewTime()
    On Error GoTo ErrorHandler
    Dim ColorNum As Integer, ScanNum As Integer
    Dim LogTen As Double
    Dim R(2), G(2), B(2), N(2), C(2), M(2), Y(2), T(2), P(2) As Double
    Dim TP(2), DET(2), Dx(2), LTi(2), ETi(2), EXTi(2), EXPTi(2) As Double
    Dim OF, DI, DD As Double
    Dim EXTR, EXTG, EXTB, EXPTR, EXPTG, EXPTB As Double

    If MsgBox("Calculate new exposure time from Daily Calibration Readings?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") <> vbYes Then
        Exit Sub
    End If

    ScanNum = CInt(PatchNum.Text)               'The Number of the Scan To use
    LogTen = Log(10#)                           'Simplify logic by putting log 10 into variable

    '---- Load Probe Times into Probe Array
    With DB.rsExposureTime
        TP(0) = Val(.Fields("Red8x10").Value)
        TP(1) = Val(.Fields("Green8x10").Value)
        TP(2) = Val(.Fields("Blue8x10").Value)
    End With
    
    '---- Load Color Circle into Arrays
    With DB.rsRingAroundsAutoScan
        .MoveFirst
        If IsLog.Value = ssCBChecked Then
            N(0) = .Fields("ExpRed" & ScanNum).Value        'Red
            N(1) = .Fields("ExpGreen" & ScanNum).Value      'Green
            N(2) = .Fields("ExpBlue" & ScanNum).Value       'Blue
        Else
            N(0) = Log(Val(.Fields("ExpRed" & ScanNum).Value)) / LogTen  ' Rpict
            N(1) = Log(Val(.Fields("ExpGreen" & ScanNum).Value)) / LogTen  ' Rpict
            N(2) = Log(Val(.Fields("ExpBlue" & ScanNum).Value)) / LogTen  ' Rpict
        End If
        .MoveNext
        If IsLog.Value = ssCBChecked Then
            R(0) = .Fields("ExpRed" & ScanNum).Value        'Red
            R(1) = .Fields("ExpGreen" & ScanNum).Value      'Green
            R(2) = .Fields("ExpBlue" & ScanNum).Value       'Blue
        Else
            R(0) = Log(Val(.Fields("ExpRed" & ScanNum).Value)) / LogTen  ' Rpict
            R(1) = Log(Val(.Fields("ExpGreen" & ScanNum).Value)) / LogTen  ' Rpict
            R(2) = Log(Val(.Fields("ExpBlue" & ScanNum).Value)) / LogTen  ' Rpict
        End If
        .MoveNext
        If IsLog.Value = ssCBChecked Then
            G(0) = .Fields("ExpRed" & ScanNum).Value        'Red
            G(1) = .Fields("ExpGreen" & ScanNum).Value      'Green
            G(2) = .Fields("ExpBlue" & ScanNum).Value       'Blue
        Else
            G(0) = Log(Val(.Fields("ExpRed" & ScanNum).Value)) / LogTen  ' Rpict
            G(1) = Log(Val(.Fields("ExpGreen" & ScanNum).Value)) / LogTen  ' Rpict
            G(2) = Log(Val(.Fields("ExpBlue" & ScanNum).Value)) / LogTen  ' Rpict
        End If
        .MoveNext
        If IsLog.Value = ssCBChecked Then
            B(0) = .Fields("ExpRed" & ScanNum).Value        'Red
            B(1) = .Fields("ExpGreen" & ScanNum).Value      'Green
            B(2) = .Fields("ExpBlue" & ScanNum).Value       'Blue
        Else
            B(0) = Log(Val(.Fields("ExpRed" & ScanNum).Value)) / LogTen  ' Rpict
            B(1) = Log(Val(.Fields("ExpGreen" & ScanNum).Value)) / LogTen  ' Rpict
            B(2) = Log(Val(.Fields("ExpBlue" & ScanNum).Value)) / LogTen  ' Rpict
        End If
        .MoveNext
        If IsLog.Value = ssCBChecked Then
            C(0) = .Fields("ExpRed" & ScanNum).Value        'Red
            C(1) = .Fields("ExpGreen" & ScanNum).Value      'Green
            C(2) = .Fields("ExpBlue" & ScanNum).Value       'Blue
        Else
            C(0) = Log(Val(.Fields("ExpRed" & ScanNum).Value)) / LogTen  ' Rpict
            C(1) = Log(Val(.Fields("ExpGreen" & ScanNum).Value)) / LogTen  ' Rpict
            C(2) = Log(Val(.Fields("ExpBlue" & ScanNum).Value)) / LogTen  ' Rpict
        End If
        .MoveNext
        If IsLog.Value = ssCBChecked Then
            M(0) = .Fields("ExpRed" & ScanNum).Value        'Red
            M(1) = .Fields("ExpGreen" & ScanNum).Value      'Green
            M(2) = .Fields("ExpBlue" & ScanNum).Value       'Blue
        Else
            M(0) = Log(Val(.Fields("ExpRed" & ScanNum).Value)) / LogTen  ' Rpict
            M(1) = Log(Val(.Fields("ExpGreen" & ScanNum).Value)) / LogTen  ' Rpict
            M(2) = Log(Val(.Fields("ExpBlue" & ScanNum).Value)) / LogTen  ' Rpict
        End If
        .MoveNext
        If IsLog.Value = ssCBChecked Then
            Y(0) = .Fields("ExpRed" & ScanNum).Value        'Red
            Y(1) = .Fields("ExpGreen" & ScanNum).Value      'Green
            Y(2) = .Fields("ExpBlue" & ScanNum).Value       'Blue
        Else
            Y(0) = Log(Val(.Fields("ExpRed" & ScanNum).Value)) / LogTen  ' Rpict
            Y(1) = Log(Val(.Fields("ExpGreen" & ScanNum).Value)) / LogTen  ' Rpict
            Y(2) = Log(Val(.Fields("ExpBlue" & ScanNum).Value)) / LogTen  ' Rpict
        End If
        
    End With

    '---- Load Probe Times into Probe Array
    With DB.rsDailyCalibration
        .MoveFirst
        If IsLog.Value = ssCBChecked Then
            P(0) = Val(.Fields("ExpRed").Value)
            P(1) = Val(.Fields("ExpGreen").Value)
            P(2) = Val(.Fields("ExpBlue").Value)
        Else
            P(0) = Log(Val(.Fields("ExpRed").Value)) / LogTen ' Probe
            P(1) = Log(Val(.Fields("ExpGreen").Value)) / LogTen ' Probe
            P(2) = Log(Val(.Fields("ExpBlue").Value)) / LogTen ' Probe
        End If
    End With

    '---- Load Lab Aim into Target Array
    With DB.rsLabAim
        If IsLog.Value = ssCBChecked Then
            T(0) = Val(.Fields("AimRed").Value)
            T(1) = Val(.Fields("AimGreen").Value)
            T(2) = Val(.Fields("AimBlue").Value)
        Else
            T(0) = Log(Val(.Fields("AimRed").Value)) / LogTen ' Target
            T(1) = Log(Val(.Fields("AimGreen").Value)) / LogTen ' Target
            T(2) = Log(Val(.Fields("AimBlue").Value)) / LogTen ' Target
        End If
    End With

    '---- Probe - Target
    For ColorNum = 0 To 2
        Dif(ColorNum) = P(ColorNum) - T(ColorNum)
    Next

    '---- Create Minus-Slope U and Plus-Slope O
    For ColorNum = 0 To 2
        U(ColorNum, 0) = N(ColorNum) - R(ColorNum)
        O(ColorNum, 0) = C(ColorNum) - N(ColorNum)
        U(ColorNum, 1) = N(ColorNum) - G(ColorNum)
        O(ColorNum, 1) = M(ColorNum) - N(ColorNum)
        U(ColorNum, 2) = N(ColorNum) - B(ColorNum)
        O(ColorNum, 2) = Y(ColorNum) - N(ColorNum)
    Next

    DD = Determinant(-1)
    DET(0) = Determinant(0)  ' Red
    DET(1) = Determinant(1)  ' Green
    DET(2) = Determinant(2)  ' Blue

    For ColorNum = 0 To 2
        Dx(ColorNum) = DET(ColorNum) / DD
    Next
    
    For ColorNum = 0 To 2
        LTi(ColorNum) = (Log(TP(ColorNum))) / LogTen
    Next
    
    For ColorNum = 0 To 2
        ETi(ColorNum) = LTi(ColorNum) - Val(OFT.Text) * Dx(ColorNum)
    Next
    
    For ColorNum = 0 To 2
        EXTi(ColorNum) = 10 ^ ETi(ColorNum)
    Next

    ''For I = 0 To 2: EXPTi(ColorNum) = 0.1 * Int(10 * EXTi(ColorNum) + 0.5): ET(ColorNum) = Str$(EXPTi(ColorNum)): Next
    
    '---- Calculate new Exposure Times
    For ColorNum = 0 To 2
        EXPTi(ColorNum) = Int(EXTi(ColorNum) + 0.5)
        '---- this is where peter's code updated the display  - store to new exposure times:      ET(ColorNum) = Str$(EXPTi(ColorNum))
    Next
    
    
    '---- Copy current exposure time record to history
    With DB.rsExposureTimeHistory
        .AddNew
        .Fields("PrinterName").Value = PrinterName
        .Fields("DateModified").Value = Now
        .Fields("Red8x10").Value = DB.rsExposureTime.Fields("Red8x10").Value
        .Fields("Green8x10").Value = DB.rsExposureTime.Fields("Green8x10").Value
        .Fields("Blue8x10").Value = DB.rsExposureTime.Fields("Blue8x10").Value
        .Fields("Red11x14").Value = DB.rsExposureTime.Fields("Red11x14").Value
        .Fields("Green11x14").Value = DB.rsExposureTime.Fields("Green11x14").Value
        .Fields("Blue11x14").Value = DB.rsExposureTime.Fields("Blue11x14").Value
        .Fields("Red8x10_Density").Value = DB.rsDailyCalibration.Fields("ExpRed").Value
        .Fields("Green8x10_Density").Value = DB.rsDailyCalibration.Fields("ExpGreen").Value
        .Fields("Blue8x10_Density").Value = DB.rsDailyCalibration.Fields("ExpBlue").Value
        DB.rsDailyCalibration.MoveNext
        .Fields("Lamphouse_Red").Value = DB.rsDailyCalibration.Fields("ExpRed").Value
        .Fields("Lamphouse_Green").Value = DB.rsDailyCalibration.Fields("ExpGreen").Value
        .Fields("Lamphouse_Blue").Value = DB.rsDailyCalibration.Fields("ExpBlue").Value
        DB.rsDailyCalibration.MovePrevious
        .UpdateBatch adAffectCurrent
    End With
    
    '---- Update current exposure time
    With DB.rsExposureTime
        .Fields("Red8x10").Value = EXPTi(0)
        .Fields("Green8x10").Value = EXPTi(1)
        .Fields("Blue8x10").Value = EXPTi(2)
        .Fields("Red11x14").Value = EXPTi(0)
        .Fields("Green11x14").Value = EXPTi(1)
        .Fields("Blue11x14").Value = EXPTi(2)
        .Fields("DateModified").Value = Now
        .UpdateBatch adAffectCurrent
    End With
    
    AppLog DebugMsg, "Time Change   DT = " + Str$(DI)
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, " "
    AppLog DebugMsg, "Color - Circle - Values"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "RED - Picture     (Tred/" + DT + " = Red Time):"
    AppLog DebugMsg, "  cyan   R(0) = " + Str$(R(0))
    AppLog DebugMsg, "  mag.   R(1) = " + Str$(R(1))
    AppLog DebugMsg, "  yell.  R(2) = " + Str$(R(2))
    AppLog DebugMsg, "GREEN - Picture   (Tgreen/" + DT + " = Green Time):"
    AppLog DebugMsg, "  cyan   G(0) = " + Str$(G(0))
    AppLog DebugMsg, "  mag.   G(1) = " + Str$(G(1))
    AppLog DebugMsg, "  yell.  G(2) = " + Str$(G(2))
    AppLog DebugMsg, "BLUE - Picture    (Tblue/" + DT + " = Blue Time):"
    AppLog DebugMsg, "  cyan   B(0) = " + Str$(B(0))
    AppLog DebugMsg, "  mag.   B(1) = " + Str$(B(1))
    AppLog DebugMsg, "  yell.  B(2) = " + Str$(B(2))
    AppLog DebugMsg, "NORM - Picture    (Normal Time):"
    AppLog DebugMsg, "  cyan   N(0) = " + Str$(N(0))
    AppLog DebugMsg, "  mag.   N(1) = " + Str$(N(1))
    AppLog DebugMsg, "  yell.  N(2) = " + Str$(N(2))
    AppLog DebugMsg, "CYAN - Picture    (Tred*" + DT + " = Red Time):"
    AppLog DebugMsg, "  cyan   C(0) = " + Str$(C(0))
    AppLog DebugMsg, "  mag.   C(1) = " + Str$(C(1))
    AppLog DebugMsg, "  yell.  C(2) = " + Str$(C(2))
    AppLog DebugMsg, "MAGENTA - Picture (Tgreen*" + DT + " = Green Time):"
    AppLog DebugMsg, "  cyan   M(0) = " + Str$(M(0))
    AppLog DebugMsg, "  mag.   M(1) = " + Str$(M(1))
    AppLog DebugMsg, "  yell.  M(2) = " + Str$(M(2))
    AppLog DebugMsg, "YELLOW - Picture  (Tblue*" + DT + " = Blue Time):"
    AppLog DebugMsg, "  cyan   Y(0) = " + Str$(Y(0))
    AppLog DebugMsg, "  mag.   Y(1) = " + Str$(Y(1))
    AppLog DebugMsg, "  yell.  Y(2) = " + Str$(Y(2))
    AppLog DebugMsg, " "
    AppLog DebugMsg, "U-Matrix = Minus - Slope"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "        Npict-Rpict      Npict-Gpict      Npict-Bpict"
    AppLog DebugMsg, "  cyan  U(0,0) = " + Str$(U(0, 0)) + "    U(0,1) = " + Str$(U(0, 1)) + "    U(0,1) = " + Str$(U(0, 2))
    AppLog DebugMsg, "  mag.  U(1,0) = " + Str$(U(1, 0)) + "    U(1,1) = " + Str$(U(1, 1)) + "    U(1,1) = " + Str$(U(1, 2))
    AppLog DebugMsg, "  yell. U(2,0) = " + Str$(U(2, 0)) + "    U(2,1) = " + Str$(U(2, 1)) + "    U(2,1) = " + Str$(U(2, 2))
    AppLog DebugMsg, " "
    AppLog DebugMsg, "O-Matrix = Plus - Slope"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "        Cpict-Npict      Mpict-Npict      Ypict-Npict"
    AppLog DebugMsg, "  cyan  O(0,0) = " + Str$(O(0, 0)) + "    O(0,1) = " + Str$(O(0, 1)) + "    O(0,1) = " + Str$(O(0, 2))
    AppLog DebugMsg, "  mag.  O(1,0) = " + Str$(O(1, 0)) + "    O(1,1) = " + Str$(O(1, 1)) + "    O(1,1) = " + Str$(O(1, 2))
    AppLog DebugMsg, "  yell. O(2,0) = " + Str$(O(2, 0)) + "    O(2,1) = " + Str$(O(2, 1)) + "    O(2,1) = " + Str$(O(2, 2))
    AppLog DebugMsg, " "
    AppLog DebugMsg, "Probe  and  Target                    Difference  Probe - Target"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "cyan   P(0) = " + Str$(P(0)) + "     T(0) = " + Str$(T(0)) + "     Dif(0) = " + Str$(Dif(0))
    AppLog DebugMsg, "mag.   P(1) = " + Str$(P(1)) + "     T(1) = " + Str$(T(1)) + "     Dif(1) = " + Str$(Dif(1))
    AppLog DebugMsg, "yell.  P(2) = " + Str$(P(2)) + "     T(2) = " + Str$(T(2)) + "     Dif(2) = " + Str$(Dif(2))
    AppLog DebugMsg, " "
    AppLog DebugMsg, "Exposure Time Probe"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "Red   ETiP(0) = " + Str$(TP(0)) + "    LTiP(0) = Log(TP(0))/L = " + Str$(LTi(0))
    AppLog DebugMsg, "Green ETiP(1) = " + Str$(TP(1)) + "    LTiP(1) = Log(TP(1))/L = " + Str$(LTi(1))
    AppLog DebugMsg, "Blue  ETiP(2) = " + Str$(TP(2)) + "    LTiP(2) = Log(TP(2))/L = " + Str$(LTi(2))
    AppLog DebugMsg, " "
    AppLog DebugMsg, "Determinats  D  DXred  DXgreen  DXblue"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "D = det D(m,n) = " + Str$(DD)
    AppLog DebugMsg, " "
    AppLog DebugMsg, "DXred   = " + Str$(DET(0)) + "       Xred   =   DXred/D = " + Str$(Dx(0))
    AppLog DebugMsg, "DXgreen = " + Str$(DET(1)) + "       Xgreen = DXgreen/D = " + Str$(Dx(1))
    AppLog DebugMsg, "DXblue  = " + Str$(DET(2)) + "       Xblue  =  DXblue/D = " + Str$(Dx(2))
    AppLog DebugMsg, " "
    AppLog DebugMsg, "Output - Factor   OF = " + Str$(OF)
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, " "
    AppLog DebugMsg, "Exponent ETx = LTx - OF * Ex        10 ^ ETx"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "  Red     ETi(0) = " + Str$(ETi(0)) + "    10 ^ ETi(0) = " + Str$(EXTi(0))
    AppLog DebugMsg, "  Green   ETi(1) = " + Str$(ETi(1)) + "    10 ^ ETi(1) = " + Str$(EXTi(1))
    AppLog DebugMsg, "  Blue    ETi(2) = " + Str$(ETi(2)) + "    10 ^ ETi(2) = " + Str$(EXTi(2))
    AppLog DebugMsg, " "
    AppLog DebugMsg, "New Exposure Times"
    AppLog DebugMsg, "========================================================"
    AppLog DebugMsg, "  Red     EXPTi(0) = " & EXPTi(0)
    AppLog DebugMsg, "  Green   EXPTi(1) = " & EXPTi(1)
    AppLog DebugMsg, "  Blue    EXPTi(2) = " & EXPTi(2)
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":CalculateNewTime", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Determinant                                           **
'**                                                                        **
'**  Description..:  This routine finds determinant of O/U Matrixes.       **
'**                                                                        **
'****************************************************************************
Private Function Determinant(RGB As Integer) As Double
    On Error GoTo ErrorHandler
    Dim ColorNum As Integer
    For ColorNum = 0 To 2
        If Dif(0) < 0 Then d(0, ColorNum) = U(0, ColorNum) Else d(0, ColorNum) = O(0, ColorNum) ' cyan
        If Dif(1) < 0 Then d(1, ColorNum) = U(1, ColorNum) Else d(1, ColorNum) = O(1, ColorNum) ' mag.
        If Dif(2) < 0 Then d(2, ColorNum) = U(2, ColorNum) Else d(2, ColorNum) = O(2, ColorNum) ' yell.
    Next
    If RGB >= 0 Then
        For ColorNum = 0 To 2
            d(ColorNum, RGB) = Dif(ColorNum)
        Next
    End If
    Determinant = d(0, 0) * d(1, 1) * d(2, 2) - d(0, 0) * d(1, 2) * d(2, 1) + _
                  d(0, 1) * d(1, 2) * d(2, 0) - d(0, 1) * d(1, 0) * d(2, 2) + _
                  d(0, 2) * d(1, 0) * d(2, 1) - d(0, 2) * d(1, 1) * d(2, 0)
    Exit Function
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":Determinant", Err.Number, Err.LastDllError, Err.Source, Err.Description, False
End Function

'--- Test function
'Private Sub Rtest_Click()
'    Dim I As Integer
'    For I = 0 To 2
'        PT(ColorNum).Text = RT(ColorNum).Text ' Probe = REDpicture        'rt becomes gt,bt,ct,mt,yt
'        TT(ColorNum).Text = NT(ColorNum).Text ' Target = NORMpict
'        TP(ColorNum) = Val(OT(ColorNum).Text) ' Exp.Times for Probe
'    Next I
'    If Not ValuesOK Then Exit Sub
'    TP(0) = TP(0) / DI ' Tred/DI            0=red,1=green,2=blue,3=cyan,4=magenta,5=yellow
'    Calculation
'End Sub


