VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{00130003-B1BA-11CE-ABC6-F5B2E79D9E3F}#1.0#0"; "ltocx13n.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#8.0#0"; "PVPrgbar.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Object = "{14ACBB92-9C4A-4C45-AFD2-7AE60E71E5B3}#4.0#0"; "IGSplitter40.ocx"
Begin VB.Form SizeSettingsForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12075
   ScaleWidth      =   17595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSSplitter.SSSplitter PrintSizeSplitter 
      Height          =   12075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17595
      _ExtentX        =   31036
      _ExtentY        =   21299
      _Version        =   262144
      AutoSize        =   1
      SplitterResizeStyle=   1
      PaneTree        =   "SizeSettingsForm.frx":0000
      Begin Threed.SSFrame PreviewFrame 
         Height          =   7725
         Left            =   7695
         TabIndex        =   21
         Top             =   4320
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   13626
         _Version        =   262144
         Caption         =   "Preview"
         Begin PVProgressBarLib.PVProgressBar UpdateProgress 
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   7350
            Width           =   7335
            _Version        =   524288
            _ExtentX        =   12938
            _ExtentY        =   556
            _StockProps     =   237
            Value           =   0
            FillColor       =   12582912
         End
         Begin LEADLib.LEAD TargetImage 
            Height          =   7095
            Left            =   60
            TabIndex        =   22
            Top             =   210
            Width           =   9690
            _Version        =   65540
            _ExtentX        =   17092
            _ExtentY        =   12515
            _StockProps     =   229
            BorderStyle     =   1
            ScaleHeight     =   471
            ScaleWidth      =   644
            DataField       =   ""
            BitmapDataPath  =   ""
            AnnDataPath     =   ""
            PaintSizeMode   =   3
            PanWinTitle     =   "PanWindow"
            PanWinPointer   =   5
            CLeadCtrl       =   0
            AutoPan         =   -1  'True
         End
         Begin MSComctlLib.Slider ZoomSlider 
            Height          =   285
            Left            =   7350
            TabIndex        =   23
            Top             =   7350
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            _Version        =   393216
            Max             =   100
         End
      End
      Begin VB.Frame PrintSizeFrame 
         Caption         =   "Print Size Settings"
         Height          =   4200
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   17535
         Begin UltraGrid.SSUltraGrid PrintSizeSettings 
            Bindings        =   "SizeSettingsForm.frx":0072
            Height          =   3915
            Left            =   60
            TabIndex        =   2
            Top             =   210
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   6906
            _Version        =   131072
            GridFlags       =   17040388
            UpdateMode      =   1
            LayoutFlags     =   72613908
            BorderStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Bands           =   "SizeSettingsForm.frx":008E
            Override        =   "SizeSettingsForm.frx":087D
            Appearance      =   "SizeSettingsForm.frx":08FB
            Caption         =   "PrintSizeSettings"
         End
         Begin Threed.SSFrame PaperMaskFrame 
            Height          =   4005
            Left            =   13200
            TabIndex        =   3
            Top             =   120
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   7064
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Paper Mask Control"
            Begin VB.CheckBox LeftMaskHome 
               Caption         =   "Left Switch"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   90
               TabIndex        =   13
               Top             =   2580
               Width           =   1905
            End
            Begin VB.CheckBox RightMaskHome 
               Caption         =   "Right Switch"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   2100
               TabIndex        =   12
               Top             =   2580
               Width           =   1995
            End
            Begin Threed.SSCommand CropCommand 
               Height          =   885
               Index           =   2
               Left            =   60
               TabIndex        =   4
               Top             =   240
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   1561
               _Version        =   262144
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "SizeSettingsForm.frx":0937
               Caption         =   "Mask Left In"
               Alignment       =   8
               PictureAlignment=   11
            End
            Begin Threed.SSCommand CropCommand 
               Height          =   885
               Index           =   3
               Left            =   60
               TabIndex        =   5
               Top             =   1185
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   1561
               _Version        =   262144
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "SizeSettingsForm.frx":0D89
               Caption         =   "Mask Left Out"
               Alignment       =   8
               PictureAlignment=   11
            End
            Begin Threed.SSCommand CropCommand 
               Height          =   885
               Index           =   4
               Left            =   2100
               TabIndex        =   6
               Top             =   240
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   1561
               _Version        =   262144
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "SizeSettingsForm.frx":11DB
               Caption         =   "Mask Right In"
               Alignment       =   8
               PictureAlignment=   11
            End
            Begin Threed.SSCommand CropCommand 
               Height          =   885
               Index           =   5
               Left            =   2100
               TabIndex        =   7
               Top             =   1185
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   1561
               _Version        =   262144
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "SizeSettingsForm.frx":162D
               Caption         =   "Mask Right Out"
               Alignment       =   8
               PictureAlignment=   11
            End
            Begin Threed.SSCommand CropCommand 
               Height          =   1035
               Index           =   6
               Left            =   60
               TabIndex        =   8
               Top             =   2880
               Width           =   4035
               _ExtentX        =   7117
               _ExtentY        =   1826
               _Version        =   262144
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "SizeSettingsForm.frx":1A7F
               Caption         =   "Set For Selected Size"
               Alignment       =   8
               PictureAlignment=   11
            End
            Begin VB.Label LeftMaskPosition 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   90
               TabIndex        =   10
               Top             =   2130
               Width           =   1965
            End
            Begin VB.Label RightMaskPosition 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   2100
               TabIndex        =   9
               Top             =   2130
               Width           =   1995
            End
         End
      End
      Begin Threed.SSFrame TestImageFrame 
         Height          =   7725
         Left            =   30
         TabIndex        =   14
         Top             =   4320
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   13626
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Select Image"
         Begin VB.DirListBox ImageFilePath 
            Height          =   1215
            Left            =   60
            TabIndex        =   16
            Top             =   570
            Width           =   7425
         End
         Begin VB.DriveListBox ImageFileDrive 
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   210
            Width           =   7455
         End
         Begin FPSpread.vaSpread ImageFileList 
            Height          =   5205
            Left            =   60
            TabIndex        =   17
            Top             =   1830
            Width           =   7425
            _Version        =   393216
            _ExtentX        =   13097
            _ExtentY        =   9181
            _StockProps     =   64
            AutoClipboard   =   0   'False
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridShowHoriz   =   0   'False
            GridSolid       =   0   'False
            MaxCols         =   7
            MaxRows         =   20
            OperationMode   =   3
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "SizeSettingsForm.frx":1D99
            UserResize      =   0
         End
         Begin Threed.SSFrame ViewingOptionsFrame 
            Height          =   615
            Left            =   60
            TabIndex        =   18
            Top             =   7050
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   1085
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Layout Options"
            Begin VB.TextBox FixedMarkerText 
               Height          =   300
               Left            =   780
               TabIndex        =   19
               Top             =   210
               Width           =   6585
            End
            Begin VB.Label Label1 
               Caption         =   "Caption"
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
               Left            =   60
               TabIndex        =   20
               Top             =   210
               Width           =   675
            End
         End
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SizeSettingsToolBars 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   3
      ToolsCount      =   30
      Tools           =   "SizeSettingsForm.frx":22A9
      ToolBars        =   "SizeSettingsForm.frx":26661
   End
   Begin LEADLib.LEAD DigitalCutCode 
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   13200
      Visible         =   0   'False
      Width           =   345
      _Version        =   65540
      _ExtentX        =   609
      _ExtentY        =   609
      _StockProps     =   229
      Enabled         =   0   'False
      ScaleHeight     =   23
      ScaleWidth      =   23
      DataField       =   ""
      BitmapDataPath  =   ""
      AnnDataPath     =   ""
      PanWinTitle     =   "PanWindow"
      CLeadCtrl       =   0
   End
End
Attribute VB_Name = "SizeSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Digital VP-2                                              **
'**                                                                        **
'** Module.....: SizeSettings                                              **
'**                                                                        **
'** Description: User Control to provide Print Size Configuration.         **
'**    This module also provides all image processing functions for the    **
'**    DVP2 Printer Console (Captioning, Cut Code, etc.)                   **
'**                                                                        **
'** History....:                                                           **
'**    09/25/03 v1.00 RDR Implemented Class from existing code.            **
'**                                                                        **
'** (c) 2002-2003 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit                                     'Require explicit variable declaration
Private UpdatingImage As Boolean                    'Flag to prevent recursion in image proc routine

Private Sub Form_Load()
    '---- Initialize Controls
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SizeSettingsToolBars_ToolClick                        **
'**                                                                        **
'**  Description..:  This routine handles adding/erasing print sizes.      **
'**                                                                        **
'****************************************************************************
Private Sub SizeSettingsToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    'MsgBox Tool.ID
    Dim FileName As String, SizeString As String
    Select Case Tool.ID
        Case "ID_Add"
            DB.AddNewSize
        Case "ID_Erase"
            With DB.rsPrintSizes
                If Not .BOF And Not .EOF Then
                    SizeString = .Fields("PrintSize").Value
                    If MsgBox("Delete Print Size [" & SizeString & "] from List?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") = vbYes Then
                        AppLog InfoMsg, "PrintSizeSettingButton_Click,Removed " & Trim(SizeString) & " from print size list."
                        .Delete adAffectCurrent
                        PrintSizeSettings.Refresh ssRefetchAndFireInitializeRow
                    End If
                End If
            End With
        Case "ID_Print"
            With ImageFileList
                .Row = .ActiveRow
                .Col = 1
                FileName = ImageFilePath.Path & IIf(Right(Trim(ImageFilePath.Path), 1) = "\", "", "\") & Trim(.Text)
                .Col = 7
                SizeString = Trim(.Text)
            End With
            DiagnosticsForm.BitON DiagnosticsForm.RightPanShutterBit
            DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, True
            If FileSystemHandle.FileExists(FileName) = True Then
                MotorDiagnostics.PaperMotorTorqueON
                DiagnosticsForm.MakeExposure FileName, SizeString, True, "", "", 1
            End If
            DiagnosticsForm.BitOFF DiagnosticsForm.RightPanShutterBit
            DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, False
        Case "ID_Refresh"
            UpdateImage ImageFileList.ActiveRow
        Case "ID_MotorOn"
            If Tool.State = ssChecked Then
                MotorDiagnostics.PaperMotorTorqueOFF
                Tool.Name = "Motor Off"
            Else
                MotorDiagnostics.PaperMotorTorqueON
                Tool.Name = "Motor On"
            End If
            Sleep 500                                               'Avoid debounce
        Case "ID_AdvancePaper"
            Dim length As Single
            length = DB.PaperAdvanceLength
            With MotorDiagnostics
                .PaperMotorTorqueON
                .AdvancePaper IIf(length < 40, length, 40)           'Advance the paper - no more than 20 inches
                .WaitForPaperAdvance
                Sleep 500
            End With
        Case "ID_RemovePaper"
            With MotorDiagnostics
                If DB.PaperAdvanceTearOffLength < 50 Then           'Check for tear off < 50in.
                    .AdvancePaper DB.PaperAdvanceTearOffLength
                Else
                    AppLog ErrorMsg, "TearOffButton_Click,Tear Off Length too long (>50) " & DB.PaperAdvanceTearOffLength
                    .AdvancePaper 33                                'Provide standard tear off length
                End If
                .WaitForPaperAdvance
                .PaperMotorTorqueOFF
            End With
            If DB.PlatenCylinderEnable = True Then
                DiagnosticsForm.BitON DiagnosticsForm.PlattenCylinderBit
            End If
            DiagnosticsForm.LastSize = ""
        Case "ID_HomeBoth"
            MotorDiagnostics.PaperMaskHome
            LeftMaskPosition.Caption = 0
            RightMaskPosition.Caption = 0
        Case "ID_HomeLeft"
            MotorDiagnostics.HomeLeftMask
            MotorDiagnostics.WaitForPaperMask
            LeftMaskPosition.Caption = 0
        Case "ID_HomeRight"
            MotorDiagnostics.HomeRightMask
            MotorDiagnostics.WaitForPaperMask
            RightMaskPosition.Caption = 0
        Case "ID_Red"
            If Tool.State = ssChecked Then
                DiagnosticsForm.MakeColor 1
            Else
                DiagnosticsForm.AllFiltersOff
            End If
        Case "ID_Green"
            If Tool.State = ssChecked Then
                DiagnosticsForm.MakeColor 2
            Else
                DiagnosticsForm.AllFiltersOff
            End If
        Case "ID_Blue"
            If Tool.State = ssChecked Then
                DiagnosticsForm.MakeColor 3
            Else
                DiagnosticsForm.AllFiltersOff
            End If
        Case "ID_Lamp"
            If Tool.State = ssChecked Then
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Opening Lamp Shutter"
                DiagnosticsForm.BitON DiagnosticsForm.LampShutterBit24
                Sleep 20
                DiagnosticsForm.BitON DiagnosticsForm.LampShutterBit12
                DiagnosticsForm.BitOFF DiagnosticsForm.LampShutterBit24
            Else
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Closing Lamp Shutter"
                DiagnosticsForm.BitOFF DiagnosticsForm.LampShutterBit24
                DiagnosticsForm.BitOFF DiagnosticsForm.LampShutterBit12
            End If
        Case "ID_Iris"
            If Tool.State = ssChecked Then
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Opening Iris Shutter"
                DiagnosticsForm.BitON DiagnosticsForm.IrisShutterBit24
            Else
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Closing Iris Shutter"
                DiagnosticsForm.BitOFF DiagnosticsForm.IrisShutterBit24
            End If
        Case "ID_Pan"
            If Tool.State = ssChecked Then
                DiagnosticsForm.BitON DiagnosticsForm.RightPanShutterBit
                If DiagnosticsForm.IsInputEnabled(DiagnosticsForm.RightPanOpen) = True Then       'Wait for Pan Shutter Switch
                    DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, True
                Else
                    Sleep 1000                                    'Wait for Fixed Delay
                    AppLog DebugMsg, "ExposureSpread_ButtonClicked,Opening Pan Shutter with 1000ms delay."
                End If
            Else
                DiagnosticsForm.BitOFF DiagnosticsForm.RightPanShutterBit
                If DiagnosticsForm.IsInputEnabled(DiagnosticsForm.RightPanOpen) = True Then       'Wait for Pan Shutter Switch
                    DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, False
                Else
                    Sleep 1000                                    'Wait for Fixed Delay
                    AppLog DebugMsg, "ExposureSpread_ButtonClicked,Closing Pan Shutter with 1000ms delay."
                End If
            End If
        Case "ID_Rotate"
            If Tool.State = ssChecked Then
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Rotating table 90 degrees"
                DiagnosticsForm.BitON DiagnosticsForm.RotateBit
            Else
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Rotating table to home"
                DiagnosticsForm.BitOFF DiagnosticsForm.RotateBit
            End If
        Case "ID_11x14"
            If Tool.State = ssChecked Then
                If DB.StepperMaskInstalled = False Then
                    AppLog DebugMsg, "ExposureSpread_ButtonClicked,Deactivating 8x10 for Nord"
                    DiagnosticsForm.BitOFF DiagnosticsForm.Lens8x10Bit
                    Sleep 100
                End If
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Activating 11x14"
                DiagnosticsForm.BitON DiagnosticsForm.Lens11x14Bit
            Else
                AppLog DebugMsg, "ExposureSpread_ButtonClicked,Deactivating 11x14"
                DiagnosticsForm.BitOFF DiagnosticsForm.Lens11x14Bit
                If DB.StepperMaskInstalled = False Then
                    AppLog DebugMsg, "ExposureSpread_ButtonClicked,Activating 8x10 for Nord"
                    Sleep 100
                    DiagnosticsForm.BitON DiagnosticsForm.Lens8x10Bit
                End If
            End If
        Case "ID_LCDRed"
        Case "ID_LCDGreen"
        Case "ID_LCDBlue"
        Case "ID_LCDPlane0"
        Case "ID_LCDPlane1"
        Case "ID_LCDPlane2"
        Case "ID_LCDPlane3"
        Case "ID_FitImage"
            If Tool.State = ssChecked Then
                TargetImage.PaintSizeMode = PAINTSIZEMODE_FITSIDES
                ZoomSlider.Enabled = False
                ZoomSlider.Visible = False
            Else
                TargetImage.AutoSetRects = True
                TargetImage.PaintSizeMode = PAINTSIZEMODE_ZOOM
                ZoomSlider.Value = 100
                TargetImage.PaintZoomFactor = 100
                ZoomSlider.Enabled = True
                ZoomSlider.Visible = True
            End If
        Case "ID_Ruler"
        Case "ID_Focus"
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SizeSettingsToolBars_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CropCommand_Click                                     **
'**                                                                        **
'**  Description..:  This routine handles the Crop Command.                **
'**                                                                        **
'****************************************************************************
Private Sub CropCommand_Click(index As Integer)
    On Error GoTo ErrorHandler
    Select Case index
        Case 0, 1                   'Smaller, Larger   (NO LONGER USED)
        Case 2                      'Mask Left (In)
            If Val(LeftMaskPosition.Caption) - DB.MaskStepsOnClick > DB.MaskLeftLimit Then     'Negative #
                LeftMaskPosition.Caption = LeftMaskPosition.Caption - DB.MaskStepsOnClick
                MotorDiagnostics.MoveLeftMask LeftMaskPosition.Caption
            Else
                MsgBox "Left-mask limit reached (CLOSED).", vbApplicationModal + vbInformation + vbOKOnly, "WARNING"
            End If
            
        Case 3                      'Mask Left (Out)
        
            If Val(LeftMaskPosition.Caption) + DB.MaskStepsOnClick <= 0 Then      'Negative #
                LeftMaskPosition.Caption = LeftMaskPosition.Caption + DB.MaskStepsOnClick
                MotorDiagnostics.MoveLeftMask LeftMaskPosition.Caption
            Else
                MsgBox "Left-mask limit reached (OPEN).", vbApplicationModal + vbInformation + vbOKOnly, "WARNING"
            End If
            
        Case 4                      'Mask Right (In)
        
            If Val(RightMaskPosition.Caption) - DB.MaskStepsOnClick >= 0 Then
                RightMaskPosition.Caption = RightMaskPosition.Caption - DB.MaskStepsOnClick
                MotorDiagnostics.MoveRightMask RightMaskPosition.Caption
            Else
                MsgBox "Right-mask limit reached (CLOSED).", vbApplicationModal + vbInformation + vbOKOnly, "WARNING"
            End If
            
        Case 5                      'Mask Right (Out)
        
            If Val(RightMaskPosition.Caption) + DB.MaskStepsOnClick <= DB.MaskRightLimit Then
                RightMaskPosition.Caption = RightMaskPosition.Caption + DB.MaskStepsOnClick
                MotorDiagnostics.MoveRightMask RightMaskPosition.Caption
            Else
                MsgBox "Right-mask limit reached (OPEN).", vbApplicationModal + vbInformation + vbOKOnly, "WARNING"
            End If
            
        Case 6                      'Set MASK for current exposure
            If MsgBox("Update exposure table with current MASK Settings?", vbApplicationModal + vbQuestion + vbDefaultButton2 + vbYesNo, "Are you sure?") = vbYes Then
                With DB.rsPrintSizes
                    .Fields("LeftMaskPosition").Value = LeftMaskPosition.Caption
                    .Fields("RightMaskPosition").Value = RightMaskPosition.Caption
                End With
            End If
    End Select
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":CropCommand_Click_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Private Sub Form_Resize()
    On Error GoTo ErrorHandler
    If Me.Width - 100 > 0 Then
        Me.PrintSizeSplitter.Width = Me.Width - 100
        Me.PrintSizeSplitter.Height = Me.Height - 100
        
        Me.PrintSizeFrame.Width = Me.PrintSizeSplitter.Panes("Pane A").Width - 20
        Me.PrintSizeSettings.Height = Me.PrintSizeSplitter.Height - 20
        
        Me.PreviewFrame.Width = Me.PrintSizeSplitter.Panes("Pane C").Width - 200
        Me.PreviewFrame.Height = Me.PrintSizeSplitter.Panes("Pane C").Height - 200
        
'        Me.TargetImage.Width = Me.PreviewFrame.Width - 100
        'Me.TargetImage.Height = Me.PreviewFrame.Height - 100
    End If
    Exit Sub
ErrorHandler:
    Resume Next
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes form controls.               **
'**                                                                        **
'****************************************************************************
Public Sub Setup()

    On Error GoTo ErrorHandler
    UpdateSizeSettingsGrid
    
    Me.SizeSettingsToolBars.Tools("ID_FitImage").State = ssChecked
    UpdatingImage = False
    
    '---- Set initial directory for calibration folder (no longer active)
    If FileSystemHandle.FolderExists(DB.CalibrationImagePath) = True Then
        ImageFilePath.Path = DB.CalibrationImagePath                'Set the initial image path to setting in DB
    End If
    
    If DB.StepperMaskInstalled = True Then
        '--- DVP-2 has a stepper mask, show the diagnostics
        PaperMaskFrame.Enabled = True
        PaperMaskFrame.Visible = True
    Else
        'Nord does not have a stepper mask, hide the diagnostics
        PaperMaskFrame.Enabled = False
        PaperMaskFrame.Visible = False
    End If

    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "SizeSettings:Setup,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateSizeSettingsGrid                                **
'**                                                                        **
'**  Description..:  This routine handles updates the print size grid.     **
'**                                                                        **
'****************************************************************************
Public Sub UpdateSizeSettingsGrid()
    On Error GoTo ErrorHandler
    
    '---- The print size settings grid
    With PrintSizeSettings
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
        If DB.StepperMaskInstalled = True Then
            .Bands(0).Columns(6).Header.Caption = "11x14"
        Else
            .Bands(0).Columns(6).Header.Caption = "10x13"
        End If
        .Bands(0).Columns(6).Width = 600
        .Bands(0).Columns(7).Header.Caption = "Video"
        .Bands(0).Columns(7).Width = 700
        
        If DB.StepperMaskInstalled = True Then
            .Bands(0).Columns(8).Header.Caption = "Left Mask"
            .Bands(0).Columns(8).Width = 1100
            .Bands(0).Columns(9).Header.Caption = "Right Mask"
            .Bands(0).Columns(9).Width = 1100
            .Bands(0).Columns(10).Hidden = True
            .Bands(0).Columns(11).Hidden = True
            .Bands(0).Columns(12).Hidden = True
            .Bands(0).Columns(15).Header.Caption = "Punch Feed"
            .Bands(0).Columns(15).Width = 1200
        Else
            .Bands(0).Columns(8).Hidden = True
            .Bands(0).Columns(9).Hidden = True
            .Bands(0).Columns(10).Header.Caption = "8x10 Flap"
            .Bands(0).Columns(10).Width = 1000
            .Bands(0).Columns(11).Header.Caption = "7x10 Flap"
            .Bands(0).Columns(11).Width = 1000
            .Bands(0).Columns(12).Header.Caption = "Right Flap"
            .Bands(0).Columns(12).Width = 1000
            .Bands(0).Columns(15).Hidden = True
        End If
        
        .Bands(0).Columns(13).Header.Caption = "PreFeed"
        .Bands(0).Columns(13).Width = 1000
        .Bands(0).Columns(14).Header.Caption = "Feed"
        .Bands(0).Columns(14).Width = 800
    End With
    Exit Sub
ErrorHandler:
    AppLog ErrorMsg, "UpdateSizeSettingsGrid,Error=" & Err.Number & " Src=" & Err.Source & "Desc=" & Err.Description
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ZoomSlider_Click                                      **
'**                                                                        **
'**  Description..:  This routine handles the Zoom slider control.         **
'**                                                                        **
'****************************************************************************
Private Sub ZoomSlider_Click()
    On Error GoTo ErrorHandler
    TargetImage.PaintZoomFactor = ZoomSlider.Value
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "SizeSettings:ZoomSlider_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ImageFileDrive_Change                                 **
'**                                                                        **
'**  Description..:  This routine handles drive change on size tab.        **
'**                                                                        **
'****************************************************************************
Private Sub ImageFileDrive_Change()
    On Error Resume Next
    ImageFilePath.Path = ImageFileDrive.Drive
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ImageFilePath_Change                                  **
'**                                                                        **
'**  Description..:  This routine handles folder change on size tab.       **
'**                                                                        **
'****************************************************************************
Private Sub ImageFilePath_Change()
    On Error Resume Next
    Dim FolderHandle As Scripting.Folder, TempFileHandle As Scripting.File
    With ImageFileList
        .MaxRows = 0
        '---- Add a row to the spreadsheet for each IMAGE file in the directory
        If FileSystemHandle.FolderExists(ImageFilePath.Path) = True Then
            Set FolderHandle = FileSystemHandle.GetFolder(ImageFilePath.Path)
            '---- Loop for each image file in the folder
            For Each TempFileHandle In FolderHandle.Files
                If (TempFileHandle.Attributes And Directory) = False And (TempFileHandle.Attributes And System) = False Then
                    PrinterConsole.ImagePreview.GetFileInfo Trim(ImageFilePath.Path) & IIf(Right(Trim(ImageFilePath.Path), 1) = "\", "", "\") & Trim(TempFileHandle.Name), 0, FILEINFO_TOTALPAGES
                    If PrinterConsole.ImagePreview.InfoBits >= 7 Then                  'Must be at least 8bpp to be considered an image
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                        .SetActiveCell 1, .Row
                        .SetText 1, .Row, TempFileHandle.Name
                        .SetText 2, .Row, PrinterConsole.ImagePreview.InfoXRes
                        .SetInteger 3, .Row, CLng(PrinterConsole.ImagePreview.InfoHeight)
                        .SetInteger 4, .Row, CLng(PrinterConsole.ImagePreview.InfoWidth)
                        .SetText 5, .Row, Format(CLng(PrinterConsole.ImagePreview.InfoWidth) / CLng(PrinterConsole.ImagePreview.InfoHeight), "####.###")
                        .SetText 6, .Row, PrinterConsole.ImagePreview.InfoSizeDisk
                        .SetText 7, .Row, DB.GetPrintSize(PrinterConsole.ImagePreview.InfoHeight, PrinterConsole.ImagePreview.InfoWidth)
                        .Col = 1
                        PrinterConsole.ImagePreview.Bitmap = 0
                    End If
                End If
            Next
        Else
            AppLog ErrorMsg, "ImageFilePath_Change,Folder " & ImageFilePath.Path & " access error."
        End If
        ImageFileList_LeaveCell 1, 2, 1, 1, False  'Force an update of the current image
    End With
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ImageFileList_LeaveCell                               **
'**                                                                        **
'**  Description..:  This routine handles image file list.                 **
'**                                                                        **
'****************************************************************************
Private Sub ImageFileList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow <> -1 Then
        UpdateImage NewRow
    Else
        ImageFileList.SetActiveCell 1, Row
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdateImage                                           **
'**                                                                        **
'**  Description..:  This routine displays image on lcd with proper mask.  **
'**                                                                        **
'****************************************************************************
Private Sub UpdateImage(ImageRow As Long)
    If UpdatingImage = False Then
        UpdatingImage = True
        
        Dim ColorPlane As Integer, ShiftPlane As Integer, Plane As Integer, FileName As String, SizeString As String
        With ImageFileList
            .SetActiveCell 1, ImageRow
            .Row = ImageRow
            .Col = 1
            FileName = ImageFilePath.Path & IIf(Right(Trim(ImageFilePath.Path), 1) = "\", "", "\") & Trim(.Text)
            .Col = 7
            SizeString = Trim(.Text)
        End With
        
        '---- If the image has a valid size string, the display it
        If SizeString <> "" Then
            If DiagnosticsForm.PrepareToPrintImage(FileName, SizeString, False, 0) <> -1 Then       'This creates CurrentPrintFile if successful
                '---- Now update the LCD with the current image
                'For Plane = 0 To 2
                '    If LcdColorPlane(Plane).Value = True Then
                        ColorPlane = 1
                '        Exit For
                '    End If
                'Next
                'For Plane = 0 To 3
                '    If LcdShiftPlane(Plane).Value = True Then
                        ShiftPlane = 1          'Plane
                '        Exit For
                '    End If
                'Next
                AppLog DebugMsg, "UpdateImage,Outputting frame " & (4 * ColorPlane) + ShiftPlane & " to LCD."
                OutputFrame (4 * ColorPlane) + ShiftPlane      'Output the frame
            End If
            ImageTimeOut = 0
            WatchForImageTimeOut = True
        End If
        UpdatingImage = False
    End If
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MakeDigitalCutCode                                    **
'**                                                                        **
'**  Description..:  This routine creates a digital package punch code.    **
'**                                                                        **
'****************************************************************************
Private Sub MakeDigitalCutCode(CutValue As Byte)
    Dim BitNum As Byte, Row As Integer, Col As Integer, DotLeft As Integer, BitVal As Byte
    Const DotTop As Integer = 10
    Const DotWidth As Integer = 10
    Const DotHeight As Integer = 10
    
    With DigitalCutCode
        .Bitmap = 0                          'Clear old bitmap
        .ReleaseBitmapDC
        .ReleaseClientDC
        .ScaleMode = 3
        .CreateBitmap DotWidth * 21, DotHeight * 3, 24
        .Fill RGB(255, 255, 255)
        .DrawPersistence = True
        .DrawFillColor = RGB(0, 0, 0)
        .DrawFillStyle = DRAWFILLSTYLE_SOLID
        '---- Draw prefix alignment pattern
        DotLeft = DotWidth                  'Start one dot into the bitmap
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        '---- Draw the dots
        For BitNum = 0 To 7
            BitVal = CutValue And (2 ^ BitNum)
            If BitVal = (2 ^ BitNum) Then
                .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
            End If
            DotLeft = DotLeft + DotWidth    'Add for last dot
        Next
        '---- Draw the postfix alignment pattern
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
    End With
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MakeSingleDigitalCutCode                              **
'**                                                                        **
'**  Description..:  This routine creates a digital package punch code.    **
'**                                                                        **
'****************************************************************************
Private Sub MakeSingleDigitalCutCode()
    Dim BitNum As Byte, Row As Integer, Col As Integer, DotLeft As Integer, BitVal As Byte
    Const DotTop As Integer = 0         'Was 10
    Const DotWidth As Integer = 10
    Const DotHeight As Integer = 5
    
    With DigitalCutCode
        .Bitmap = 0                          'Clear old bitmap
        .ReleaseBitmapDC
        .ReleaseClientDC
        .ScaleMode = 3
        .CreateBitmap DotWidth * 21, DotHeight * 2, 24
        .Fill RGB(255, 255, 255)
        .DrawPersistence = True
        .DrawFillColor = RGB(0, 0, 0)
        .DrawFillStyle = DRAWFILLSTYLE_SOLID
        '---- Draw prefix alignment pattern
        DotLeft = DotWidth                  'Start one dot into the bitmap
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
        DotLeft = DotLeft + DotWidth        'Add for last dot
        DotLeft = DotLeft + DotWidth        'Add for last white space between dots
        .DrawRectangle DotLeft, DotTop, DotWidth, DotHeight
    End With
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ProcessImage                                          **
'**                                                                        **
'**  Description..:  This routine processes bitmaps into proper format for **
'**                  outputting to LCD.  This is the ONLY routine that     **
'**                  modifies the bitmap for printing.                     **
'**                                                                        **
'****************************************************************************
Public Function ProcessImage(FileName As String) As Currency
    On Error GoTo ErrorHandler
    Dim VidCropTop As Integer, VidCropLeft As Integer, VidCropRight As Integer, VidCropBottom As Integer, TotalTime As Currency
    Dim mPC As PerformanceCounter
    TotalTime = -1
    Set mPC = New PerformanceCounter
    mPC.StartTimer True
    UpdateProgress.Value = 0
    VidCropBottom = 2400
    VidCropRight = 3200
    
    '----- Image needs to be moved X pixels such that caption is not in image!!
    
    With DB.rsPrintSizes
        '---- Determine image position based on rotation
        If .Fields("RotateTable").Value = True Then
            VidCropTop = .Fields("VideoPosition").Value
            VidCropLeft = 0
            'If .Fields("Caption").Value = True Then
            'Else
            'End If
        Else
            If .Fields("Caption").Value = True Then
                VidCropTop = 50
            Else
                VidCropTop = 0
            End If
            VidCropLeft = .Fields("VideoPosition").Value
        End If
    End With
    
    With PrinterConsole.ImagePreview
        '---- Retrieve fresh copy of image (or create new gray pattern if selected for focus)
        AppLog DebugMsg, "ProcessImage,Loading & processing bitmap image."
        .AutoRepaint = False                            'Turn off auto-repaint, it just slows us down
        .Bitmap = 0                                     'Release the bitmap
        .ReleaseBitmapDC                                'Release handle
        .ReleaseClientDC                                'Release handle
        .DrawPersistence = True                         'Ensure we are drawing on the bitmap - not just display
        .PaintSizeMode = PAINTSIZEMODE_NORMAL           'Set size mode to normal for proper scaling
        .ForceRepaint                                   'MUST force repaint for these to take effect
        .AutoSetRects = True
        .Load FileName, 0, 0, 1                         'This is the ONLY place where an image is loaded!!
        .PaintSizeMode = PAINTSIZEMODE_FIT
        .AutoRepaint = True
    End With
    UpdateProgress.Value = 25
    
    With TargetImage                                    'Copy image to target image (TARGET IS ALWAYS 3200x2400)
        AppLog DebugMsg, "ProcessImage,Combining bitmap images..."
        .AutoRepaint = False                            'Turn off auto-repaint, it just slows us down
        .Bitmap = 0                                     'Release the bitmap
        .ReleaseBitmapDC                                'Release handle
        .ReleaseClientDC                                'Release handle
        .DrawPersistence = True                         'Ensure we are drawing on the bitmap - not just display
        .PaintSizeMode = PAINTSIZEMODE_NORMAL           'Set paintsize mode to normal for scaling
        .ForceRepaint                                   'Must force repaint to obtain proper dimensioning for next functions
        .CreateBitmap 3200, 2400, 24                    'Create a fresh copy
        .Fill RGB(255, 255, 255)                        'Create White White-space!
        .Combine VidCropLeft, VidCropTop, PrinterConsole.ImagePreview.BitmapWidth, PrinterConsole.ImagePreview.BitmapHeight, PrinterConsole.ImagePreview.Bitmap, 0, 0, CB_OP_ADD + CB_DST_0
        UpdateProgress.Value = 50
        
        If SizeSettingsToolBars.Tools("ID_Focus").State = ssChecked Then                'If show focus pattern is checked
            
            .DrawPenColor = RGB(0, 0, 0)                'Set line pen color to black
            .DrawPenStyle = DRAWPENSTYLE_SOLID          'Set solid line style
            .DrawPenWidth = 1                           'Set line width to 1 pixel
            .DrawFillStyle = DRAWFILLSTYLE_TRANSPARENT
            .DrawRectangle 0, 0, 3200, 2400
            .DrawLine 0, 0, 3200, 2400                  'X Pattern
            .DrawLine 0, 2400, 3200, 0                  'X Pattern
            .DrawLine 0, (2400 / 2) + (0 / 2), 3200, (2400 / 2) + (0 / 2) 'H Crosshair
            .DrawLine (3200 / 2) + (0 / 2), 0, (3200 / 2) + (0 / 2), 2400 'V Crosshair
            
            
            'If SizeSettingsToolBars.Tools("ID_Ruler").State = ssChecked Then
            '    Dim xpixel As Integer, ypixel As Integer
            '    For xpixel = VidCropLeft To VidCropRight Step 10
            '        .DrawLine xpixel, VidCropTop, xpixel, IIf(xpixel Mod 100 = 0, VidCropTop + 10, VidCropTop + 5)
            '        .DrawLine xpixel, VidCropBottom, xpixel, IIf(xpixel Mod 100 = 0, VidCropBottom - 10, VidCropBottom - 5)
            '    Next
            '    For ypixel = VidCropTop To VidCropBottom Step 10
            '        .DrawLine VidCropLeft, ypixel, IIf(ypixel Mod 100 = 0, VidCropLeft + 10, VidCropLeft + 5), ypixel
            '        .DrawLine VidCropRight, ypixel, IIf(ypixel Mod 100 = 0, VidCropRight - 10, VidCropRight - 5), ypixel
            '    Next
            'End If
        End If
        
        UpdateProgress.Value = 75
        
        
        '---- Fill Sides of Image with gray to prevent reflections
        If DB.ApplyGrayBorder = True Then
            .DrawFillColor = RGB(64, 64, 64)
            .DrawFillStyle = DRAWFILLSTYLE_SOLID
            If DB.rsPrintSizes.Fields("Lens11x14").Value = True Then
                '---- Left & Right of 11x14 (top & bottom on screen)
                .DrawRectangle 0, 0, VidCropLeft, 2400
                .DrawRectangle VidCropLeft + PrinterConsole.ImagePreview.BitmapWidth, 0, VidCropRight, 2400
            Else
                '---- Top & Bottom of 8x10 (left & right sides on screen)
                .DrawRectangle 0, 0, VidCropLeft, 2400
                .DrawRectangle VidCropLeft + PrinterConsole.ImagePreview.BitmapWidth, 0, VidCropRight, 2400
            End If
        End If
        
        If DB.rsPrintSizes.Fields("Caption").Value = True Then 'Caption the image if necessary
            Dim OutText As String
            OutText = Format(Now, "mm/dd/yyyy") & " " & Format(Now, "hh:mm:ssAMPM") & " " & Trim(FileName)
            If DB.rsPrintSizes.Fields("Lens11x14").Value = False Then
                OutText = OutText & ", Red=" & DB.rsExposureTime.Fields("Red8x10").Value
                OutText = OutText & ", Grn=" & DB.rsExposureTime.Fields("Green8x10").Value
                OutText = OutText & ", Blu=" & DB.rsExposureTime.Fields("Blue8x10").Value
            Else
                OutText = OutText & ", Red=" & DB.rsExposureTime.Fields("Red11x14").Value
                OutText = OutText & ", Grn=" & DB.rsExposureTime.Fields("Green11x14").Value
                OutText = OutText & ", Blu=" & DB.rsExposureTime.Fields("Blue11x14").Value
            End If
            OutText = OutText & ", Crop L=" & VidCropLeft & ", T=" & VidCropTop & ", B=" & VidCropBottom & ", R=" & VidCropRight
            OutText = OutText & IIf(Len(FixedMarkerText.Text) > 0, ", " & Trim(FixedMarkerText.Text), "")
            .Font.Name = "Arial"
            .Font.Size = 24
            .Font.Bold = False
            .Font.Italic = False
            .DrawFontColor = RGB(0, 0, 0)
            If DB.rsPrintSizes.Fields("Lens11x14").Value = True Then
                .TextTop = VidCropTop
                .TextLeft = VidCropLeft + PrinterConsole.ImagePreview.BitmapWidth
                .TextWidth = 100 * 2
                .TextHeight = VidCropBottom
                .TextAngle = 900                    'Angle is 100th of degree
            Else
                .TextTop = IIf(VidCropTop > 0, VidCropTop - 1.5 * .Font.Size, 0)
                .TextLeft = VidCropLeft + 100
                .TextWidth = .BitmapWidth
                .TextHeight = .Font.Size * 2
                .TextAngle = 0
            End If
            .TextAlign = EFX_TEXTALIGN_LEFT_TOP
            AppLog DebugMsg, "ProcessImage,Setting caption text to: L=" & .TextLeft & ",T=" & .TextTop & ",W=" & .TextWidth & _
                ",H=" & .TextHeight & ",Angle=" & .TextAngle & ",Font=" & .Font.Name & ",Size=" & .Font.Size & _
                ",Bold=" & Font.Bold & ",Italic=" & Font.Italic & ",Underline=" & Font.Underline & ",Text=[" & OutText & "]" & ", " & Trim(FixedMarkerText.Text)
            .DrawText OutText, 0
            
        End If
        
        '---- Add Digital Cut Code
        If DB.RenderCutCode = True Then
            Dim DigitalCutValue As Single

            If DB.rsPrintQue.RecordCount > 0 Then
                DigitalCutValue = DB.rsPrintQue.Fields("PunchCode").Value
                If DigitalCutValue = 0 Then DigitalCutValue = 1
            Else
                DigitalCutValue = 1
            End If
            If DigitalCutValue > 0 And DigitalCutValue < 256 Then
                
                If DB.rsPrintSizes.Fields("Lens11x14").Value = True Then
                
                    '---- Digital Package Punch
                    'MakeDigitalCutCode CByte(DigitalCutValue)
                    
                    DigitalCutCode.Rotate -9000, ROTATE_RESIZE, RGB(255, 255, 255)
                    
                    '---- Original
                    '.Combine PrinterConsole.ImagePreview.BitmapWidth, VidCropBottom - DigitalCutCode.BitmapHeight - DB.DigitalCut11x14Offset, DigitalCutCode.BitmapWidth, DigitalCutCode.BitmapHeight, DigitalCutCode.Bitmap, 0, 0, CB_OP_ADD + CB_DST_0
                    
                    '---- New for Package Cut Code
                    .Combine VidCropLeft, PrinterConsole.ImagePreview.BitmapHeight - DigitalCutCode.BitmapHeight - DB.DigitalCut11x14Offset, DigitalCutCode.BitmapWidth, DigitalCutCode.BitmapHeight, DigitalCutCode.Bitmap, 0, 0, CB_OP_ADD + CB_DST_0
                    
                    DigitalCutCode.Rotate 9000, ROTATE_RESIZE, RGB(255, 255, 255)       'Rotate back for next use!
                    
                Else
                    
                    '---- Digital Package Punch
                    'MakeDigitalCutCode CByte(DigitalCutValue)
                    
                    MakeSingleDigitalCutCode
                    DigitalCutCode.Reverse
                    
                    '----- Original
                    '.Combine VidCropLeft + PrinterConsole.ImagePreview.BitmapWidth - DigitalCutCode.BitmapWidth - DB.DigitalCut8x10Offset, IIf(VidCropTop > 0, VidCropTop - DigitalCutCode.BitmapHeight, 0), DigitalCutCode.BitmapWidth, DigitalCutCode.BitmapHeight, DigitalCutCode.Bitmap, 0, 0, CB_OP_ADD + CB_DST_0
                    
                    '----- New for Package Cut Code
                    .Combine VidCropLeft + PrinterConsole.ImagePreview.BitmapWidth - DigitalCutCode.BitmapWidth - DB.DigitalCut8x10Offset, PrinterConsole.ImagePreview.BitmapHeight - DigitalCutCode.BitmapHeight, DigitalCutCode.BitmapWidth, DigitalCutCode.BitmapHeight, DigitalCutCode.Bitmap, 0, 0, CB_OP_ADD + CB_DST_0
                    
                End If
            End If
        End If
        
        .AutoRepaint = True
        .ForceRepaint
        .Save CurrentPrintFile, FILE_BMP, 0, 0, 0       'Save the image to display on the LCD
    End With
    
    SizeSettingsToolBars.Tools("ID_FitImage").State = ssChecked
    TargetImage.PaintSizeMode = PAINTSIZEMODE_FITSIDES
    ZoomSlider.Enabled = False
    ZoomSlider.Visible = False
        
    
    UpdateProgress.Value = 100
    TotalTime = mPC.StopTimer
    AppLog DebugMsg, "ProcessImage,Timed," & Format(TotalTime, "####.####")
    Set mPC = Nothing
    ProcessImage = TotalTime
    Exit Function
ErrorHandler:
    mPC.StopTimer
    Set mPC = Nothing
End Function

Public Sub DisableSetupSettings()
    SizeSettingsToolBars.Tools("ID_Focus").State = ssUnchecked
    SizeSettingsToolBars.Tools("ID_Ruler").State = ssUnchecked
    FixedMarkerText.Text = ""
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ClearImage                                            **
'**                                                                        **
'**  Description..:  This routine clears the image on the LCD with Gray.   **
'**                                                                        **
'****************************************************************************
Public Sub ClearImage()
    
    '----- DONT KNOW WHY THIS IS HERE!!! CHECK INTO THIS!!
    'DiagnosticsForm.MakeExposure SettingsFolder & "Images\DVP2_LUT_Calibration.psd", "Calibration", True
    
    If DiagnosticsForm.PrepareToPrintImage(SettingsFolder & "Images\DVP2_Gray.psd", "Calibration", False, 0) <> -1 Then
        AppLog DebugMsg, "ClearImage,Outputting clear image to LCD"
        OutputFrame 0
        ImageTimeOut = 0
        WatchForImageTimeOut = False
    End If
End Sub

