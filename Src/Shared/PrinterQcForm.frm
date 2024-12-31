VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form PrinterQcForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame CurrentLUTFrame 
      Caption         =   "Current LUT Setting"
      Height          =   795
      Left            =   300
      TabIndex        =   1
      Top             =   1170
      Width           =   4515
      Begin VB.Label LUTLabel 
         Caption         =   "LUTLabel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4365
      End
   End
   Begin ActiveToolBars.SSActiveToolBars QcToolBars 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "PrinterQcForm.frx":0000
      ToolBars        =   "PrinterQcForm.frx":3F86
   End
   Begin Threed.SSCheck UserMode 
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
      Caption         =   "Q.C. Mode"
   End
End
Attribute VB_Name = "PrinterQcForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: DVP2                                                      **
'**                                                                        **
'** Module.....: PrinterQCForm                                             **
'**                                                                        **
'** Description: This form provides quality control functions.             **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Public PrintingLUT As Boolean                          'Set TRUE if printing LUT image (linear LUT)

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Setup                                                 **
'**                                                                        **
'**  Description..:  This routine initializes form variables & controls.   **
'**                                                                        **
'****************************************************************************
Public Sub Setup()
    PrintingLUT = False                                 'Indicates if printing LUT file
End Sub

Private Sub Form_Activate()
    If DB.ApplyMullerSohnLUT = True Then
        Me.LUTLabel.Caption = "48-Step MuellerSOHN"
    Else
        Me.LUTLabel.Caption = "72-Step PictoGraphics"
    End If
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  PrintCalibrationImage_Click                           **
'**                                                                        **
'**  Description..:  This routine prints calibration images.               **
'**                                                                        **
'****************************************************************************
Private Sub QcToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    On Error GoTo ErrorHandler
    Dim ColorPlane As Integer, ShiftPlane As Integer, Plane As Integer, FileName As String, SizeString As String
    MainForm.StatusLabel(0).Caption = "Printing calibration..."
    With DB.rsPrintSizes
        If .RecordCount = 0 Then
            MsgBox "No print sizes defined"
            Me.Enabled = True
            Exit Sub
        End If
        .MoveFirst
        .Find "PrintSize='Calibration'"
        If Not .EOF Then
            .Fields("RotateTable").Value = -1
            .Fields("Lens11x14").Value = 0
            .UpdateBatch adAffectCurrent
        Else
            MsgBox "Calibration print size not defined"
            Me.Enabled = True
            Exit Sub
        End If
        DiagnosticsForm.BitON DiagnosticsForm.RightPanShutterBit
        DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, True
        
        Select Case Tool.ID
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
            Case "ID_PrintDailyTarget"
                .MoveFirst
                .Find "PrintSize='Calibration'"
                If Not .EOF Then
                    .Fields("Caption").Value = -1
                    .Fields("RotateTable").Value = -1
                    .Fields("Lens11x14").Value = 0
                    .UpdateBatch adAffectCurrent
                Else
                    MsgBox "Calibration print size not defined"
                    Me.Enabled = True
                    Exit Sub
                End If
                DiagnosticsForm.MakeExposure SettingsFolder & "Images\DVP2_Target_Calibration.psd", "Calibration", True, "", "", 1
                .MoveFirst
                .Find "PrintSize='Calibration'"
                If Not .EOF Then
                    .Fields("Caption").Value = -1
                    .Fields("RotateTable").Value = 0
                    .Fields("Lens11x14").Value = -1
                    .UpdateBatch adAffectCurrent
                Else
                    MsgBox "Calibration print size not defined"
                    Me.Enabled = True
                    Exit Sub
                End If
                DiagnosticsForm.MakeExposure SettingsFolder & "Images\DVP2_Target_Calibration.psd", "Calibration", True, "", "", 1
            
            Case "ID_PrintOffset"
                PrintingLUT = True
                .Fields("Caption").Value = 0
                DiagnosticsForm.MakeExposure SettingsFolder & "Images\DVP2_Offset_Calibration.psd", "Calibration", True, "", "", 1
                PrintingLUT = False
            Case "ID_PrintLUT"
                If DB.ApplyMullerSohnLUT = True Then
                    PrintingLUT = True
                    .Fields("Caption").Value = -1
                    DiagnosticsForm.MakeExposure SettingsFolder & "Images\DVP2_LUT_Calibration.psd", "Calibration", True, "", "", 1
                    PrintingLUT = False
                Else
                    PrintingLUT = False             'Use goof LUT when printing picto!
                    .Fields("Caption").Value = 0
                    DiagnosticsForm.MakeExposure SettingsFolder & "Images\DVP2_LUT_Picto_Calibration.psd", "Calibration", True, "", "", 1
                    PrintingLUT = False
                End If
            Case "ID_PrintRingAround"
                FileName = SettingsFolder & "Images\DVP2_RingAround_Calibration.psd"
                If DB.StepperMaskInstalled = True Then
                    SizeString = "3x10"
                Else
                    SizeString = "7x10"
                End If
                If FileSystemHandle.FileExists(FileName) = True Then
                    DiagnosticsForm.MakeExposure FileName, SizeString, True, "", "", 1      'This is normal reference
                    PrintColorRing "Red8x10", FileName, SizeString                          'Red Over-Under
                    PrintColorRing "Green8x10", FileName, SizeString                        'Green Over-Under
                    PrintColorRing "Blue8x10", FileName, SizeString                         'Blue Over-Under
                End If
        End Select
    End With
    DiagnosticsForm.BitOFF DiagnosticsForm.RightPanShutterBit
    DiagnosticsForm.CheckForInput DiagnosticsForm.RightPanOpen, False
    SizeSettingsForm.ClearImage
    MainForm.StatusLabel(0).Caption = PrinterIdleMessage
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrinterQcForm:QcToolBars_ToolClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PrintColorRing                                        **
'**                                                                        **
'**  Description..:  This routine prints over/under for color ring.        **
'**                                                                        **
'****************************************************************************
Private Sub PrintColorRing(ColorField As String, FileName As String, SizeString As String)
    On Error GoTo ErrorHandler
    Dim CurrentColorValue As Long
    CurrentColorValue = CLng(DB.rsExposureTime.Fields(ColorField).Value)
    DB.rsExposureTime.Fields(ColorField).Value = CLng(CurrentColorValue * 1.32)
    DB.rsExposureTime.UpdateBatch adAffectCurrent
    DiagnosticsForm.MakeExposure FileName, SizeString, True, "", "", 1
    DB.rsExposureTime.Fields(ColorField).Value = CLng(CurrentColorValue / 1.32)
    DB.rsExposureTime.UpdateBatch adAffectCurrent
    DiagnosticsForm.MakeExposure FileName, SizeString, True, "", "", 1
    DB.rsExposureTime.Fields(ColorField).Value = CurrentColorValue
    DB.rsExposureTime.UpdateBatch adAffectCurrent
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrinterQcForm:PrintColorRing", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
