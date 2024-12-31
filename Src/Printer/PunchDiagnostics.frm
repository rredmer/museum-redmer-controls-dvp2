VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.Form PunchDiagnostics 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars PunchToolBars 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      ToolBarsCount   =   1
      ToolsCount      =   1
      Tools           =   "PunchDiagnostics.frx":0000
      ToolBars        =   "PunchDiagnostics.frx":0CEE
   End
   Begin VB.TextBox PunchCodeTest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      Text            =   "1"
      Top             =   1170
      Width           =   495
   End
   Begin Threed.SSCommand PunchButton 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   1170
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
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
      Caption         =   "Test Punch"
      Alignment       =   1
   End
End
Attribute VB_Name = "PunchDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Setup()
    '---- Initialize Punch Settings
    
End Sub

Private Sub PunchButton_Click()
    DiagnosticsForm.MakePunch CByte(Val(PunchCodeTest.Text))
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PunchToolBars_ToolClick                               **
'**                                                                        **
'**  Description..:  This routine handles the Punch Toolbar.               **
'**                                                                        **
'****************************************************************************
Private Sub PunchToolBars_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Select Case Tool.ID
        Case "ID_TestPunch"
            DiagnosticsForm.MakePunch CByte(Val(PunchCodeTest.Text))
    End Select
End Sub
