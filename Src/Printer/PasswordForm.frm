VERSION 5.00
Begin VB.Form PasswordForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Mode"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2220
      TabIndex        =   2
      Top             =   840
      Width           =   825
   End
   Begin VB.TextBox Password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2220
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   270
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   2115
   End
End
Attribute VB_Name = "PasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Password.Text = ""
End Sub

Private Sub OK_Click()
    Me.Hide
End Sub
