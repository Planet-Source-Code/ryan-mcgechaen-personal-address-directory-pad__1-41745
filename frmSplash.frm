VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Interval        =   5000
      Left            =   1800
      Top             =   960
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Address Directory"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label lblvers 
      BackStyle       =   0  'Transparent
      Caption         =   "v2.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Ryan McGechaen, Copywrite 2002"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   2040
      Width           =   3375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrStart_Timer()
    Unload Me
    frmMain.Show
End Sub
