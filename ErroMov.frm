VERSION 5.00
Begin VB.Form frmErroMov 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4395
   ClientLeft      =   1140
   ClientTop       =   1230
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "ErroMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "&Não"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Sim"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblAtention 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ATENÇÃO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Label lblAtention 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ATENÇÃO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ErroMov.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Infopar Quick Store"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmErroMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gbResult As Boolean

Public Function gbContinue() As Boolean
  Screen.MousePointer = vbDefault
  Me.Show vbModal
  gbContinue = gbResult
End Function

Private Sub cmdNo_Click()
  gbResult = False
  Unload Me
End Sub

Private Sub cmdYes_Click()
  gbResult = True
  Unload Me
End Sub
