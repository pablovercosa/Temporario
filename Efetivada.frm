VERSION 5.00
Begin VB.Form frmEfetivada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ATENÇÃO"
   ClientHeight    =   2505
   ClientLeft      =   2985
   ClientTop       =   2835
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Efetivada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Efetivada.frx":4E95A
   ScaleHeight     =   2505
   ScaleWidth      =   6690
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1860
      Width           =   5955
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"Efetivada.frx":4EEE4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   450
      TabIndex        =   0
      Top             =   195
      Width           =   5850
   End
End
Attribute VB_Name = "frmEfetivada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub


Private Sub Form_Load()
  Call CenterForm(Me)
End Sub
