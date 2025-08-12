VERSION 5.00
Begin VB.Form frmComenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Comentários"
   ClientHeight    =   4410
   ClientLeft      =   975
   ClientTop       =   3900
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Comenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4410
   ScaleWidth      =   10125
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3795
      Width           =   9885
   End
   Begin VB.TextBox Comenta 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      Height          =   3600
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   75
      Width           =   9915
   End
End
Attribute VB_Name = "frmComenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  gsBuffer = Comenta.Text
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub
