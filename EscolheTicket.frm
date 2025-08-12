VERSION 5.00
Begin VB.Form frmEscolheTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolha o modelo de ticket a imprimir"
   ClientHeight    =   4170
   ClientLeft      =   4500
   ClientTop       =   2280
   ClientWidth     =   6675
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EscolheTicket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6675
   Begin VB.CommandButton B_Cancelar 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3630
      Width           =   6540
   End
   Begin VB.CommandButton B_OK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   6540
   End
   Begin VB.ListBox Lista1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   6540
   End
   Begin VB.Label Retorno 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   6060
      TabIndex        =   3
      Top             =   3810
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "frmEscolheTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Cancelar_Click()
  gsRetornoDoc = "CANCELADO"
  Unload Me
End Sub

Private Sub B_OK_Click()
  If Lista1.ListIndex = -1 Then
    gsRetornoDoc = "CANCELADO"
  Else
    gsRetornoDoc = Lista1.List(Lista1.ListIndex)
  End If
  Unload Me
End Sub

Private Sub Form_Activate()
  Dim Texto As String
  Dim Aux_Str As String
  Dim Fim As Integer
  
  Lista1.Clear
  
  Rem Enche Combo_Tickets
  Aux_Str = gsConfigPath & "*.CTI"
  Fim = False
  Texto = Dir(Aux_Str)
  If Texto = "" Then Exit Sub
  Lista1.AddItem Texto
  Do
    Texto = Dir
    If Texto = "" Then Fim = True
    If Fim = False Then Lista1.AddItem Texto
  Loop Until Fim = True
 
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub
