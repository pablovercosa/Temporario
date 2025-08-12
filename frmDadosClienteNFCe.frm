VERSION 5.00
Begin VB.Form frmDadosClienteNFCe 
   Caption         =   " Dados do Cliente"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDadosClienteNFCe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   7170
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   6915
   End
   Begin VB.CommandButton btnCancela 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3015
      Width           =   6885
   End
   Begin VB.CommandButton btnOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2430
      Width           =   6885
   End
   Begin VB.TextBox txtIE 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   1830
      Width           =   6915
   End
   Begin VB.TextBox txtCpf_Cnpj 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6915
   End
   Begin VB.Label lblNome 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lblIE 
      BackStyle       =   0  'Transparent
      Caption         =   "Inscrição Estadual (Somente Numeros)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1590
      Width           =   2895
   End
   Begin VB.Label lblCPF_CPNJ 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ (Somente números)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2325
   End
End
Attribute VB_Name = "frmDadosClienteNFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancela_Click()
  gsRetornoDoc = "NÃO"
  Unload Me
End Sub

Private Sub btnOk_Click()
  gsNomeCliente = txtNome.Text
  gsCPF_Cnpj = txtCpf_Cnpj.Text
  gsIE = txtIE.Text
  gsRetornoDoc = "OK"
  Unload Me
End Sub

Private Sub Form_Load()
  gsRetornoDoc = ""
  txtNome.Text = gsNomeCliente
  txtCpf_Cnpj.Text = gsCPF_Cnpj
  txtIE.Text = gsIE
End Sub
