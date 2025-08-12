VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAguardeAtualizacaoClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Store"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
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
   Icon            =   "frmAguardeAtualizacaoClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblAviso 
      Caption         =   "Aguarde, nos próximos minutos o sistema estará automaticamente atualizando a classificação de seus Clientes..."
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmAguardeAtualizacaoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  pgbProgress.Value = 0
End Sub
