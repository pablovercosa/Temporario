VERSION 5.00
Begin VB.Form frmErro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ocorreu um erro em seu aplicativo"
   ClientHeight    =   3735
   ClientLeft      =   2010
   ClientTop       =   1920
   ClientWidth     =   6630
   ControlBox      =   0   'False
   HelpContextID   =   50
   Icon            =   "frmErro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturnErro 
      Caption         =   "&Sair"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturnErro 
      Caption         =   "&Encerrar"
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturnErro 
      Caption         =   "&Prosseguir"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturnErro 
      Caption         =   "&Repetir"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Image imgErro 
      Height          =   480
      Left            =   240
      Picture         =   "frmErro.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblModulo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblErrNumber 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Módulo"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Número do erro"
      Height          =   195
      Left            =   4920
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblTitulo 
      Caption         =   $"frmErro.frx":044E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'30/04/2003 - mpdea
'Ocultado botão Prosseguir, impedindo que o usuário possa prosseguir
'a operação com erro

Private gnReturn As Integer

Public Function gnShowErr(ByVal nErro As Integer, ByVal sModulo As String) As Integer
  lblModulo.Caption = sModulo
  lblErrNumber.Caption = nErro
  txtDescription.Text = Error(nErro)
  Me.Show vbModal
  gnShowErr = gnReturn
End Function

Private Sub cmdReturnErro_Click(Index As Integer)
  gnReturn = Index
  Unload Me
End Sub
