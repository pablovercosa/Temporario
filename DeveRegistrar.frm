VERSION 5.00
Begin VB.Form frmDeveRegistrar 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cópia não registrada."
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "DeveRegistrar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Registrar depois"
      Height          =   400
      Left            =   2205
      TabIndex        =   3
      Top             =   2820
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Default         =   -1  'True
      Height          =   400
      Left            =   4110
      TabIndex        =   2
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   120
      Picture         =   "DeveRegistrar.frx":058A
      Top             =   180
      Width           =   2745
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"DeveRegistrar.frx":CE54
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1800
      Left            =   3000
      TabIndex        =   1
      Top             =   135
      Width           =   4500
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Usuários REGISTRADOS poderão usufruir de promoções especiais nas próximas versões do software e em outros produtos Infopar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   150
      TabIndex        =   0
      Top             =   1905
      Width           =   7245
   End
End
Attribute VB_Name = "frmDeveRegistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gsPrefix As String

Private Sub cmdCancel_Click()
  gbProdutoRegistrado = False
  Unload Me
End Sub

Private Sub cmdRegistrar_Click()
  Call gbConsoleLicencas(gsPrefix)
  Unload Me
End Sub

Private Sub Form_Activate()
  cmdRegistrar.SetFocus
End Sub

