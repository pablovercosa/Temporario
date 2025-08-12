VERSION 5.00
Begin VB.Form frmMotivoCancelamento 
   Caption         =   " Motivo do Cancelamento"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMotivoCancelamento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNao 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Não"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1470
      Width           =   4815
   End
   Begin VB.TextBox txtMotivoCancelamento 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   480
      Width           =   10695
   End
   Begin VB.CommandButton cmdSim 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sim"
      Default         =   -1  'True
      Height          =   375
      Left            =   510
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1470
      Width           =   4815
   End
   Begin VB.Label lblConfirma 
      Caption         =   "Confirma Canelamento?"
      Height          =   255
      Left            =   540
      TabIndex        =   4
      Top             =   1110
      Width           =   2655
   End
   Begin VB.Label lblMotivoCancelamento 
      Caption         =   "Informe o motivo do cancelamento (minimo 15 caracteres)"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMotivoCancelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNao_Click()
gsRetornoDoc = "NÂO"
Unload Me
End Sub

Private Sub cmdSim_Click()
  If Len(txtMotivoCancelamento.Text) >= 15 Then
    strMotivoCancelamento = txtMotivoCancelamento.Text
    gsRetornoDoc = "OK"
    Unload Me
  Else
    MsgBox "O motivo de cancelamento precisa ter no minimo 15 caracteres"
    gsRetornoDoc = "NÂO"
  End If
End Sub
