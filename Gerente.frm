VERSION 5.00
Begin VB.Form frmGerente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Senha do Gerente"
   ClientHeight    =   3420
   ClientLeft      =   4170
   ClientTop       =   3750
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Gerente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   6645
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   -15
      TabIndex        =   4
      Top             =   1035
      Width           =   6675
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "SENHA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5400
         TabIndex        =   6
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Caso a senha do gerente ainda não foi cadastrada na tela de [Parâmetros da Filial/Empresa], então digite a palavra"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   435
         TabIndex        =   5
         Top             =   240
         Width           =   5850
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   -75
      TabIndex        =   2
      Top             =   -75
      Width           =   6735
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Para o uso desta função é necessária a digitação da senha do gerente."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1815
         TabIndex        =   3
         Top             =   270
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   570
         Picture         =   "Gerente.frx":4E95A
         Top             =   315
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   6675
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "•"
      TabIndex        =   0
      Text            =   "12345678"
      Top             =   2205
      Width           =   2850
   End
End
Attribute VB_Name = "frmGerente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gbPWS As String
Private gbPressOK As Boolean

Public Function gbSenhaGerente() As Boolean
  
  '22/01/2003 - mpdea
  'Em modo limitado não solicita senha
  If (gsSenhaGerente <> "") And gblnQuickFull Then
    gbSenhaGerente = False
    txtSenha.Text = ""
    Me.Show vbModal
    If gbPressOK Then
      If gbPWS = gsSenhaGerente Then
        gbSenhaGerente = True
      Else
        MsgBox "Senha não confere!", vbExclamation, "Atenção"
      End If
    End If
  Else
    gbSenhaGerente = True
  End If
  
End Function

Private Sub cmdOK_Click()
  gbPressOK = True
  gbPWS = txtSenha.Text
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  gbPressOK = False
  gbPWS = ""
  Call CenterForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmGerente = Nothing
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    Unload Me
  End If
End Sub
