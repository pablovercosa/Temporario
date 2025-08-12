VERSION 5.00
Begin VB.Form frmTrocaSenhaGerente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Troca de Senha do Gerente"
   ClientHeight    =   2910
   ClientLeft      =   1665
   ClientTop       =   2760
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TrocaSenhaGerente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   6795
   Begin VB.TextBox txtConfirm 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2730
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1350
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
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
      Height          =   435
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2370
      Width           =   6375
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
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
      Height          =   435
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1830
      Width           =   6375
   End
   Begin VB.TextBox txtSenhaGerente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2730
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   892
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Entre com um valor alfanumérico de até 8 caracteres. Maiúsculas e minúsculas SÃO DIFERENTES. Não use a expressão SENHA."
      Height          =   525
      Left            =   150
      TabIndex        =   6
      Top             =   195
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "Confirmação"
      Height          =   225
      Left            =   1590
      TabIndex        =   5
      Top             =   1410
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Nova Senha"
      Height          =   225
      Left            =   1590
      TabIndex        =   4
      Top             =   960
      Width           =   960
   End
End
Attribute VB_Name = "frmTrocaSenhaGerente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
  If txtSenhaGerente.Text = "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Senha do Gerente incorreta."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    txtSenhaGerente.SetFocus
    Exit Sub
  End If
  
  If Len(Trim(txtConfirm.Text)) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Confirmação de Senha incorreta."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    txtConfirm.SetFocus
    Exit Sub
  End If
  
  If txtSenhaGerente.Text <> txtConfirm.Text Then
    gsTitle = LoadResString(201)
    gsMsg = "Confirmação e Senha não conferem. Reentre."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    txtSenhaGerente.SetFocus
    Exit Sub
  End If
 
  frmParametros.gsSenhaGerenteAtual = txtSenhaGerente.Text
  
  DisplayMsg "Agora clique no ícone SALVAR para concluir a troca da senha do gerente."
  
  Unload Me
 
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

