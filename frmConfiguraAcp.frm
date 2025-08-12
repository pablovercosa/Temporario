VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ACTBAR.OCX"
Begin VB.Form frmConfiguraAcp 
   Caption         =   "Configurações Acp"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "frmConfiguraAcp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5280
   Begin VB.TextBox txtProxyAdress 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CheckBox chkProxy 
      Caption         =   "Usar um servidor Proxy para a rede local"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtDte 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CheckBox chkSenha 
      Caption         =   "Mascarar a senha"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblAdress 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmConfiguraAcp.frx":05CA
   End
   Begin VB.Label Label3 
      Caption         =   "DTE"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Senha"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Usuário"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmConfiguraAcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
 Select Case Tool.Name
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
 End Select
 
 End Sub

Private Sub chkProxy_Click()

If chkProxy.Value = 0 Then
   txtProxyAdress.Enabled = False
   lblAdress.Enabled = False
Else
   txtProxyAdress.Enabled = True
   lblAdress.Enabled = True
End If
End Sub

Private Sub chkSenha_Click()

If chkSenha.Value = 1 Then
   txtSenha.PasswordChar = "*"
Else
   txtSenha.PasswordChar = ""
End If

End Sub
Private Sub Form_Load()
  
  Dim sRet As String
 
 
 Call CenterForm(Me)
 If chkSenha.Value = 1 Then txtSenha.PasswordChar = "*"
 
 sRet = GetSetting("QuickStore", "ACP", "Dte", "")
 txtDte.Text = sRet
 
 sRet = GetSetting("QuickStore", "ACP", "User", "")
 txtUser.Text = sRet
 
 sRet = GetSetting("QuickStore", "ACP", "Senha", "")
 txtSenha.Text = sRet
 
 sRet = GetSetting("QuickStore", "ACP", "Adress", "")
 txtProxyAdress.Text = sRet
  
 sRet = GetSetting("QuickStore", "ACP", "Proxy", "")
 If sRet = "True" Then
    chkProxy.Value = 1
 Else
    chkProxy.Value = 0
 End If
  
 
 If chkProxy.Value = 0 Then
    txtProxyAdress.Enabled = False
    lblAdress.Enabled = False
 End If
 
 
End Sub
Private Sub UpdateRecord()
Dim bProxy As Boolean

If chkProxy.Value = 0 Then
   bProxy = False
Else
   bProxy = True
End If

SaveSetting "QuickStore", "ACP", "Dte", txtDte.Text
SaveSetting "QuickStore", "ACP", "User", txtUser.Text
SaveSetting "QuickStore", "ACP", "Senha", txtSenha.Text
SaveSetting "QuickStore", "ACP", "Proxy", bProxy
SaveSetting "QuickStore", "ACP", "Adress", txtProxyAdress.Text

End Sub

Private Sub ClearScreen()

txtDte.Text = ""
txtUser.Text = ""
txtSenha.Text = ""

txtProxyAdress.Text = ""
chkProxy.Value = 0





End Sub
