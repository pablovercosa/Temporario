VERSION 5.00
Begin VB.Form frmEmailConfigurar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Configuração para envio de Email"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailConfigurar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4230
      Width           =   9735
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3690
      Width           =   9735
   End
   Begin VB.Frame fraPadrao 
      Caption         =   "Padrão"
      Height          =   1125
      Left            =   60
      TabIndex        =   12
      Top             =   2430
      Width           =   9735
      Begin VB.TextBox txtEmailRemetente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4950
         MaxLength       =   255
         TabIndex        =   16
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtNomeExibicaoRemetente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         MaxLength       =   255
         TabIndex        =   14
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Endereço de email do remetente"
         Height          =   195
         Index           =   5
         Left            =   4950
         TabIndex        =   15
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Nome para exibição do remetente"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2430
      End
   End
   Begin VB.Frame fraAutenticacao 
      Caption         =   "Autenticação"
      Height          =   1125
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   9735
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   7620
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   450
         Width           =   1845
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4950
         MaxLength       =   255
         TabIndex        =   10
         Top             =   450
         Width           =   2595
      End
      Begin VB.CheckBox chkAutenticacaoPop3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Requer autenticação POP3"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2190
         TabIndex        =   7
         Top             =   510
         Width           =   2355
      End
      Begin VB.CheckBox chkAutenticacao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Requer autenticação"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "(Opcional)"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   3630
         TabIndex        =   19
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Index           =   3
         Left            =   7620
         TabIndex        =   9
         Top             =   210
         Width           =   450
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Usuário"
         Height          =   195
         Index           =   2
         Left            =   4950
         TabIndex        =   8
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame fraServidor 
      Caption         =   "Servidor"
      Height          =   1125
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   9735
      Begin VB.TextBox txtServidorPop3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4920
         MaxLength       =   255
         TabIndex        =   4
         Top             =   570
         Width           =   4575
      End
      Begin VB.TextBox txtServidorSmtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   210
         MaxLength       =   255
         TabIndex        =   2
         Top             =   570
         Width           =   4575
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "POP3 (Opcional)"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   3
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "SMTP"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmEmailConfigurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'30/01/2009 - mpdea
'Tela para configuração de envio de e-mail

Private m_int_codigo_filial As Integer

Public Property Let CodigoFilial(ByVal Value As Integer)
  m_int_codigo_filial = Value
End Property

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim rstEmail As Recordset
  Dim strSQL As String
  
  
  On Error GoTo ErrHandler
  
  
  strSQL = "SELECT * FROM Email WHERE Filial = " & m_int_codigo_filial
  Set rstEmail = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rstEmail
    If Not (.BOF And .EOF) Then
      .Edit
    Else
      .AddNew
      .Fields("Filial").Value = m_int_codigo_filial
    End If
    .Fields("ServidorSmtp").Value = txtServidorSmtp.Text
    .Fields("ServidorPop3").Value = txtServidorPop3.Text
    .Fields("Autenticacao").Value = IIf(chkAutenticacao.Value = vbChecked, True, False)
    .Fields("AutenticacaoPop3").Value = IIf(chkAutenticacaoPop3.Value = vbChecked, True, False)
    .Fields("Usuario").Value = txtUsuario.Text
    .Fields("Senha").Value = txtSenha.Text
    .Fields("NomeExibicaoRemetente").Value = txtNomeExibicaoRemetente.Text
    .Fields("EmailRemetente").Value = txtEmailRemetente.Text
    .Update
    .Close
  End With
  Set rstEmail = Nothing
  Unload Me
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstEmail Is Nothing Then
    rstEmail.Close
    Set rstEmail = Nothing
  End If
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub Form_Load()
  Dim obj_config_envio_email As ConfigEnvioEmail

  On Error GoTo ErrHandler
  
  obj_config_envio_email = LoadConfigEnvioEmail(m_int_codigo_filial)
  With obj_config_envio_email
    txtServidorSmtp.Text = .ServidorSmtp
    txtServidorPop3.Text = .ServidorPop3
    chkAutenticacao.Value = IIf(.Autenticacao, vbChecked, vbUnchecked)
    chkAutenticacaoPop3.Value = IIf(.AutenticacaoPop3, vbChecked, vbUnchecked)
    txtUsuario.Text = .usuario
    txtSenha.Text = .Senha
    txtNomeExibicaoRemetente.Text = .NomeExibicaoRemetente
    txtEmailRemetente.Text = .EmailRemetente
  End With
  
  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub
