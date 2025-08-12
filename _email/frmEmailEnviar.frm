VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEmailEnviar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Enviar Email"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailEnviar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5550
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEnviar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enviar"
      Height          =   435
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   8295
   End
   Begin VB.Frame fraAssunto 
      Caption         =   "Assunto"
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   1320
      Width           =   8295
      Begin VB.TextBox txtAssunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame fraMensagem 
      Caption         =   "Mensagem"
      Height          =   2835
      Left            =   60
      TabIndex        =   7
      Top             =   2100
      Width           =   8295
      Begin VB.TextBox txtMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   2475
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   270
         Width           =   7815
      End
   End
   Begin VB.Frame fraHeader 
      Caption         =   "Para"
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8295
      Begin VB.TextBox txtDestinatarioEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtDestinatarioNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmEmailEnviar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'30/01/2009 - mpdea
'Tela para envio de e-mail

Private WithEvents SendMail As vbSendMail.clsSendMail
Attribute SendMail.VB_VarHelpID = -1

Public Sub LoadEmail(ByVal strDestinatarioNome As String, ByVal strDestinatarioEmail As String, _
  ByVal strAssunto As String, ByVal strMensagem As String)
  
  txtDestinatarioNome.Text = strDestinatarioNome
  txtDestinatarioEmail.Text = strDestinatarioEmail
  txtAssunto.Text = strAssunto
  txtMensagem.Text = strMensagem
End Sub

Private Sub cmdEnviar_Click()
  Dim obj_config_envio_email As ConfigEnvioEmail

  On Error GoTo ErrHandler

  cmdEnviar.Enabled = False
  stbStatus.SimpleText = ""
  Screen.MousePointer = vbHourglass

  obj_config_envio_email = LoadConfigEnvioEmail(gnCodFilial)

  Set SendMail = New clsSendMail
  With SendMail
    .SMTPHost = obj_config_envio_email.ServidorSmtp
    .POP3Host = obj_config_envio_email.ServidorPop3
    .UseAuthentication = obj_config_envio_email.Autenticacao
    .UsePopAuthentication = obj_config_envio_email.AutenticacaoPop3
    .UserName = obj_config_envio_email.usuario
    .Password = obj_config_envio_email.Senha
    .From = obj_config_envio_email.EmailRemetente
    .FromDisplayName = obj_config_envio_email.NomeExibicaoRemetente
    .Recipient = txtDestinatarioEmail.Text
    .RecipientDisplayName = txtDestinatarioNome.Text
    .Subject = txtAssunto.Text
    .Message = txtMensagem.Text
    .Connect
    .Send
  End With
  Set SendMail = Nothing

  cmdEnviar.Enabled = True
  stbStatus.SimpleText = ""
  Screen.MousePointer = vbDefault
  
  
  Exit Sub
  
ErrHandler:
  Set SendMail = Nothing
  cmdEnviar.Enabled = True
  stbStatus.SimpleText = ""
  Screen.MousePointer = vbDefault
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub SendMail_SendFailed(Explanation As String)
  If Len(Explanation) > 2 Then
    MsgBox "Não foi possível enviar o e-mail: " & Explanation, vbExclamation, "Atenção"
  End If
End Sub

Private Sub SendMail_SendSuccesful()
  Unload Me
End Sub

Private Sub SendMail_Status(Status As String)
  stbStatus.SimpleText = Status
End Sub
