VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSenhaFuncionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senha requerida"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
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
   Icon            =   "frmSenhaFuncionario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   45
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3105
      Width           =   6555
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   470
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2565
      Width           =   6555
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   75
      PasswordChar    =   "•"
      TabIndex        =   0
      Text            =   "12345678"
      Top             =   2040
      Width           =   2820
   End
   Begin Threed.SSPanel sspUserCode 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   1320
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "0"
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.01
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
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
      Left            =   -120
      TabIndex        =   3
      Top             =   -120
      Width           =   6735
      Begin VB.Image imgSenha 
         Height          =   480
         Left            =   570
         Picture         =   "frmSenhaFuncionario.frx":4E95A
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Informe a senha do funcionário abaixo para prosseguir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1635
         TabIndex        =   4
         Top             =   360
         Width           =   3735
      End
   End
   Begin Threed.SSPanel sspUserName 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   " ABC"
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.01
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   1
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   75
      TabIndex        =   9
      Top             =   1800
      Width           =   510
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   1080
      Width           =   555
   End
End
Attribute VB_Name = "frmSenhaFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'29/12/2003 - mpdea
'
'Objeto    : Form
'Nome      : frmSenhaFuncionario
'
'Descrição : Utilizado para exigir a senha do funcionário sempre que necessário
'
'Uso       : Executar a função CheckSenha, informando o código do funcionário
'            A função retornará True se a senha for informada corretamente

Private m_blnReturnOK As Boolean
Private m_dblValorP As Double

Public Function CheckSenha(ByVal intUserCode As Integer) As Boolean
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Nome, ValorP FROM Funcionários WHERE Código = " & intUserCode
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If (.BOF And .EOF) Then
      sspUserName.Caption = " Usuário inexistente"
      sspUserName.Font.Bold = True
      cmdOK.Enabled = False
    Else
      sspUserName.Caption = " " & .Fields("Nome").Value & ""
      Call IsDataType(dtDouble, .Fields("ValorP").Value, m_dblValorP)
    End If
    .Close
  End With
  Set rs = Nothing
  
  m_blnReturnOK = False
  txtSenha.Text = ""
  sspUserCode.Caption = intUserCode
  
  Me.Show vbModal
  
  CheckSenha = m_blnReturnOK
  
End Function

Private Sub cmdCancel_Click()
  m_blnReturnOK = False
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
  If CriptografaSenha(txtSenha.Text) = m_dblValorP Then
    m_blnReturnOK = True
    Unload Me
  Else
    MsgBox "Senha não confere.", vbExclamation, "Atenção"
    Call SelectAllText(txtSenha, True)
  End If

End Sub
