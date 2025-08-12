VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmVendaParaFuncionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda para Funcionário"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
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
   Icon            =   "frmVendaParaFuncionario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5115
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Caso necessite retornar para a tela de venda rápida antes de avançar tecle em Cancelar ou pressione {ESC}."
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Utilize os campos abaixo para identificar o funcionário caso seja uma venda para funcionário. Caso contrário deixe em branco."
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdAvançar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Avançar >>"
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
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame fraN 
      Caption         =   "Identificação do Funcionário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4935
      Begin VB.TextBox txtNomeFuncionario 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   3615
      End
      Begin VB.Data datFuncionarios 
         Caption         =   "datFuncionarios"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome, Apelido FROM Funcionários ORDER BY Código"
         Top             =   240
         Visible         =   0   'False
         Width           =   1380
      End
      Begin SSDataWidgets_B.SSDBCombo cboFuncionario 
         Bindings        =   "frmVendaParaFuncionario.frx":000C
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   885
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   1561
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   65535
         DataFieldToDisplay=   "Código"
      End
   End
End
Attribute VB_Name = "frmVendaParaFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFuncionario_CloseUp()
  cboFuncionario.Text = cboFuncionario.Columns(0).Text
  cboFuncionario_LostFocus
End Sub

Private Sub cboFuncionario_LostFocus()
  Dim rstFuncionarios As Recordset
  
  txtNomeFuncionario.Text = ""
  
  If Not IsNumeric(cboFuncionario.Text) Then Exit Sub
  
  Set rstFuncionarios = db.OpenRecordset("SELECT Código, Nome, Apelido FROM Funcionários WHERE Código = " & CInt(cboFuncionario.Text), dbOpenDynaset)

  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      txtNomeFuncionario.Text = .Fields("Nome") & ""
    End If
  End With

  rstFuncionarios.Close
  Set rstFuncionarios = Nothing

End Sub

Private Sub cmdAvançar_Click()
  '18/01/2006 - mpdea
  'Alterado objeto frmVendaRap2 -> g_frmVendaRapida
  'Adicionado With
  With g_frmVendaRapida
    If Len(txtNomeFuncionario.Text) > 0 Then
      .g_intCodigoFuncComprador = CInt(cboFuncionario.Text)
    Else
      .g_intCodigoFuncComprador = 0
    End If
    .g_blnStatusVendaFunc = False
  End With
  
  Unload Me
End Sub

Private Sub cmdCancelar_Click()
  '18/01/2006 - mpdea
  'Alterado objeto frmVendaRap2 -> g_frmVendaRapida
  g_frmVendaRapida.g_blnRetornar = True
  Unload Me
End Sub

Private Sub Form_Load()
  
  datFuncionarios.DatabaseName = gsQuickDBFileName
  
  Call CenterForm(Me)
  
End Sub
