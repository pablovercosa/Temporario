VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmMensagensFiltroSelecao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione o item"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8055
   Icon            =   "frmMensagensFiltroSelecao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtItemNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5535
   End
   Begin VB.Data datItens 
      Caption         =   "datItens"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo cboItem 
      Bindings        =   "frmMensagensFiltroSelecao.frx":058A
      DataSource      =   "datItens"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2175
      DataFieldList   =   "Código"
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3810
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Código"
      Columns(0).FieldLen=   256
      Columns(1).Width=   9155
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Nome"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmMensagensFiltroSelecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'27/01/2006 - mpdea
'Tela para seleção do valor a ser utilizado no filtro
'da tela de inclusão de regras para mensagem


'Tipos de seleção para o filtro
Public Enum g_enuTipoFiltro
  mtfGrupoFiscal = 1
  mtfClasse
  mtfSubClasse
  mtfProduto
  mtfOpSaida
End Enum

'Tipo de filtro selecionado
Private m_enuTipoFiltroSelec As g_enuTipoFiltro

'Item selecionado
Private m_strRetCodigo As String
Private m_strRetNome As String

'Obtém o item
Public Function GetItem(ByVal enuTipoFiltro As g_enuTipoFiltro, _
  ByRef strCodigo As String, ByRef strNome As String) As Boolean
  
  m_strRetCodigo = ""
  m_strRetNome = ""
  m_enuTipoFiltroSelec = enuTipoFiltro
  Load Me
  Me.Show vbModal
  strCodigo = m_strRetCodigo
  strNome = m_strRetNome
  GetItem = Not (m_strRetCodigo = "")
End Function

Private Sub cboItem_CloseUp()
  cboItem.Text = cboItem.Columns(0).Text
  Call cboItem_LostFocus
End Sub

Private Sub cboItem_LostFocus()
  Dim intItem As Integer
  
  
  On Error GoTo ErrHandler
  
  
  txtItemNome.Text = ""
  
  If cboItem.Text <> "" Then
    
    Select Case m_enuTipoFiltroSelec
      Case mtfGrupoFiscal, mtfClasse, mtfSubClasse, mtfOpSaida
        If Not IsDataType(dtInteger, cboItem.Text, intItem) Then
          DisplayMsg "Código inválido."
          cboItem.Text = ""
          Exit Sub
        End If
        
        If m_enuTipoFiltroSelec = mtfOpSaida Then
          If intItem < 1 Or intItem > 999 Then
            DisplayMsg "Código inválido."
            cboItem.Text = ""
            Exit Sub
          End If
        Else
          If intItem < 1 Or intItem > 9999 Then
            DisplayMsg "Código inválido."
            cboItem.Text = ""
            Exit Sub
          End If
        End If
        
        With datItens.Recordset
          .FindFirst "Código = " & intItem
          If Not .NoMatch Then
            txtItemNome.Text = .Fields("Nome").Value & ""
          End If
        End With
        
      Case mtfProduto
        With datItens.Recordset
          .FindFirst "Código = '" & cboItem.Text & "'"
          If Not .NoMatch Then
            txtItemNome.Text = .Fields("Nome").Value & ""
          End If
        End With
        
    End Select
    
  End If
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmdCancelar_Click()
  m_strRetCodigo = ""
  m_strRetNome = ""
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If cboItem.Text = "" Then
    DisplayMsg "Escolha um item."
    cboItem.SetFocus
  Else
    Call cboItem_LostFocus
    If txtItemNome.Text <> "" Then
      m_strRetCodigo = cboItem.Text
      m_strRetNome = txtItemNome.Text
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim strSQL As String
  Const TB As String = "TABLE_NAME"
  
  On Error GoTo ErrHandler
  
  
  Call StatusMsg("")
  
  Call CenterForm(Me)
  
  'Seta Banco de dados para controle Data
  datItens.DatabaseName = gsQuickDBFileName
  
  'SQL auxiliar
  strSQL = "SELECT Código, Nome FROM TABLE_NAME ORDER BY Nome;"
  
  'Tipo de filtro selecionado
  Select Case m_enuTipoFiltroSelec
    Case mtfGrupoFiscal
      Me.Caption = "Grupo Fiscal"
      datItens.RecordSource = Replace(strSQL, TB, "GrupoFiscal")
    Case mtfClasse
      Me.Caption = "Classe"
      datItens.RecordSource = Replace(strSQL, TB, "Classes")
    Case mtfSubClasse
      Me.Caption = "Sub Classe"
      datItens.RecordSource = Replace(strSQL, TB, "[Sub Classes]")
    Case mtfProduto
      Me.Caption = "Produto"
      datItens.RecordSource = Replace(strSQL, TB, "Produtos")
    Case mtfOpSaida
      Me.Caption = "Operação de Saída"
      datItens.RecordSource = Replace(strSQL, TB, "[Operações Saída]")
  End Select
  
  datItens.Refresh
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub
