VERSION 5.00
Begin VB.Form frmMensagensIncluirRegra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incluir Mensagem para Nota Fiscal"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7590
   Icon            =   "frmMensagensIncluirRegra.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMensagem 
      Height          =   615
      Left            =   120
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   4440
      Width           =   7335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   5280
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   7275
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   360
      Width           =   7335
      Begin VB.Frame fraUF 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   480
         TabIndex        =   25
         Top             =   3000
         Width           =   6615
         Begin VB.ComboBox cboUF 
            Height          =   315
            Left            =   2160
            TabIndex        =   15
            Top             =   210
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optFiltroUF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Específico"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optFiltroUF 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame fraOpSaida 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   480
         TabIndex        =   24
         Top             =   1920
         Width           =   6615
         Begin VB.OptionButton optFiltroOpSaidas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Específica"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optFiltroOpSaidas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Grupo Fiscal"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optFiltroOpSaidas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cllque aqui para selecionar o Grupo Fiscal"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   2160
            MouseIcon       =   "frmMensagensIncluirRegra.frx":058A
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Tag             =   "Cllque aqui para selecionar o Grupo Fiscal"
            Top             =   270
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cllque aqui para selecionar a Operação de Saída"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   2160
            MouseIcon       =   "frmMensagensIncluirRegra.frx":0E54
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Tag             =   "Cllque aqui para selecionar a Operação de Saída"
            Top             =   510
            Visible         =   0   'False
            Width           =   3495
         End
      End
      Begin VB.Frame fraProdutos 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Width           =   6615
         Begin VB.OptionButton optFiltroProdutos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Específico"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optFiltroProdutos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Classe e/ou Sub Classe"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   3
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton optFiltroProdutos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Grupo Fiscal"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optFiltroProdutos 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cllque aqui para selecionar a Sub Classe"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   2160
            MouseIcon       =   "frmMensagensIncluirRegra.frx":171E
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Tag             =   "Cllque aqui para selecionar a Sub Classe"
            Top             =   750
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cllque aqui para selecionar o Grupo Fiscal"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   2160
            MouseIcon       =   "frmMensagensIncluirRegra.frx":1FE8
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Tag             =   "Cllque aqui para selecionar o Grupo Fiscal"
            Top             =   270
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cllque aqui para selecionar a Classe"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   2160
            MouseIcon       =   "frmMensagensIncluirRegra.frx":28B2
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Tag             =   "Cllque aqui para selecionar a Classe"
            Top             =   510
            Visible         =   0   'False
            Width           =   2550
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cllque aqui para selecionar o Produto"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   2160
            MouseIcon       =   "frmMensagensIncluirRegra.frx":317C
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Tag             =   "Cllque aqui para selecionar o Produto"
            Top             =   990
            Visible         =   0   'False
            Width           =   2640
         End
      End
      Begin VB.Label lblTitleRule 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "3) Filtro para Estado (UF) do Cliente"
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
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   3075
      End
      Begin VB.Label lblTitleRule 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "2) Filtro para Operações de Saída"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lblTitleRule 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "1) Filtro para Produtos"
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
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Width           =   1905
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
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
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Regra para a mensagem"
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
      TabIndex        =   26
      Top             =   120
      Width           =   2070
   End
End
Attribute VB_Name = "frmMensagensIncluirRegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'27/01/2006 - mpdea
'Tela para inclusão de Mensagens e sua regras de aplicação


Private m_objMensagemNotaFiscal As clsMensagemNotaFiscal

'Variável auxiliar para retorno da opção de filtro selecionada
'
'0 = Produtos - Grupo Fiscal
'1 = Produtos - Classe
'2 = Produtos - Sub Classe
'3 = Produtos - Específico
'4 = Op. Saída - Grupo Fiscal
'5 = Op. Saída - Específica
Private m_strFiltroTemp(5) As String
'Constantes de identificação do Link para Filtro (LF)
Private Const LF_PROD_GRUPO_FISCAL As Integer = 0
Private Const LF_PROD_CLASSE As Integer = 1
Private Const LF_PROD_SUBCLASSE As Integer = 2
Private Const LF_PROD_ESPECIFICO As Integer = 3
Private Const LF_OPSAIDA_GRUPO_FISCAL As Integer = 4
Private Const LF_OPSAIDA_ESPECIFICA As Integer = 5

'01/02/2006 - mpdea
'Retorno do botão OK
Private m_blnOK As Boolean

'Obtém a Mensagem com a Regra
Public Function GetMensagemNotaFiscal() As clsMensagemNotaFiscal
  Set m_objMensagemNotaFiscal = New clsMensagemNotaFiscal
  
  'Padrão
  With m_objMensagemNotaFiscal
    .TipoFiltroProduto = tfpTodos
    .TipoFiltroOpSaida = tfoTodas
    .TipoFiltroUF = tfuTodos
  End With
  m_blnOK = False
  
  'Zera filtro auxiliar
  Erase m_strFiltroTemp
  
  'Carrega e exibe tela
  Load Me
  Me.Show vbModal
  
  If Not m_blnOK Then
    Set m_objMensagemNotaFiscal = Nothing
  End If
  
  Set GetMensagemNotaFiscal = m_objMensagemNotaFiscal
End Function

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  'Validação
  If m_blnValidarDados Then
    With m_objMensagemNotaFiscal
      'Ordem
      .Ordem = 0
      '1) Filtro para Produtos
      Select Case .TipoFiltroProduto
        Case tfpTodos
          .FiltroProduto = "*"
        Case tfpGrupoFiscal
          .FiltroProduto = m_strFiltroTemp(LF_PROD_GRUPO_FISCAL)
        Case tfpClasseSubClasse
          .FiltroProduto = m_strFiltroTemp(LF_PROD_CLASSE) & "|" & _
            m_strFiltroTemp(LF_PROD_SUBCLASSE)
        Case tfpEspecifico
          .FiltroProduto = m_strFiltroTemp(LF_PROD_ESPECIFICO)
      End Select
      '2) Filtro para Operações de Saída
      Select Case .TipoFiltroOpSaida
        Case tfoTodas
          .FiltroOpSaida = "*"
        Case tfoGrupoFiscal
          .FiltroOpSaida = m_strFiltroTemp(LF_OPSAIDA_GRUPO_FISCAL)
        Case tfoEspecifica
          .FiltroOpSaida = m_strFiltroTemp(LF_OPSAIDA_ESPECIFICA)
      End Select
      '3) Filtro para Estado (UF) do Cliente
      Select Case .TipoFiltroUF
        Case tfuTodos
          .FiltroUF = "*"
        Case tfuEspecifico
          .FiltroUF = cboUF.Text
      End Select
      'Mensagem
      .Mensagem = txtMensagem.Text
    End With
    
    'Descarrega form
    m_blnOK = True
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  
  On Error GoTo ErrHandler

  Call StatusMsg("")
  
  Call CenterForm(Me)
    
  Exit Sub
  
ErrHandler:
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
    
End Sub

Private Sub LoadEstados()
  Dim rstEstados As Recordset
  
  On Error GoTo ErrHandler
  
  'Preenche Estados
  cboUF.Clear
  Set rstEstados = db.OpenRecordset("SELECT Estado FROM Estados ORDER BY Estado", dbOpenDynaset, dbReadOnly)
  With rstEstados
    If Not (.BOF And .EOF) Then
      Do While Not .EOF
        cboUF.AddItem .Fields("Estado").Value & ""
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rstEstados = Nothing
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstEstados Is Nothing Then
    rstEstados.Close
    Set rstEstados = Nothing
  End If
  'Exibe mensagem de erro
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub lblLink_Click(Index As Integer)
  Dim enuTipoFiltro As g_enuTipoFiltro
  Dim strCodigo As String
  Dim strNome As String
  
  
  On Error GoTo ErrHandler
  
  
  'Tipo de item
  Select Case Index
    Case 0, 4
      enuTipoFiltro = mtfGrupoFiscal
    Case 1
      enuTipoFiltro = mtfClasse
    Case 2
      enuTipoFiltro = mtfSubClasse
    Case 3
      enuTipoFiltro = mtfProduto
    Case 5
      enuTipoFiltro = mtfOpSaida
  End Select
  
  'Obtém o item
  If frmMensagensFiltroSelecao.GetItem(enuTipoFiltro, strCodigo, strNome) Then
    If strCodigo = "" Then
      'Restaura texto se retornar sem informação
      lblLink(Index).Caption = lblLink(Index).Tag
      Exit Sub
    Else
      'Exibe item selecionado
      lblLink(Index).Caption = "Item selecionado: " & strCodigo & " - " & strNome
      'Armazena item selecionado
      m_strFiltroTemp(Index) = strCodigo
    End If
  End If
  
  Exit Sub

ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub optFiltroUF_Click(Index As Integer)
  'Carrega Estados
  If Index = 1 And optFiltroUF(1).Tag <> "1" Then
    Call LoadEstados
    'Marca como carregado a combo de estados
    optFiltroUF(1).Tag = "1"
  End If
  'Seta tipo selecionado
  m_objMensagemNotaFiscal.TipoFiltroUF = Index
  'Exibição dos objetos
  cboUF.Visible = Index = 1 'Estado
End Sub

Private Sub optFiltroOpSaidas_Click(Index As Integer)
  'Seta tipo selecionado
  m_objMensagemNotaFiscal.TipoFiltroOpSaida = Index
  'Exibição dos objetos
  lblLink(4).Visible = Index = 1 'Grupo Fiscal
  lblLink(5).Visible = Index = 2 'Operação de Saída
End Sub

Private Sub optFiltroProdutos_Click(Index As Integer)
  'Seta tipo selecionado
  m_objMensagemNotaFiscal.TipoFiltroProduto = Index
  'Exibição dos objetos
  lblLink(0).Visible = Index = 1 'Grupo Fiscal
  lblLink(1).Visible = Index = 2 'Classe
  lblLink(2).Visible = Index = 2 'Sub Classe
  lblLink(3).Visible = Index = 3 'Produto
End Sub

Private Function m_blnValidarDados() As Boolean
  Dim intX As Integer
  Dim intIndex As Integer
  
  
  On Error GoTo ErrHandler
    
  
  '1) Filtro para Produtos
  For intX = optFiltroProdutos.LBound To optFiltroProdutos.UBound
    If optFiltroProdutos(intX).Value Then
      Select Case intX
        Case 0 'Todos
          'OK
          Exit For
        Case 1 'Grupo Fiscal
          If m_strFiltroTemp(LF_PROD_GRUPO_FISCAL) = "" Then
            DisplayMsg "Selecione o Grupo Fiscal para filtro dos Produtos."
            optFiltroProdutos(intX).SetFocus
            Exit Function
          End If
        Case 2 'Classe e function Classe
          If m_strFiltroTemp(LF_PROD_CLASSE) = "" And _
            m_strFiltroTemp(LF_PROD_SUBCLASSE) = "" Then
            DisplayMsg "Selecione a Classe e/ou Sub-Classe para filtro dos Produtos."
            optFiltroProdutos(intX).SetFocus
            Exit Function
          End If
        Case 3 'Específico
          If m_strFiltroTemp(LF_PROD_ESPECIFICO) = "" Then
            DisplayMsg "Selecione o Produto para filtro dos Produtos."
            optFiltroProdutos(intX).SetFocus
            Exit Function
          End If
      End Select
    End If
  Next intX
  
  '2) Filtro para Operações de Saída
  For intX = optFiltroOpSaidas.LBound To optFiltroOpSaidas.UBound
    If optFiltroOpSaidas(intX).Value Then
      Select Case intX
        Case 0 'Todos
          'OK
          Exit For
        Case 1 'Grupo Fiscal
          If m_strFiltroTemp(LF_OPSAIDA_GRUPO_FISCAL) = "" Then
            DisplayMsg "Selecione o Grupo Fiscal para filtro das Operações de Saída."
            optFiltroOpSaidas(intX).SetFocus
            Exit Function
          End If
        Case 2 'Específica
          If m_strFiltroTemp(LF_OPSAIDA_ESPECIFICA) = "" Then
            DisplayMsg "Selecione a Operação de Saída para filtro das Operações de Saída."
            optFiltroOpSaidas(intX).SetFocus
            Exit Function
          End If
      End Select
    End If
  Next intX
  
  '3) Filtro para Estado (UF) do Cliente
  For intX = optFiltroUF.LBound To optFiltroUF.UBound
    If optFiltroUF(intX).Value Then
      Select Case intX
        Case 0 'Todos
          'OK
          Exit For
        Case 1 'Específico
          If cboUF.Text = "" Then
            DisplayMsg "Selecione o Estado (UF) para filtro dos Estados (UF) dos Clientes."
            cboUF.SetFocus
            Exit Function
          End If
      End Select
    End If
  Next intX
  
  'Mensagem
  If Trim(txtMensagem.Text) = "" Then
    DisplayMsg "Informe a mensagem a ser utilizada."
    txtMensagem.SetFocus
    Exit Function
  End If
  
    
  'Retorna resultado
  m_blnValidarDados = True
  
  Exit Function

ErrHandler:
  m_blnValidarDados = False
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function
