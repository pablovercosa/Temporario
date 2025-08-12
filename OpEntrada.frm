VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmOpEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"OpEntrada.frx":0000
   ClientHeight    =   8070
   ClientLeft      =   2640
   ClientTop       =   1545
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1220
   Icon            =   "OpEntrada.frx":00B2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8070
   ScaleWidth      =   11475
   Begin VB.TextBox txtCSOSN 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      MaxLength       =   3
      TabIndex        =   43
      ToolTipText     =   "Obrigatório somente no caso de se encaixar como Simples Nacional Ex: 101 ou 102 ou 400"
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impostos"
      Height          =   1335
      Left            =   90
      TabIndex        =   26
      Top             =   4500
      Width           =   11235
      Begin VB.CheckBox O_ICMS_IPI 
         Appearance      =   0  'Flat
         Caption         =   "Incidir ICMS sobre IPI"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6420
         TabIndex        =   32
         Top             =   900
         Width           =   2100
      End
      Begin VB.CheckBox O_Base_ICMS_Frete 
         Appearance      =   0  'Flat
         Caption         =   "Considerar total da nota + valor do frete para ICMS"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   4350
      End
      Begin VB.CheckBox IPI_TOT 
         Appearance      =   0  'Flat
         Caption         =   "Calcula IPI somente p/Total"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2850
         TabIndex        =   31
         Top             =   900
         Width           =   2745
      End
      Begin VB.CheckBox O_Base_IPI 
         Appearance      =   0  'Flat
         Caption         =   "Base ICMS deve somar IPI"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   300
         Width           =   2415
      End
      Begin VB.CheckBox IPI 
         Appearance      =   0  'Flat
         Caption         =   "Calcula IPI"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   900
         Width           =   1095
      End
      Begin VB.CheckBox ICM 
         Appearance      =   0  'Flat
         Caption         =   "Calcula ICM (Nao Esta Visivel Para Compatibilidade)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6420
         TabIndex        =   28
         Top             =   300
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin VB.Frame fraReceita 
      Caption         =   "Receita Estadual"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6030
      TabIndex        =   41
      Top             =   6840
      Width           =   5295
      Begin VB.CheckBox chkInformante 
         Appearance      =   0  'Flat
         Caption         =   "Informante próprio (P)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   510
         TabIndex        =   42
         Top             =   285
         Width           =   2685
      End
   End
   Begin VB.Frame fraP 
      Caption         =   "Permitir alteração no preço de Venda"
      Height          =   615
      Left            =   90
      TabIndex        =   37
      Top             =   6840
      Width           =   5895
      Begin VB.CheckBox chkPermitir 
         Appearance      =   0  'Flat
         Caption         =   "Sim"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   255
         Width           =   615
      End
      Begin VB.Data datTabela 
         Caption         =   "datTabela"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4620
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Tabela FROM [Tabela de Preços] ORDER BY Tabela"
         Top             =   150
         Visible         =   0   'False
         Width           =   1380
      End
      Begin SSDataWidgets_B.SSDBCombo cboTabela 
         Bindings        =   "OpEntrada.frx":4EA0C
         Height          =   315
         Left            =   2250
         TabIndex        =   40
         Top             =   240
         Width           =   3285
         DataFieldList   =   "Tabela"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorOdd    =   8454143
         Columns(0).Width=   3200
         _ExtentX        =   5794
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Tabela"
      End
      Begin VB.Label lblTabela 
         AutoSize        =   -1  'True
         Caption         =   "Tabela"
         Height          =   195
         Left            =   1680
         TabIndex        =   38
         Top             =   255
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimentação"
      Height          =   3600
      Left            =   90
      TabIndex        =   5
      Top             =   870
      Width           =   8325
      Begin VB.TextBox txtModeloDocumentoFiscal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E5E5E5&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   7410
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1080
         Width           =   765
      End
      Begin VB.CheckBox chkPrecoCustoCalculado 
         Appearance      =   0  'Flat
         Caption         =   "Calcular CUSTO conforme campo ""Preço de Custo Calculado"" (Pasta cálculos do Produto)"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   210
         TabIndex        =   17
         Top             =   2940
         Width           =   7275
      End
      Begin VB.CheckBox chkGravaCustoPrecoListaSemIPI 
         Appearance      =   0  'Flat
         Caption         =   "Gravar CUSTO no campo ""Preço de lista sem IPI"" (Pasta cálculos do Produto)"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   210
         TabIndex        =   13
         Top             =   1635
         Width           =   6495
      End
      Begin VB.CheckBox chkEmitirNFManualmente 
         Appearance      =   0  'Flat
         Caption         =   "Emitir nota fiscal manualmente (Para situações em que o sistema ficou fora do ar por algum tempo)"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   210
         TabIndex        =   15
         Top             =   2250
         Width           =   7995
      End
      Begin VB.CheckBox chkEstorno 
         Appearance      =   0  'Flat
         Caption         =   "Realiza Estorno da Reserva"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   16
         Top             =   2625
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox Senha 
         Appearance      =   0  'Flat
         Caption         =   "Exige senha do gerente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   1950
         Width           =   2535
      End
      Begin VB.CheckBox Comissão 
         Appearance      =   0  'Flat
         Caption         =   "Diminui a comissão do vendedor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   990
         Width           =   2955
      End
      Begin VB.CheckBox Grava_Custo 
         Appearance      =   0  'Flat
         Caption         =   "Gravar preço na tabela CUSTO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   1305
         Width           =   2895
      End
      Begin VB.CheckBox Nota 
         Caption         =   "Emite nota (Nao Esta Visivel Para Compatibilidade)"
         Height          =   255
         Left            =   7800
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.CheckBox Dinheiro 
         Appearance      =   0  'Flat
         Caption         =   "Movimenta dinheiro (sai/entra no caixa ou contas a pagar)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   675
         Width           =   4845
      End
      Begin VB.CheckBox Estoque 
         Appearance      =   0  'Flat
         Caption         =   "Aumenta estoque"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Modelo documento fiscal"
         Height          =   195
         Index           =   0
         Left            =   5370
         TabIndex        =   8
         Top             =   1125
         Width           =   1755
      End
   End
   Begin VB.Frame Frete 
      Caption         =   "Frete - para operações de compra"
      Height          =   915
      Left            =   90
      TabIndex        =   33
      Top             =   5880
      Width           =   11235
      Begin VB.CheckBox chkSomarFreteCustoProduto 
         Appearance      =   0  'Flat
         Caption         =   "Somar frete no custo dos produtos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6420
         TabIndex        =   35
         Top             =   240
         Width           =   3135
      End
      Begin VB.CheckBox Somar_Frete_Custo 
         Appearance      =   0  'Flat
         Caption         =   "Ao gravar preço de CUSTO somar o frete ao valor dos produtos. (Nao Esta Visivel Para Compatibilidade)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   570
         Visible         =   0   'False
         Width           =   9615
      End
      Begin VB.CheckBox Somar_Frete_Total 
         Appearance      =   0  'Flat
         Caption         =   "Somar frete no total a pagar (contas, etc.)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   270
         Width           =   3585
      End
   End
   Begin VB.TextBox Código_Fiscal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1845
      MaxLength       =   4
      TabIndex        =   4
      ToolTipText     =   "Código de operação fiscal  EX: 5102 para venda estadual"
      Top             =   480
      Width           =   1635
   End
   Begin VB.TextBox Nome 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1845
      MaxLength       =   50
      TabIndex        =   2
      Top             =   105
      Width           =   9495
   End
   Begin VB.TextBox Código 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   735
      MaxLength       =   3
      TabIndex        =   1
      Top             =   105
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Classificação"
      Height          =   3585
      Left            =   8490
      TabIndex        =   18
      Top             =   870
      Width           =   2835
      Begin VB.OptionButton O_Devolução 
         Appearance      =   0  'Flat
         Caption         =   "&Devolução"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   24
         Top             =   2370
         Width           =   1575
      End
      Begin VB.OptionButton O_Pedido 
         Appearance      =   0  'Flat
         Caption         =   "&Pedido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   1950
         Width           =   855
      End
      Begin VB.OptionButton O_Empréstimo 
         Appearance      =   0  'Flat
         Caption         =   "&Recebimento de Empréstimo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   25
         Top             =   2790
         Width           =   2535
      End
      Begin VB.OptionButton O_Grátis_Entrada 
         Appearance      =   0  'Flat
         Caption         =   "&Grátis Entrada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   22
         Top             =   1530
         Width           =   1455
      End
      Begin VB.OptionButton O_Ajuste_Entrada 
         Appearance      =   0  'Flat
         Caption         =   "&Ajuste de Entrada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   21
         Top             =   1125
         Width           =   1695
      End
      Begin VB.OptionButton O_Trans_Entrada 
         Appearance      =   0  'Flat
         Caption         =   "&Transferência de Entrada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   20
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton O_Compra 
         Appearance      =   0  'Flat
         Caption         =   "&Compra"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Label lblCSO_SimplesNacional 
      Caption         =   "Simples Nacional - CSO"
      Height          =   255
      Left            =   3930
      TabIndex        =   44
      Top             =   510
      Width           =   1785
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   210
      Top             =   7260
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
      Bands           =   "OpEntrada.frx":4EA24
   End
   Begin VB.Label Label2 
      Caption         =   "Código Fiscal (CFOP)"
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   510
      Width           =   1650
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmOpEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsOpEntradas As Recordset
Dim Num_Registro As Variant

'27/02/2004 - Daniel
'Flag para indicar se é o cliente PSV Informática
Dim m_blnPSV      As Boolean
'Flag para guardar o Código da Op. de Entrada
'para comparação na ChecaEstorno
Dim m_intCodigo   As Integer
'Flag para gerenciamento do cliente Agrofarm
Dim m_blnAgrofarm As Boolean

Private Sub ShowRecord()

  Código.Text = rsOpEntradas("Código")
  Nome.Text = rsOpEntradas("Nome") & ""
   
  If rsOpEntradas("Tipo") = "C" Then O_Compra.Value = True
  If rsOpEntradas("Tipo") = "T" Then O_Trans_Entrada.Value = True
  If rsOpEntradas("Tipo") = "A" Then O_Ajuste_Entrada.Value = True
  If rsOpEntradas("Tipo") = "G" Then O_Grátis_Entrada.Value = True
  If rsOpEntradas("Tipo") = "E" Then O_Empréstimo.Value = True
  If rsOpEntradas("Tipo") = "P" Then O_Pedido.Value = True
  If rsOpEntradas("Tipo") = "D" Then O_Devolução.Value = True
  
  '17/09/2009 - mpdea
  'Modelo de documento fiscal
  txtModeloDocumentoFiscal.Text = rsOpEntradas.Fields("ModeloDocumentoFiscal").Value & ""
  
  Estoque.Value = -rsOpEntradas("Estoque")
  Dinheiro.Value = -rsOpEntradas("Dinheiro")
  Comissão.Value = -rsOpEntradas("Comissão")
  ' Etiquetas.Value = -rsOpEntradas("Etiquetas")
  ' Nota.Value = -rsOpEntradas("Nota")
  ' ICM.Value = -rsOpEntradas("ICM")
  IPI.Value = -rsOpEntradas("IPI")
  IPI_TOT.Value = -rsOpEntradas("IPI TOT")
  
  O_Base_IPI.Value = -rsOpEntradas("Base ICM com IPI")
  '17/11/2006 - Anderson
  'Solicitação - Technomax
  'Utilizado para somar o valor do frete no calculo de icms para movimentos de entrada.
  O_Base_ICMS_Frete.Value = -rsOpEntradas("BaseICMSFrete")
  
  '01/12/2006 - Anderson
  'Solicitação - Technomax
  'Incide ICMS sobre IPI.
  O_ICMS_IPI.Value = -rsOpEntradas("ICMSSobreIPI")
  
  Senha.Value = -rsOpEntradas("Senha")
  
  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio Calçados - Otimização liberada
  '             para todos usuários do Quick Store
  '
  'Tratamento para o campo Emitir NF automaticamente
  If rsOpEntradas.Fields("EmitirNFManualmente").Value Then
    chkEmitirNFManualmente.Value = vbChecked
  Else
    chkEmitirNFManualmente.Value = vbUnchecked
  End If
  
  '27/02/2004 - Daniel
  'Case: PSV
  If m_blnPSV Then
    chkEstorno.Enabled = O_Ajuste_Entrada.Value
    chkEstorno.Value = -rsOpEntradas.Fields("Estorno").Value
  End If
  '---------------------------------------------------------
  
  Grava_Custo.Value = -rsOpEntradas("Gravar Custo")
  
  
  '----------------------------------------------------------------------------
  'Data.............: 22/09/2005
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Gravar o preço de Custo no campo Preço de Lista sem IPI
  '                   utilizado na pasta Cálculos no Cadastro de Produtos
  If chkGravaCustoPrecoListaSemIPI.Visible Then
    chkGravaCustoPrecoListaSemIPI.Value = IIf((rsOpEntradas.Fields("GravaCustoPrecoListaSemIPI").Value), vbChecked, vbUnchecked)
  End If
  '----------------------------------------------------------------------------
  
  
  '----------------------------------------------------------------------------
  'Data.............: 16/05/2006
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Somar frete no custo dos produtos
  chkSomarFreteCustoProduto.Value = IIf((rsOpEntradas.Fields("SomarFreteCustoProduto").Value), vbChecked, vbUnchecked)
  '----------------------------------------------------------------------------
  
  
  Somar_Frete_Total.Value = -rsOpEntradas("Somar Frete ao Total")
  Somar_Frete_Custo.Value = -rsOpEntradas("Somar Frete ao Custo")
  
  '23/08/2004 - Daniel
  'Adicionado campo para o tratamento de permissão
  'para alterar o preço de venda
  If Len(rsOpEntradas("Tabela").Value) > 0 Then
    cboTabela.Enabled = True
    lblTabela.Enabled = True
  Else
    cboTabela.Enabled = False
    lblTabela.Enabled = False
  End If
    
  cboTabela.Text = rsOpEntradas("Tabela").Value & ""
  
  If rsOpEntradas("PermitirAlterPreco").Value Then
    chkPermitir.Value = vbChecked
  Else
    chkPermitir.Value = vbUnchecked
  End If
  
  Código_Fiscal.Text = rsOpEntradas("Código Fiscal") & ""
  
  '30/03/2011 - Andrea
  txtCSOSN.Text = rsOpEntradas("CSO") & ""
  
  '18/02/2005 - Daniel
  '
  'Solicitante: Agrofarm - RS
  '
  'Gerenciamento do campo Informante próprio (P)
  If m_blnAgrofarm Then chkInformante.Value = IIf(rsOpEntradas("InformanteProprio").Value, vbChecked, vbUnchecked)
  
  '29/08/2007 - Anderson
  'Implementação do campo para automatização do preço de custo.
  'Solicitante: Candy Clean
  chkPrecoCustoCalculado.Value = -rsOpEntradas("PrecoCustoCalculado")
  
  Num_Registro = rsOpEntradas.Bookmark
  
End Sub

Private Sub DeleteRecord()
  Dim Aux_Filial As Integer
  Dim Aux_Sequência As Integer
  Dim rsEntradas As Recordset

  If IsNull(Num_Registro) Then
    DisplayMsg "Encontre uma operação antes."
    Exit Sub
  End If
  
  'Operação bloqueada
  If rsOpEntradas.Fields("Locked").Value Then
    DisplayMsg "Não é possível excluir. Operação bloqueada pelo sistema."
    Exit Sub
  End If
  
  Call StatusMsg("Aguarde, verificando se este Tipo de Operacão não está em uso...")
  
  Set rsEntradas = db.OpenRecordset("SELECT * FROM Entradas WHERE Operação = " & rsOpEntradas("Código").Value, dbOpenDynaset)
  
  If Not rsEntradas.EOF Then
    DisplayMsg "Este Tipo de Operação não pode ser apagado por estar em uso."
    rsEntradas.Close
    Set rsEntradas = Nothing
    Exit Sub
  End If
  
  rsEntradas.Close
  Set rsEntradas = Nothing
  
  gsTitle = LoadResString(201)
  gsMsg = "Você deseja realmente apagar este Tipo de Operação?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    rsOpEntradas.Delete
    DisplayMsg "Tipo de Operação apagado com sucesso."
    Call ClearScreen
  End If

End Sub

Private Sub UpdateRecord()
  Dim Erro As Integer
  
  Call StatusMsg("")
  
  On Error GoTo Processa_Erro
  
  '23/08/2004 - Daniel
  'Tratamento para Tabela
  If Len(cboTabela.Text) > 0 Then
    If O_Compra.Value = False Then
      MsgBox "Habilite a Classificação como Compra para gravação.", vbExclamation, "Permitir alteração no Preço de Venda"
      Exit Sub
    End If
  End If
  '--------------------------------------------------------------------------------------------------------------------
  
  'Operação bloqueada
  If Not IsNull(Num_Registro) Then
    If rsOpEntradas.Fields("Locked").Value Then
      DisplayMsg "Operação bloqueada pelo sistema. Necessário senha do Gerente."
      If Not frmGerente.gbSenhaGerente Then
        Exit Sub
      End If
    End If
  End If
  
  Rem Verifica Conta
  Erro = False
  If IsNull(Código.Text) Then Erro = True
  If Erro = False Then If Código.Text = "" Then Erro = True
  If Erro = False Then If Not IsNumeric(Código.Text) Then Erro = True
  If Erro = False Then If Val(Código.Text) <= 0 Or Val(Código.Text) > 499 Then Erro = True
  
  If Erro = True Then
    
    '=============================================================
    ' Se operEntrada = -1 ou -2 (pode atualizar apenas CFOP e CSO
    If Val(Código.Text) = -1 Or Val(Código.Text) = -2 Then
        Call StatusMsg("Gravando ...")
        DoEvents
        
        'Inicia transação
        ws.BeginTrans
        
        rsOpEntradas.Edit
        rsOpEntradas.Fields("CSO") = txtCSOSN.Text
        rsOpEntradas.Fields("Código Fiscal") = Código_Fiscal.Text
        rsOpEntradas.Update
        
        Num_Registro = rsOpEntradas.LastModified
        rsOpEntradas.Bookmark = Num_Registro

        
        'Efetua registro do Log
        If rsOpEntradas.Fields("Locked").Value Then
          db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & _
            Now & "#, 'Operação de Entrada bloqueada foi alterada pelo usuário " & _
            gnUserCode & "-" & gsUserName & " através de senha do gerente', 'CADASTRO')", dbFailOnError
        End If
  
        'Finaliza transação
        ws.CommitTrans
  
        MsgBox "Apenas os campos CFOP e CSO podem e foram alterados com sucesso para esta Operação de Entrada.", vbInformation, "Atenção"
  
        Call StatusMsg("")
        Exit Sub
    End If
    '=============================================================
    
    DisplayMsg "Escolha códigos entre 1 e 499"
    Código.SetFocus
    Exit Sub
  End If
  
  
  Erro = False
  If IsNull(Nome.Text) Then Erro = True
  If Erro = False Then If Nome.Text = "" Then Erro = True
  If Erro = True Then
    DisplayMsg "Por favor digite o nome da operação."
    Nome.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Gravando ...")
  DoEvents
  
  'Inicia transação
  ws.BeginTrans
  
  With rsOpEntradas
    If IsNull(Num_Registro) Then
      .AddNew
      .Fields("Código") = Val(Código.Text)
    Else
      .Edit
    End If
    
    .Fields("Nome") = Nome.Text
    
    '17/09/2009 - mpdea
    'Modelo de documento fiscal
    rsOpEntradas.Fields("ModeloDocumentoFiscal").Value = txtModeloDocumentoFiscal.Text
    
    If O_Compra.Value = True Then .Fields("Tipo") = "C"
    If O_Trans_Entrada.Value = True Then .Fields("Tipo") = "T"
    If O_Ajuste_Entrada.Value = True Then .Fields("Tipo") = "A"
    If O_Grátis_Entrada.Value = True Then .Fields("Tipo") = "G"
    If O_Empréstimo.Value = True Then .Fields("Tipo") = "E"
    If O_Pedido.Value = True Then .Fields("Tipo") = "P"
    If O_Devolução.Value = True Then .Fields("Tipo") = "D"
    
    If Estoque.Value = 1 Then .Fields("Estoque") = True
    If Estoque.Value = 0 Then .Fields("Estoque") = False
    
    If Dinheiro.Value = 1 Then .Fields("Dinheiro") = True
    If Dinheiro.Value = 0 Then .Fields("Dinheiro") = False
    
    If Comissão.Value = 1 Then .Fields("Comissão") = True
    If Comissão.Value = 0 Then .Fields("Comissão") = False
    
    ' If Etiquetas.Value = 1 Then .Fields("Etiquetas") = True
    ' If Etiquetas.Value = 0 Then .Fields("Etiquetas") = False
    
    
    ' If Nota.Value = 1 Then .Fields("Nota") = True
    ' If Nota.Value = 0 Then .Fields("Nota") = False
    
    ' If ICM.Value = 1 Then .Fields("ICM") = True
    ' If ICM.Value = 0 Then .Fields("ICM") = False
    
    If IPI.Value = 1 Then .Fields("IPI") = True
    If IPI.Value = 0 Then .Fields("IPI") = False
         
    If IPI_TOT.Value = 1 Then .Fields("IPI TOT") = True
    If IPI_TOT.Value = 0 Then .Fields("IPI TOT") = False
     
         
    If Grava_Custo.Value = 1 Then .Fields("Gravar Custo") = True
    If Grava_Custo.Value = 0 Then .Fields("Gravar Custo") = False
    
    
    '----------------------------------------------------------------------------
    'Data.............: 22/09/2005
    'Desenvolvedor....: mpdea
    'Solicitante......: Pavinato
    'Descrição........: Gravar o preço de Custo no campo Preço de Lista sem IPI
    '                   utilizado na pasta Cálculos no Cadastro de Produtos
    If chkGravaCustoPrecoListaSemIPI.Visible Then
      .Fields("GravaCustoPrecoListaSemIPI").Value = IIf((chkGravaCustoPrecoListaSemIPI.Value = vbChecked), True, False)
    End If
    '----------------------------------------------------------------------------
    
    
    '----------------------------------------------------------------------------
    'Data.............: 16/05/2006
    'Desenvolvedor....: mpdea
    'Solicitante......: Pavinato
    'Descrição........: Somar frete no custo dos produtos
    .Fields("SomarFreteCustoProduto").Value = IIf((chkSomarFreteCustoProduto.Value = vbChecked), True, False)
    '----------------------------------------------------------------------------
    
    If O_Base_IPI.Value = 1 Then .Fields("Base ICM com IPI") = True
    If O_Base_IPI.Value = 0 Then .Fields("Base ICM com IPI") = False
    
    '17/11/2006 - Anderson
    'Solicitação - Technomax
    'Utilizado para somar o valor do frete no calculo de icms para movimentos de entrada.
    If O_Base_ICMS_Frete.Value = 1 Then .Fields("BaseICMSFrete") = True
    If O_Base_ICMS_Frete.Value = 0 Then .Fields("BaseICMSFrete") = False
        
    '01/12/2006 - Anderson
    'Solicitação - Technomax
    'Incide ICMS sobre IPI.
    If O_ICMS_IPI.Value = 1 Then .Fields("ICMSSobreIPI") = True
    If O_ICMS_IPI.Value = 0 Then .Fields("ICMSSobreIPI") = False
    
    If Senha.Value = 1 Then .Fields("Senha") = True
    If Senha.Value = 0 Then .Fields("Senha") = False
    
    '19/05/2005 - Daniel
    '
    'Solicitante: Pedágio Calçados - Otimização liberada
    '             para todos usuários do Quick Store
    '
    'Tratamento para o campo Emitir NF automaticamente
    If chkEmitirNFManualmente.Value = vbChecked Then
      .Fields("EmitirNFManualmente").Value = True
    Else
      .Fields("EmitirNFManualmente").Value = False
    End If
    
    If Somar_Frete_Total.Value = 1 Then .Fields("Somar Frete ao Total") = True
    If Somar_Frete_Total.Value = 0 Then .Fields("Somar Frete ao Total") = False
    
    '23/08/2004 - Daniel
    'Adicionado campo para o tratamento de permissão
    'para alterar o preço de venda
    '19/07/2007 - Anderson
    'Alteração realizada para gravar a tabela selecionada no cadastro de operações de Entrada.
    'If Len(cboTabela.Text) Then .Fields("Tabela").Value = Trim(cboTabela.Text)
    .Fields("Tabela").Value = "" & Trim(cboTabela.Text)
    
    If chkPermitir.Value Then
      .Fields("PermitirAlterPreco").Value = True
    Else
      .Fields("PermitirAlterPreco").Value = False
    End If
    
    If Somar_Frete_Custo.Value = 1 Then .Fields("Somar Frete ao Custo") = True
    If Somar_Frete_Custo.Value = 0 Then .Fields("Somar Frete ao Custo") = False
    
    .Fields("Código Fiscal") = Código_Fiscal.Text
    
    '30/03/2011 - Andrea
    .Fields("CSO") = txtCSOSN.Text
    
    '27/02/2004 - Daniel
    'Case: PSV
    If m_blnPSV Then
      m_intCodigo = CInt(Código.Text)  'txt do Código
    
      .Fields("Estorno").Value = (chkEstorno.Enabled And chkEstorno.Value = 1 And chkEstorno.Visible And ChecaEstorno)
    End If
    '---------------------------------------------------------------------------------------------------------------
      
    '18/02/2005 - Daniel
    '
    'Solicitante: Agrofarm - RS
    '
    'Gerenciamento do campo Informante próprio (P)
    If m_blnAgrofarm Then
      If chkInformante.Value = vbChecked Then
        .Fields("InformanteProprio").Value = True
      Else
        .Fields("InformanteProprio").Value = False
      End If
    End If
    
    '29/08/2007 - Anderson
    'Implementação do campo para automatização do preço de custo.
    'Solicitante: Candy Clean
    If chkPrecoCustoCalculado.Value = vbChecked Then
      .Fields("PrecoCustoCalculado").Value = True
    Else
      .Fields("PrecoCustoCalculado").Value = False
    End If
        
    .Update
    Num_Registro = .LastModified
    .Bookmark = Num_Registro
  
  End With
  
  'Efetua registro do Log
  If rsOpEntradas.Fields("Locked").Value Then
    db.Execute "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & _
      Now & "#, 'Operação de Entrada bloqueada foi alterada pelo usuário " & _
      gnUserCode & "-" & gsUserName & " através de senha do gerente', 'CADASTRO')", dbFailOnError
  End If
  
  'Finaliza transação
  ws.CommitTrans
  
  Call StatusMsg("")
  
  Exit Sub
  
Processa_Erro:
  ws.Rollback
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar gravar registro."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & " - " & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub


Private Sub ClearScreen()

  Call StatusMsg("")
  Código.Text = ""
  Nome.Text = ""
  
  '17/09/2009 - mpdea
  'Modelo de documento fiscal
  txtModeloDocumentoFiscal.Text = ""
  
  O_Compra.Value = True
  
  Estoque.Value = 0
  Dinheiro.Value = 0
  ' Etiquetas.Value = 0
  Comissão.Value = 0
  'Nota.Value = 0
  'ICM.Value = 0
  IPI.Value = 0
  IPI_TOT.Value = 0
  
  O_Base_IPI.Value = 0
  '17/11/2006 - Anderson
  'Solicitação - Technomax
  'Utilizado para somar o valor do frete no calculo de icms para movimentos de entrada.
  O_Base_ICMS_Frete.Value = 0
  
  '01/12/2006 - Anderson
  'Solicitação - Technomax
  'Incide ICMS sobre IPI.
  O_ICMS_IPI.Value = 0
  
  Senha.Value = 0
  
  '19/05/2005 - Daniel
  '
  'Solicitante: Pedágio Calçados - Otimização liberada
  '             para todos usuários do Quick Store
  '
  'Tratamento para o campo Emitir NF automaticamente
  chkEmitirNFManualmente.Value = vbUnchecked
  
  '27/02/2004 - Daniel
  'Case: PSV
  If m_blnPSV Then
    chkEstorno.Value = vbUnchecked
    chkEstorno.Enabled = False
  End If
  '--------------------------------
  
  
  '----------------------------------------------------------------------------
  'Data.............: 22/09/2005
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Gravar o preço de Custo no campo Preço de Lista sem IPI
  '                   utilizado na pasta Cálculos no Cadastro de Produtos
  If chkGravaCustoPrecoListaSemIPI.Visible Then
    chkGravaCustoPrecoListaSemIPI.Value = vbUnchecked
  End If
  '----------------------------------------------------------------------------
  
  
  '----------------------------------------------------------------------------
  'Data.............: 16/05/2006
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Somar frete no custo dos produtos
  chkSomarFreteCustoProduto.Value = vbUnchecked
  '----------------------------------------------------------------------------
  
  
  Grava_Custo.Value = 0
  
  Somar_Frete_Total.Value = 0
  Somar_Frete_Custo.Value = 0
  
  '23/08/2004 - Daniel
  'Adicionado campo para o tratamento de permissão
  'para alterar o preço de venda
  cboTabela.Text = ""
  chkPermitir.Value = vbUnchecked
  
  Código_Fiscal.Text = ""
  
  '30/03/2011 - Andrea
  txtCSOSN.Text = ""

  
  '18/02/2005 - Daniel
  '
  'Solicitante: Agrofarm - RS
  '
  'Gerenciamento do campo Informante próprio (P)
  If m_blnAgrofarm Then chkInformante.Value = vbUnchecked
  '------------------------------------------------------
  
  '29/08/2007 - Anderson
  'Implementação do campo para automatização do preço de custo.
  'Solicitante: Candy Clean
  chkPrecoCustoCalculado.Value = vbUnchecked

  If Not rsOpEntradas.EOF Then
    On Error Resume Next
    rsOpEntradas.MoveFirst
    rsOpEntradas.MovePrevious
    On Error GoTo 0
  End If
  
  Num_Registro = Null
  
  Código.SetFocus

End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsOpEntradas
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsOpEntradas
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsOpEntradas
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsOpEntradas
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
    Case "miOpPrevious"
      Call MovePrevious
    Case "miOpNext"
      Call MoveNext
    Case "miOpLast"
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
  End Select
  
  '27/02/2004 - Daniel
  'Case: PSV
  If m_blnPSV Then VerificaClassificacao
  
End Sub

Private Sub chkPermitir_Click()
  If chkPermitir.Value = 0 Then
    cboTabela.Enabled = False
    lblTabela.Enabled = False
    '07/06/2005 - Daniel
    'Solicitado este tratamento pela TECHNOMAX Informática
    cboTabela.Text = ""
  End If
  
  If chkPermitir.Value = 1 Then
    cboTabela.Enabled = True
    lblTabela.Enabled = True
  End If
End Sub

Private Sub Código_LostFocus()
  
  If Not IsNumeric(Código.Text) Then Exit Sub
  If Val(Código) <= 0 Then Exit Sub
  
  With rsOpEntradas
    .FindFirst "Código = " & CInt(Código.Text)
    If Not .NoMatch Then
      Call ShowRecord
    Else
      Beep
    End If
  End With
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      'strfile = "D:\SoftwaresInstalados\QuickStoreHelp\QuickStoreHelp.chm"
      'Call objHelp.Show(strfile, "QuickStore10Help")
      Call objHelp.Show(strfile, "QuickStore10Help", 10005)
      Set objHelp = Nothing
  Else
      Call HandleKeyDown(KeyCode, Shift)
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  Num_Registro = Null
  
  Set rsOpEntradas = db.OpenRecordset("SELECT * FROM [Operações Entrada] ORDER BY Código", dbOpenDynaset)
  
  Call ActiveBarLoadToolTips(Me)
  
  
  '----------------------------------------------------------------------------
  'Data.............: 22/09/2005
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Gravar o preço de Custo no campo Preço de Lista sem IPI
  '                   utilizado na pasta Cálculos no Cadastro de Produtos
  chkGravaCustoPrecoListaSemIPI.Visible = g_blnGravaCustoPrecoListaSemIPI
  '----------------------------------------------------------------------------
  

  '27/02/2004 - Daniel
  'Case.......: PSV Informática
  'Finalidade.: Compôr o field Estorno em Operações Entrada
  If CheckSerialCaseMod("QS35552-811", "QS37705-639", "QS37825-830", "QS38933-772", "QS39369-521") Then
     m_blnPSV = True
     
     chkEstorno.Visible = True
     chkEstorno.Enabled = False
  End If
  '-----------------------------------------

  '23/08/2004 - Daniel
  'Implementação para permissão de alteração
  'de preço
  datTabela.DatabaseName = gsQuickDBFileName

  lblTabela.Enabled = False
  cboTabela.Enabled = False
  '-----------------------------------------

  '18/02/2005 - Daniel
  '
  'Solicitante: Agrofarm - RS
  '
  'Gerenciamento do campo Informante próprio (P). Para toda operação
  'que possuir este campo habilitado, no momento da geração do
  'arquivo para o Sintegra no registro 50 o campo emitente será
  'igual a P.
  'No caso da Agrofarm as vezes eles emitem notas contra eles mesmos
  'no momento da entrada quando alguma venda ao consumidor retorna
  'como devolução
  If CheckSerialCaseMod("QS35815-716", "QS37243-804") Then
    m_blnAgrofarm = True
    fraReceita.Enabled = True
    chkInformante.Enabled = True
  Else
    fraReceita.Enabled = False
    chkInformante.Enabled = False
  End If
  '-----------------------------------------------------------------

  '29/08/2007 - Anderson
  'Implementação do campo para automatização do preço de custo.
  'Solicitante: Candy Clean
  chkPrecoCustoCalculado.Visible = CheckSerialCaseMod("QS37957-281")

  Me.Show
  DoEvents
  
  Call ClearScreen
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsOpEntradas.Close
  Set rsOpEntradas = Nothing
End Sub

Private Sub O_Ajuste_Entrada_Click()
'  Grava_Custo.Enabled = True
  
  '22/09/2005 - mpdea
  'Habilita campos de custo
  Call HabilitaOpcaoGravarCusto(True)
  
  '27/02/2004 - Daniel
  'Case: PSV
  VerificaClassificacao
  
End Sub

Private Sub O_Compra_Click()
  'Grava_Custo.Enabled = True
  
  '22/09/2005 - mpdea
  'Habilita campos de custo
  Call HabilitaOpcaoGravarCusto(True)
End Sub

Private Sub O_Devolução_Click()
  '22/09/2005 - mpdea
  'Habilita campos de custo
  Call HabilitaOpcaoGravarCusto(True)
End Sub

Private Sub O_Empréstimo_Click()
  'Grava_Custo.Enabled = True
  
  '22/09/2005 - mpdea
  'Habilita campos de custo
  Call HabilitaOpcaoGravarCusto(True)
End Sub

Private Sub O_Grátis_Entrada_Click()
'  Grava_Custo.Value = 0
'  Grava_Custo.Enabled = False

  '22/09/2005 - mpdea
  'Desabilita campos de custo
  Call HabilitaOpcaoGravarCusto(False)
End Sub

Private Sub O_Pedido_Click()
'  Grava_Custo.Value = 0
'  Grava_Custo.Enabled = False

  '22/09/2005 - mpdea
  'Desabilita campos de custo
  Call HabilitaOpcaoGravarCusto(False)
End Sub

Private Sub O_Trans_Entrada_Click()
'  Grava_Custo.Value = 0
'  Grava_Custo.Enabled = False

  '22/09/2005 - mpdea
  'Desabilita campos de custo
  Call HabilitaOpcaoGravarCusto(False)
End Sub

Private Sub VerificaClassificacao()
'27/02/2004 - Daniel
'Case: PSV
  If m_blnPSV Then
    chkEstorno.Visible = O_Ajuste_Entrada.Value
    chkEstorno.Enabled = O_Ajuste_Entrada.Value
  End If
End Sub

'22/09/2005 - mpdea
'Centraliza a codificação para habilitar/desabilitar campos de custo
Private Sub HabilitaOpcaoGravarCusto(ByVal blnHabilitar As Boolean)

  '----------------------------------------------------------------------------
  'Data.............: 22/09/2005
  'Desenvolvedor....: mpdea
  'Solicitante......: Pavinato
  'Descrição........: Gravar o preço de Custo no campo Preço de Lista sem IPI
  '                   utilizado na pasta Cálculos no Cadastro de Produtos
  If chkGravaCustoPrecoListaSemIPI.Visible Then
    chkGravaCustoPrecoListaSemIPI.Enabled = blnHabilitar
    If Not blnHabilitar Then
      chkGravaCustoPrecoListaSemIPI.Value = vbUnchecked
    End If
  End If
  '----------------------------------------------------------------------------

  'Gravar preço na tabela CUSTO
  Grava_Custo.Enabled = blnHabilitar
  If Not blnHabilitar Then
    Grava_Custo.Value = vbUnchecked
  End If
  
  '29/08/2007 - Anderson
  'Implementação do campo para automatização do preço de custo.
  'Solicitante: Candy Clean
  chkPrecoCustoCalculado.Enabled = blnHabilitar
  If Not blnHabilitar Then
    chkPrecoCustoCalculado.Value = vbUnchecked
  End If
End Sub

Private Function ChecaEstorno() As Boolean
'Fará a verificação se já existe algum registro
'com o field Estorno Verdadeiro na tabela
'[Operações Entrada]
  Dim rstOperacoesEntrada As Recordset
  
  Set rstOperacoesEntrada = db.OpenRecordset(" SELECT Código, Nome, Estorno FROM [Operações Entrada] ", dbOpenDynaset)
  
  With rstOperacoesEntrada
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        If .Fields("Estorno").Value = True Then
          If m_intCodigo <> .Fields("Código").Value Then 'Tratamento para atualizações do registro...
            ChecaEstorno = False 'para não criar um segundo registro como True
            MsgBox "Já existe Operação de Entrada com o campo Realiza Estorno da Reserva marcado. " & .Fields("Código").Value & " - " & .Fields("Nome").Value, vbExclamation, "Quick Store"
            chkEstorno.Value = vbUnchecked
            Exit Function
          End If
        End If
      .MoveNext
      Loop
    End If
    
    .Close
  End With
  
  Set rstOperacoesEntrada = Nothing
  
  'Se não encontrou ninguém como True então
  ChecaEstorno = True

End Function
