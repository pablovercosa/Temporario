VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSaidasDevolucaoProdutos 
   Caption         =   " Devolução de Produtos"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaidasDevolucaoProdutos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   13035
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   10350
      ScaleHeight     =   1095
      ScaleWidth      =   2595
      TabIndex        =   23
      Top             =   1590
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   90
      TabIndex        =   19
      Top             =   5400
      Width           =   7425
      Begin VB.TextBox txtObsValeCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   600
         MaxLength       =   60
         TabIndex        =   25
         Top             =   690
         Width           =   6705
      End
      Begin VB.CommandButton cmd_imprimir 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imprime Vale"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   3375
      End
      Begin VB.OptionButton opt_relatorio 
         Appearance      =   0  'Flat
         Caption         =   "Modelo Relatório"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2100
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton opt_ticket 
         Appearance      =   0  'Flat
         Caption         =   "Modelo Ticket"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   180
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000000&
         Caption         =   "Obs:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   26
         Top             =   720
         Width           =   435
      End
   End
   Begin VB.TextBox txt_valorUnitarioProdutoDevolver 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5610
      TabIndex        =   17
      Top             =   2340
      Width           =   1335
   End
   Begin VB.TextBox txt_valorUnitarioProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5610
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txt_descontoVenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5610
      TabIndex        =   13
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CommandButton cmd_visualizarDevolucao 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Visualizar devolução detalhada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5970
      Width           =   3375
   End
   Begin VB.CommandButton cmd_gerarDevolucaoDeProduto 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Confirmar devolução"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2820
      Width           =   12960
   End
   Begin VB.TextBox txt_qtde 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   8
      Top             =   1020
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid gridProdutosDevolvidos 
      Height          =   1785
      Left            =   60
      TabIndex        =   2
      Top             =   3600
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   3149
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      BackColor       =   15066597
      BackColorFixed  =   8454143
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483641
      BackColorBkg    =   16250871
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000000&
      Caption         =   """logotipo.bmp"" no diretório QuickStore\Imagens"
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
      Left            =   10350
      TabIndex        =   24
      Top             =   1050
      Width           =   2595
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000000&
      Caption         =   "Confirma o Valor Unitário do Produto para Devolução"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   18
      Top             =   2370
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000000&
      Caption         =   "Valor Unitário do Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3150
      TabIndex        =   16
      Top             =   1950
      Width           =   2415
   End
   Begin VB.Label lblDescontoVenda 
      BackColor       =   &H80000000&
      Caption         =   "Esta venda foi realizada concedendo desconto total de R$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   270
      TabIndex        =   14
      Top             =   1530
      Width           =   5295
   End
   Begin VB.Label lbl_totalDevolucoes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11100
      TabIndex        =   12
      Top             =   5475
      Width           =   1905
   End
   Begin VB.Label Label3 
      Caption         =   "Total R$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10260
      TabIndex        =   11
      Top             =   5520
      Width           =   795
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000000&
      Caption         =   "Quantidade a ser devolvida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   1110
      Width           =   2505
   End
   Begin VB.Label lbl_nomeProdutoDevolucao 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4050
      TabIndex        =   6
      Top             =   570
      Width           =   8925
   End
   Begin VB.Label Label4 
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   615
      Width           =   795
   End
   Begin VB.Label lbl_produtoDevolucao 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   570
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000000&
      Caption         =   "Lista de produtos já devolvidos desta Sequência (Venda)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   3330
      Width           =   5175
   End
   Begin VB.Label lbl_sequencia 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   105
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Sequência"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmSaidasDevolucaoProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lsSequenciaVenda As Long
Public sCodigoProdutoDevolucao As String
Public sNomeProdutoDevolucao As String
Public lsQuantidade As Long
Public sDescontoVenda As String
Public sEmpresaFilial As String
Public sCliente As String
Public sDataDaVenda As String
Public sValorUnitarioProdutoDevolucao As String

Private Sub cmd_gerarDevolucaoDeProduto_Click()

    ' Criar um registro na tabela de <Entradas>
    ' Criar um registro na tabela de <Entradas - Produtos>    ** detalhe: com a OperacaoEntrada (Cod -1 fixa no quick
    ' para tratar OperacaoDevolucaoDoClienteBaseDeTrocaSequencia_original
    
    ' Também movimentar o estoque (lançar aumento de estoque do produto adicionando a quantidade desejada na tela)
    ' A quantidade informada na tela não pode ser maior que a quantidade vendida
    
    GravarDevolucao

End Sub

Private Sub GravarDevolucao()
  On Error GoTo ErrTransaction
  
  Dim nSequencia As Long
  'Variáveis de Tratamento de Erro
  Dim bSequencia As Boolean
  Dim bSeqChanged As Boolean
  Dim nRepeatUpdate3022 As Integer
  Dim nRepeatUpdateLocked As Integer
  Dim nRet As Integer
  Dim i As Integer
  Dim sMsg As String
  Dim blnInTransaction As Boolean
  Dim intRepeatUpdateLocked As Integer
  Dim dblTotalPagar As Double
  
  Dim lCliente As Long
  Dim iQuantidadeItens As Integer
  Dim dPrecoTotal As Double
  Dim dPrecoUnitario As Double
  Dim dValorUnitarioProdutoDevolver As Double
  Dim sValorAux As String
  
  Dim rsEntradas As Recordset
  Dim rsSaidas As Recordset
  Dim rsSaidasProdutos As Recordset
  Dim rsParametros As Recordset
  Dim rsComissao As Recordset
  
  Dim boTemComissao As Boolean
  
  boTemComissao = False
  
  Set rsComissao = db.OpenRecordset("Select * from Comissão where filial = " & gnCodFilial & " and sequência = " & lsSequenciaVenda, dbOpenDynaset)
    
  If Not (rsComissao.EOF And rsComissao.BOF) Then
      boTemComissao = True
  End If
  rsComissao.Close
  Set rsComissao = Nothing
  
  If sCodigoProdutoDevolucao = "0" Or sCodigoProdutoDevolucao = "" Then
      MsgBox "Selecione um produto válido a ser devolvido", vbInformation, "Atenção"
      Exit Sub
  End If
  
  Set rsParametros = db.OpenRecordset("Select * from [Parâmetros Filial] where filial = " & gnCodFilial, dbOpenDynaset)
  
  Set rsSaidas = db.OpenRecordset("Select * from Saídas where filial = " & gnCodFilial & " and sequência = " & lsSequenciaVenda, dbOpenDynaset)
    
  If Not (rsSaidas.EOF And rsSaidas.BOF) Then
      lCliente = rsSaidas.Fields("Cliente").Value
      
      If rsSaidas.Fields("Efetivada").Value = False Then
          MsgBox "Só é permitido realizar devolução de produtos para vendas com status efetivida", vbInformation, "Atenção"
          rsSaidas.Close
          Set rsSaidas = Nothing
          Exit Sub
      End If
  Else
      MsgBox "Venda original não foi encontrada", vbInformation, "Atenção"
      rsSaidas.Close
      Set rsSaidas = Nothing
      Exit Sub
  End If
  rsSaidas.Close
  Set rsSaidas = Nothing
  
  If txt_valorUnitarioProdutoDevolver.Text = "" Then
      MsgBox "Informe o valor unitário do produto a ser devolvido", vbInformation, "Atenção"
      Exit Sub
  End If
  
  Set rsSaidasProdutos = db.OpenRecordset("Select sum(Qtde), sum([Preço Final]) from [Saídas - Produtos] where filial = " & gnCodFilial & " and sequência = " & lsSequenciaVenda & " and Código = '" & sCodigoProdutoDevolucao & "' ", dbOpenDynaset)
    
  If Not (rsSaidasProdutos.EOF And rsSaidasProdutos.BOF) Then
      iQuantidadeItens = rsSaidasProdutos.Fields(0).Value
      dPrecoTotal = rsSaidasProdutos.Fields(1).Value
      dPrecoUnitario = dPrecoTotal / iQuantidadeItens
      
      sValorAux = txt_valorUnitarioProdutoDevolver.Text
      sValorAux = Replace(sValorAux, ".", ",")
      dValorUnitarioProdutoDevolver = CDbl(sValorAux)
      
      If dValorUnitarioProdutoDevolver - 0.01 > dPrecoUnitario Then
          MsgBox "Valor Unitário do Produto para Devolução deve ser igual ou menor que o valor vendido do produto.", vbInformation, "Atenção"
          rsSaidasProdutos.Close
          Set rsSaidasProdutos = Nothing
          Exit Sub
      End If
  End If
  rsSaidasProdutos.Close
  Set rsSaidasProdutos = Nothing
  
  Dim iBuscarItensDevolvidosDoProduto_NestaVenda As Integer
  iBuscarItensDevolvidosDoProduto_NestaVenda = BuscarItensDevolvidosDoProduto_NestaVenda
  
  If iBuscarItensDevolvidosDoProduto_NestaVenda + txt_qtde.Text > iQuantidadeItens Then
      MsgBox "Quantidade de itens a ser devolvido é maior que o limite.", vbInformation, "Atenção"
      rsParametros.Close
      Set rsParametros = Nothing
      Exit Sub
  End If
  
  If txt_qtde.Text < 0 Then
      MsgBox "Quantidade de itens a ser devolvido informado deve ser maior que ZERO", vbInformation, "Atenção"
      rsParametros.Close
      Set rsParametros = Nothing
      Exit Sub
  End If

  Set rsEntradas = db.OpenRecordset("SELECT * FROM Entradas WHERE Filial = " & gnCodFilial & " ORDER BY Sequência", dbOpenDynaset)
  
  Call ws.BeginTrans
  blnInTransaction = True
  
  'Pega número da nova movimentação
  nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("Última Movimentação") + 1
    
  rsParametros.Edit
  rsParametros("Última Movimentação") = nSequencia
  rsParametros.Update

  With rsEntradas
    .AddNew
    .Fields("Sequência") = nSequencia
    sMsg = "inserida"

    .Fields("Filial") = gnCodFilial
    .Fields("Data") = Format(Now, "dd/MM/yyyy")

    ' ********** ATENCAO
    ' Codigo da OperacaoEntrada é -1 SEM COMISSÃO (fixa do Quick para OperacaoDevolucaoDoClienteBaseDeTrocaSequencia_original)
    ' Codigo da OperacaoEntrada é -2 COM COMISSÃO (fixa do Quick para OperacaoDevolucaoDoClienteBaseDeTrocaSequencia_original)
'''    If boTemComissao = False Then
        .Fields("Operação") = -1
'''    Else
'''        .Fields("Operação") = -2
'''    End If

    .Fields("Digitador") = gnUserCode
    .Fields("Fornecedor") = lCliente
    .Fields("Observações") = "OperacaoDevolucaoDoClienteBaseDeTrocaSequencia_original"
    .Fields("ChaveReferenciada").Value = lsSequenciaVenda
    .Fields("Nota Fiscal") = ""
    .Fields("SerieNF").Value = ""
    .Fields("ModeloDocumentoFiscal").Value = ""
    .Fields("Pedido") = ""

'''    .Fields("Produtos") = dPrecoUnitario * txt_qtde.Text
    .Fields("Produtos") = dValorUnitarioProdutoDevolver * txt_qtde.Text
    .Fields("Desconto") = 0
    .Fields("IPI") = 0
    .Fields("Frete") = 0
    .Fields("Base ICM") = 0
    .Fields("Valor ICM") = 0
    .Fields("Base ICM Subs") = 0
    .Fields("Valor ICM Subs") = 0
'''    .Fields("Total") = dPrecoUnitario * txt_qtde.Text
    .Fields("Total") = dValorUnitarioProdutoDevolver * txt_qtde.Text
    .Fields("Caixa") = False

    .Fields("Dinheiro Caixa") = 0
    .Fields("Cheque Caixa") = 0
    .Fields("Conta") = 0
    .Fields("Num Cheque") = ""
    .Fields("Descrição") = ""
    .Fields("Data Emissão") = Format(Now, "dd/MM/yyyy")
    .Fields("Data Acerto Empréstimo") = Format(Now, "dd/MM/yyyy")

    .Fields("NumeroDI") = ""
    .Fields("CodigoExportador") = ""
    .Fields("UFDesembaracoDI") = ""
    .Fields("LocalDesembaracoDI") = ""
    .Fields("NumeroAdicaoDI") = 0
    .Fields("NumeroSeqItemAdicaoDI") = 0
    .Fields("CodigoFabricanteAdicaoDI") = 0
    .Fields("DescontoAdicaoDI") = 0
    
    .Fields("DataDeRegistroDI") = Format(Now, "dd/MM/yyyy")
    .Fields("DataDesembaracoDI") = Format(Now, "dd/MM/yyyy")
    
    .Fields("Consumidor_Final").Value = 0
    .Fields("Presenca_Comprador").Value = 0
    .Fields("FinalidadeNFe").Value = 0
    .Fields("TotalDesoneracaoICMS").Value = 0
    
    .Update
    
    bSeqChanged = False
    bSequencia = False
  End With
  
  
  ' Gravar na tabela <Entradas - Produtos>
  Dim rsEntradaProd As Recordset
  Set rsEntradaProd = db.OpenRecordset("SELECT * FROM [Entradas - Produtos] WHERE Filial = " & gnCodFilial, dbOpenDynaset)
  Dim Cód As String
  
  With rsEntradaProd
      .AddNew
      .Fields("Filial").Value = gnCodFilial
      .Fields("Sequência").Value = nSequencia
      .Fields("Linha").Value = 1
      .Fields("Código").Value = sCodigoProdutoDevolucao
      .Fields("Qtde").Value = txt_qtde.Text
      .Fields("QtdeAtual").Value = 0
      .Fields("EntradaConsignada").Value = False
'''      .Fields("Preço").Value = CSng(Format(dPrecoUnitario, "##,###,##0.00"))
      .Fields("Preço").Value = CSng(Format(dValorUnitarioProdutoDevolver, "##,###,##0.00"))
      .Fields("Desconto").Value = 0
      .Fields("ICM").Value = 0
      .Fields("IPI").Value = 0
'''      .Fields("Preço Final").Value = dPrecoUnitario * txt_qtde.Text
      .Fields("Preço Final").Value = dValorUnitarioProdutoDevolver * txt_qtde.Text
      .Fields("Etiqueta").Value = False
      .Fields("Código Sem Grade") = ""
       
      Dim Tamanho As Integer
      Dim Cor As Integer
      Dim Edição As Long
      Dim Tipo As Integer
      Dim Erro As Integer
       
       Call Acha_Produto(sCodigoProdutoDevolucao, Cód, Tamanho, Cor, Edição, Tipo, Erro)
      If Erro = 0 Then
        .Fields("Código Sem Grade") = Cód
      End If
        
      .Fields("IndiceFinanceiro").Value = 0
      .Fields("ValorIcmsRetido").Value = 0
      .Fields("ValorICMSDesonerado").Value = 0
      .Fields("Percentual_Diferimento") = 0
        
      .Update
  End With
  
  
  nRet = Efetiva_Entrada(gnCodFilial, nSequencia)
    
  If nRet <> 0 Then
      Select Case nRet
        Case -1
          'Ação cancelada
          Call StatusMsg("Ação cancelada.")
        Case 1
          Call DisplayMsg("Código da operação inexistente.")
        Case 2
          Call DisplayMsg("Funcionário inexistente.")
        Case 3
          Call DisplayMsg("Fornecedor inexistente.")
        Case Else
          Call DisplayMsg("Operação NÃO efetivada. Erro" & str(nRet))
      End Select
      ws.Rollback
      blnInTransaction = False
      Exit Sub
  End If

  Call ws.CommitTrans
  blnInTransaction = False
  
  rsParametros.Close
  rsEntradas.Close
  rsEntradaProd.Close
  Set rsParametros = Nothing
  Set rsEntradas = Nothing
  Set rsEntradaProd = Nothing
  
  MsgBox "Devolução de produto realizada com sucesso", vbInformation, "Sucesso"
  
  CarregarGrade
  
  Exit Sub

ErrTransaction:

  rsParametros.Close
  rsEntradas.Close
  rsEntradaProd.Close
  Set rsParametros = Nothing
  Set rsEntradas = Nothing
  Set rsEntradaProd = Nothing

  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Resume
        Else
          'Cancelamento da transação
          If blnInTransaction Then ws.Rollback
          Exit Sub
        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Manutenção - Contas a receber")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Exit Sub
        Case 3 'Encerrar
          End
      End Select
  End Select
End Sub

Private Sub cmd_imprimir_Click()
  On Error GoTo Erro
  
  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome As String
  Dim strNomeLPT As String
  Dim strPortaLPT As String
  Dim intX As Integer
  Dim i As Integer
  
  If opt_ticket.Value = True Then
      strNome = "TICKET"
      strNomeLPT = "NOME IMPRESSORA TICKET"
      strPortaLPT = "PORTA IMPRESSORA TICKET"
  Else
      strNome = "REL"
      strNomeLPT = "NOME IMPRESSORA REL"
      strPortaLPT = "PORTA IMPRESSORA REL"
  End If

  strImpressora = GetSetting("QuickStore", "ConfigLPT", strNomeLPT, "")
  strPorta = GetSetting("QuickStore", "ConfigLPT", strPortaLPT, "")
      
  If Len(Trim(strImpressora)) > 0 And Len(Trim(strPorta)) > 0 Then
      For Each objPrinter In Printers
        If objPrinter.DeviceName = strImpressora And objPrinter.Port = strPorta Then
            Set Printer = objPrinter
            Exit For
        End If
      Next objPrinter
  End If

  Dim sCodigoProduto As String
  Dim sNomeProduto As String
  Dim sCodigoEntrada As String
  Dim sNumItens As String
  Dim sValorUnitario As String
  Dim sValorTotal As String
  Dim sLinha As String
  Dim lContador As Long
  Dim sDataAtual As String
  
  sDataAtual = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  
  Printer.Font = "LUCIDA CONSOLE"

  If strNome = "REL" Then
      Printer.Print "   _________________________________________________________________________________________________________________"
      Printer.Print ""
      Printer.Print "                                                     VALE CRÉDITO"
      Printer.Print ""

      If Picture1.Picture <> 0 Then
          Printer.PaintPicture Picture1, 9000, 500, 2300, 1000
      End If
      
      sLinha = "   Emissão Vale : " & sDataAtual
      Printer.Print sLinha
      
      sLinha = "   Empresa      : " & gsNomeFilial
      Printer.Print sLinha
    
      sLinha = "   Atendente    : " & gsUserName
      Printer.Print sLinha
    
      sLinha = "   Sequência    : " & lsSequenciaVenda
      Printer.Print sLinha
    
      sLinha = "   Data Venda   : " & sDataDaVenda
      Printer.Print sLinha
    
      sLinha = "   Cliente      : " & sCliente
      Printer.Print sLinha
    
      Printer.Print ""
    
      sLinha = "   Código Produto       Nome                                       Cód.Entrada  Núm.itens Valor unitário Valor total"
      Printer.Print sLinha
    
      Printer.Print "   _________________________________________________________________________________________________________________"
      Printer.Print ""
    
      With gridProdutosDevolvidos
          For lContador = 1 To .Rows - 1
    
              ' ************************** ATENÇÃO ***********************************
              ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
              ' De preferência com o mesmo nome da impressora !!!
    
              sCodigoProduto = .TextMatrix(lContador, 2)
              If Len(sCodigoProduto) < 20 Then
                For i = Len(sCodigoProduto) To 19
                    sCodigoProduto = " " & sCodigoProduto
                Next
              End If
    
              sNomeProduto = .TextMatrix(lContador, 3)
              If Len(sNomeProduto) < 40 Then
                For i = Len(sNomeProduto) To 39
                    sNomeProduto = sNomeProduto & " "
                Next
              Else
                  sNomeProduto = Mid(sNomeProduto, 1, 40)
              End If
    
              sCodigoEntrada = .TextMatrix(lContador, 7)
              If Len(sCodigoEntrada) < 11 Then
                For i = Len(sCodigoEntrada) To 10
                    sCodigoEntrada = " " & sCodigoEntrada
                Next
              End If
    
              sNumItens = .TextMatrix(lContador, 4)
              If Len(sNumItens) < 9 Then
                For i = Len(sNumItens) To 8
                    sNumItens = " " & sNumItens
                Next
              End If
              
              sValorUnitario = .TextMatrix(lContador, 5)
              If Len(sValorUnitario) < 14 Then
                For i = Len(sValorUnitario) To 13
                    sValorUnitario = " " & sValorUnitario
                Next
              End If
    
              sValorTotal = .TextMatrix(lContador, 6)
              If Len(sValorTotal) < 11 Then
                For i = Len(sValorTotal) To 10
                    sValorTotal = " " & sValorTotal
                Next
              End If
    
              sLinha = sCodigoProduto
              sLinha = sLinha & " " & sNomeProduto
              sLinha = sLinha & "   " & sCodigoEntrada
              sLinha = sLinha & "  " & sNumItens
              sLinha = sLinha & " " & sValorUnitario
              sLinha = sLinha & " " & sValorTotal
    
              Printer.Print "   " & sLinha
          Next
      End With
    
      Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      Printer.Print "   TOTAL DO VALE CRÉDITO : " & lbl_totalDevolucoes.Caption
      Printer.Print "   -----------------------------------------------------------------------------------------------------------------"
      
      If Trim(txtObsValeCredito.Text) <> "" Then
          Printer.Print ""
          Printer.Print ""
          Printer.Print "   OBS: " & Trim(txtObsValeCredito.Text)
          Printer.Print ""
      End If
      
      Printer.Print "                                                      _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ "
      Printer.Print "   Assinatura do Atendente e carimbo da loja         |                                                              |"
      Printer.Print ""
      Printer.Print "                                                     |                                                              |"
      Printer.Print ""
      Printer.Print "   __________________________________________        |                                                              |"
      Printer.Print "   " & gsUserName
      Printer.Print "                                                     |_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ |"
      Printer.Print "   _________________________________________________________________________________________________________________"
    
      Printer.EndDoc
  Else
  
      ' Modelo Ticket...42 colunas
      
      Printer.Print "__________________________________________"
      
      If Picture1.Picture <> 0 Then
          Printer.PaintPicture Picture1, 1000, 500, 2300, 1000
          Printer.Print ""
          Printer.Print ""
          Printer.Print ""
          Printer.Print ""
          Printer.Print ""
          Printer.Print ""
          Printer.Print ""
          Printer.Print ""
          Printer.Print ""
      End If
      Printer.Print ""
      Printer.Print "               VALE CRÉDITO"
      Printer.Print ""
      
      
      sLinha = "Emissão  : " & sDataAtual
      Printer.Print sLinha
      
      If Len(gsNomeFilial) > 30 Then
          sLinha = "Empresa  : " & Mid(gsNomeFilial, 1, 30)
          Printer.Print sLinha
          Printer.Print Mid(gsNomeFilial, 30, Len(gsNomeFilial) - 30)
      Else
          sLinha = "Empresa  : " & gsNomeFilial
          Printer.Print sLinha
      End If

      If Len(gsUserName) > 30 Then
          sLinha = "Atendente: " & Mid(gsUserName, 1, 30)
          Printer.Print sLinha
          Printer.Print Mid(gsUserName, 30, Len(gsUserName) - 30)
      Else
          sLinha = "Atendente: " & gsUserName
          Printer.Print sLinha
      End If

      sLinha = "Sequência: " & lsSequenciaVenda
      Printer.Print sLinha
    
      sLinha = "Venda    : " & sDataDaVenda
      Printer.Print sLinha
    
      If Len(sCliente) > 30 Then
          sLinha = "Cliente  : " & Mid(sCliente, 1, 30)
          Printer.Print sLinha
          Printer.Print Mid(sCliente, 30, Len(sCliente) - 30)
      Else
          sLinha = "Cliente  : " & sCliente
          Printer.Print sLinha
      End If
    
      Printer.Print ""
    
      sLinha = "Produto Entrada Itens VlrUnitário VlrTotal"
      Printer.Print sLinha
    
      Printer.Print "__________________________________________"
      Printer.Print ""
    
      With gridProdutosDevolvidos
          For lContador = 1 To .Rows - 1
    
              ' ************************** ATENÇÃO ***********************************
              ' Para usar USB tem que COMPARTILHAR a impressora e enviar o arquivo para o compartilhamento
              ' De preferência com o mesmo nome da impressora !!!
    
              sCodigoProduto = .TextMatrix(lContador, 2)
              If Len(sCodigoProduto) < 20 Then
                For i = Len(sCodigoProduto) To 19
                    sCodigoProduto = " " & sCodigoProduto
                Next
              End If
    
              sNomeProduto = .TextMatrix(lContador, 3)
              If Len(sNomeProduto) > 42 Then
                  sNomeProduto = Mid(sNomeProduto, 1, 42)
              End If
      
              sCodigoEntrada = .TextMatrix(lContador, 7)
              If Len(sCodigoEntrada) < 11 Then
                For i = Len(sCodigoEntrada) To 10
                    sCodigoEntrada = " " & sCodigoEntrada
                Next
              End If
    
              sNumItens = .TextMatrix(lContador, 4)
              If Len(sNumItens) < 9 Then
                For i = Len(sNumItens) To 8
                    sNumItens = " " & sNumItens
                Next
              End If
              
              sValorUnitario = .TextMatrix(lContador, 5)
              If Len(sValorUnitario) < 14 Then
                For i = Len(sValorUnitario) To 13
                    sValorUnitario = " " & sValorUnitario
                Next
              End If
    
              sValorTotal = .TextMatrix(lContador, 6)
              If Len(sValorTotal) < 11 Then
                For i = Len(sValorTotal) To 10
                    sValorTotal = " " & sValorTotal
                Next
              End If
    
              sLinha = sCodigoProduto
              Printer.Print sLinha
              
              sLinha = sNomeProduto
              Printer.Print sLinha
              
              sLinha = sCodigoEntrada
              Printer.Print sLinha
              
              sLinha = sNumItens
              sLinha = sLinha & " " & sValorUnitario
              sLinha = sLinha & " " & sValorTotal
              Printer.Print sLinha
          Next
      End With
      
      Printer.Print "------------------------------------------"
      Printer.Print "TOTAL VALE CRÉDITO: " & lbl_totalDevolucoes.Caption
      Printer.Print "------------------------------------------"
      
      If Trim(txtObsValeCredito.Text) <> "" Then
          Printer.Print ""
          Printer.Print ""
          
          If Len(Trim(txtObsValeCredito.Text)) > 35 Then
              Printer.Print "OBS: " & Mid(Trim(txtObsValeCredito.Text), 1, 35)
              Printer.Print Mid(Trim(txtObsValeCredito.Text), 36, Len(Trim(txtObsValeCredito.Text)) - 35)
          Else
              Printer.Print "OBS: " & Trim(txtObsValeCredito.Text)
          End If
          Printer.Print ""
      End If
      
      Printer.Print ""
      Printer.Print "Assinatura do Atendente e carimbo da loja"
      Printer.Print ""
      Printer.Print ""
      Printer.Print ""
      Printer.Print "_____________________________________"
      Printer.Print gsUserName
      Printer.Print ""
      Printer.Print " - - - - - - - - - - - - - - - - - -"
      Printer.Print "|                                   |"
      Printer.Print "|                                   |"
      Printer.Print "|                                   |"
      Printer.Print "|                                   |"
      Printer.Print " - - - - - - - - - - - - - - - - - -"
    
      Printer.EndDoc
  End If
    
  Exit Sub
Erro:
    MsgBox "Erro na impressão do Vale " & Err.Description, vbInformation, "Atenção"
End Sub

Private Sub cmd_visualizarDevolucao_Click()
  Dim sCodEntradaDev As String
  
  If gridProdutosDevolvidos.RowSel > 0 Then
      sCodEntradaDev = gridProdutosDevolvidos.TextMatrix(gridProdutosDevolvidos.RowSel, 7)
      
      Dim objEntrada As frmEntrada
      Set objEntrada = New frmEntrada
      
      objEntrada.bTelaChamadoraDevolucaoProdutos = True
      objEntrada.sCodEntradaDevolucaoProdutos = sCodEntradaDev
      objEntrada.Show

  Else
      MsgBox "Selecione um registro na grade.", vbInformation, "Atenção"
  End If
End Sub

Private Sub Form_Activate()
  txt_qtde.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo Erro

  On Error Resume Next
  Picture1.Picture = LoadPicture(App.Path & "\Imagens\logotipo.bmp")
  
  gridProdutosDevolvidos.ColWidth(0) = 0
  gridProdutosDevolvidos.ColWidth(1) = 1000
  gridProdutosDevolvidos.ColWidth(2) = 2500
  gridProdutosDevolvidos.ColWidth(3) = 5000
  gridProdutosDevolvidos.ColWidth(4) = 900
  gridProdutosDevolvidos.ColWidth(5) = 1150
  gridProdutosDevolvidos.ColWidth(6) = 1100
  gridProdutosDevolvidos.ColWidth(7) = 1050
  
  gridProdutosDevolvidos.Row = 0
  gridProdutosDevolvidos.TextMatrix(0, 0) = ""
  gridProdutosDevolvidos.TextMatrix(0, 1) = "Data"
  gridProdutosDevolvidos.TextMatrix(0, 2) = "Código produto"
  gridProdutosDevolvidos.TextMatrix(0, 3) = "Nome produto"
  gridProdutosDevolvidos.TextMatrix(0, 4) = "Núm. itens"
  gridProdutosDevolvidos.TextMatrix(0, 5) = "Valor unitário"
  gridProdutosDevolvidos.TextMatrix(0, 6) = "Valor total"
  gridProdutosDevolvidos.TextMatrix(0, 7) = "Cód.Entrada"

'  gridProdutosDevolvidos.ColAlignment(1) = Left

  lbl_sequencia.Caption = lsSequenciaVenda
  lbl_produtoDevolucao.Caption = sCodigoProdutoDevolucao
  lbl_nomeProdutoDevolucao.Caption = sNomeProdutoDevolucao
  txt_qtde.Text = 1
  
  If sDescontoVenda <> "0,00" And sDescontoVenda <> "" Then
      txt_descontoVenda.Text = sDescontoVenda
  Else
      txt_descontoVenda.Text = "0.00"
  End If
  
  If sValorUnitarioProdutoDevolucao <> "" Then
      txt_valorUnitarioProduto.Text = FormataValorTextoOriginal(sValorUnitarioProdutoDevolucao, 2)
  End If
  
  If sValorUnitarioProdutoDevolucao <> "" Then
      txt_valorUnitarioProdutoDevolver.Text = FormataValorTextoOriginal(sValorUnitarioProdutoDevolucao, 2)
  End If
  
  CarregarGrade
  
  Exit Sub
Erro:
  MsgBox "Erro ao carregar a tela " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
  
End Sub

Private Function FormataValorTextoOriginal(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTextoOriginal = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
  
  If lngCasasDecimais = 2 Then
      If Len(FormataValorTexto) = 7 Then  ' 9999.99     = 9.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 6)
      ElseIf Len(FormataValorTexto) = 8 Then ' 99999.99    = 99.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 6)
      ElseIf Len(FormataValorTexto) = 9 Then ' 999999.99   = 999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 6)
      ElseIf Len(FormataValorTexto) = 10 Then ' 9999999.99   = 9.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 1) + "." + Mid(FormataValorTexto, 2, 3) + "." + Mid(FormataValorTexto, 5, 6)
      ElseIf Len(FormataValorTexto) = 11 Then ' 99999999.99   = 99.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 2) + "." + Mid(FormataValorTexto, 3, 3) + "." + Mid(FormataValorTexto, 6, 6)
      ElseIf Len(FormataValorTexto) = 12 Then ' 999999999.99   = 999.999.999.99
          FormataValorTexto = Mid(FormataValorTexto, 1, 3) + "." + Mid(FormataValorTexto, 4, 3) + "." + Mid(FormataValorTexto, 7, 6)
      End If
  End If
End Function

Private Sub CarregarGrade()
  Dim rsEntradasDeDev      As Recordset
  Dim strSQL                As String
  Dim dValorDevTotalLinhas  As Double

On Error GoTo ErrHandler

  dValorDevTotalLinhas = 0
  gridProdutosDevolvidos.Rows = 1

  strSQL = "SELECT E.Data, P.Código, P.Qtde, P.Preço, P.[Preço Final], PR.Nome, E.Sequência  FROM Entradas E, [Entradas - Produtos] P, Produtos PR "
  strSQL = strSQL & " WHERE E.Filial = " & gnCodFilial
  strSQL = strSQL & " and E.ChaveReferenciada = '" & lsSequenciaVenda & "' "
  strSQL = strSQL & " and E.Observações = 'OperacaoDevolucaoDoClienteBaseDeTrocaSequencia_original'"
  strSQL = strSQL & " and E.Sequência = P.Sequência"
  strSQL = strSQL & " and E.Filial = P.Filial"
  strSQL = strSQL & " and P.Código = PR.Código"
  
  Set rsEntradasDeDev = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rsEntradasDeDev
      If Not (.BOF And .EOF) Then
          .MoveFirst

          Do Until .EOF
              gridProdutosDevolvidos.AddItem "" & vbTab & .Fields("Data").Value & vbTab & _
                .Fields("Código").Value & "" & vbTab & _
                .Fields("Nome").Value & "" & vbTab & _
                .Fields("Qtde").Value & "" & vbTab & _
                FormataValorTexto(.Fields("Preço").Value, 2) & vbTab & _
                FormataValorTexto(.Fields("Preço Final").Value, 2) & vbTab & _
                .Fields("Sequência").Value

                dValorDevTotalLinhas = dValorDevTotalLinhas + .Fields("Preço Final").Value
              
              .MoveNext
          Loop
      End If
      .Close
  End With
 
  ' -------------------------------------------------
  ' Buscar produtos com grade (se houver)
  Dim rsTamanho As Recordset
  Dim rsCor As Recordset
  Dim sProdutoGradeAux As String
  Dim sTamanho As String
  Dim sCor As String
  
  Set rsTamanho = db.OpenRecordset("Tamanhos", , dbReadOnly)
  Set rsCor = db.OpenRecordset("Cores", , dbReadOnly)
  
  strSQL = "SELECT E.Data, P.Código, P.Qtde, P.Preço, P.[Preço Final], PR.Nome, E.Sequência  FROM Entradas E, [Entradas - Produtos] P, Produtos PR, [Códigos da Grade] PG "
  strSQL = strSQL & " WHERE E.Filial = " & gnCodFilial
  strSQL = strSQL & " and E.ChaveReferenciada = '" & lsSequenciaVenda & "' "
  strSQL = strSQL & " and E.Observações = 'OperacaoDevolucaoDoClienteBaseDeTrocaSequencia_original'"
  strSQL = strSQL & " and E.Sequência = P.Sequência"
  strSQL = strSQL & " and P.Código = PG.Código"
  strSQL = strSQL & " and PG.[Código Original] = PR.Código"
  
  Set rsEntradasDeDev = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rsEntradasDeDev
      If Not (.BOF And .EOF) Then
          .MoveFirst

          Do Until .EOF
              rsTamanho.Index = "Código"
              rsTamanho.Seek "=", Mid(.Fields(1).Value, Len(.Fields(1).Value) - 5, 3)
              If Not rsTamanho.NoMatch Then
                  sTamanho = rsTamanho.Fields("Nome").Value
              Else
                  sTamanho = ""
              End If
              
              rsCor.Index = "Código"
              rsCor.Seek "=", Mid(.Fields(1).Value, Len(.Fields(1).Value) - 2, 3)
              If Not rsCor.NoMatch Then
                  sCor = rsCor.Fields("Nome").Value
              Else
                  sCor = ""
              End If
              sProdutoGradeAux = .Fields("Nome").Value & " " & sTamanho & " " & sCor
          
              gridProdutosDevolvidos.AddItem "" & vbTab & .Fields("Data").Value & vbTab & _
                .Fields("Código").Value & "" & vbTab & _
                sProdutoGradeAux & "" & vbTab & _
                .Fields("Qtde").Value & "" & vbTab & _
                FormataValorTexto(.Fields("Preço").Value, 2) & vbTab & _
                FormataValorTexto(.Fields("Preço Final").Value, 2) & vbTab & _
                .Fields("Sequência").Value

                dValorDevTotalLinhas = dValorDevTotalLinhas + .Fields("Preço Final").Value
              
              .MoveNext
          Loop
      End If
      .Close
  End With
  rsTamanho.Close
  rsCor.Close
  Set rsTamanho = Nothing
  Set rsCor = Nothing
  ' -------------------------------------------------
 
 
  Set rsEntradasDeDev = Nothing
  
  lbl_totalDevolucoes.Caption = FormataValorTexto(dValorDevTotalLinhas, 2)

  Exit Sub

ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
  Exit Sub
End Sub

Private Sub txt_qtde_GotFocus()
    With txt_qtde
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Function BuscarItensDevolvidosDoProduto_NestaVenda() As Integer
  Dim rsEntradasDeDev      As Recordset
  Dim strSQL               As String

On Error GoTo ErrHandler

  strSQL = "SELECT sum(P.Qtde) FROM Entradas E, [Entradas - Produtos] P "
  strSQL = strSQL & " WHERE E.Filial = " & gnCodFilial
  strSQL = strSQL & " and E.ChaveReferenciada = '" & lsSequenciaVenda & "' "
  strSQL = strSQL & " and E.Observações = 'OperacaoDevolucaoDoClienteBaseDeTrocaSequencia_original'"
  strSQL = strSQL & " and E.Sequência = P.Sequência"
  strSQL = strSQL & " and P.Código = '" & sCodigoProdutoDevolucao & "' "
  
  
  Set rsEntradasDeDev = db.OpenRecordset(strSQL, dbOpenDynaset)

  If Not (rsEntradasDeDev.BOF And rsEntradasDeDev.EOF) Then
      If Not IsNull(rsEntradasDeDev.Fields(0).Value) Then
          BuscarItensDevolvidosDoProduto_NestaVenda = rsEntradasDeDev.Fields(0).Value
      Else
          BuscarItensDevolvidosDoProduto_NestaVenda = 0
      End If
  Else
      BuscarItensDevolvidosDoProduto_NestaVenda = 0
  End If
 
  rsEntradasDeDev.Close
  Set rsEntradasDeDev = Nothing

  Exit Function

ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Atenção"
  Exit Function
End Function

