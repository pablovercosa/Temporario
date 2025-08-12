VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.1#0"; "Codejock.CommandBars.v12.1.1.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   " Quick Store 10"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11820
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4740
      Top             =   4500
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   3360
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3960
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   5520
      Top             =   3360
      _Version        =   786433
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "mdiMain.frx":4E95A
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   4770
      Top             =   3330
      _Version        =   786433
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'13/01/2009 - mpdea
'Tela principal
'
'27/01/2009 - mpdea
'Adaptado para o novo menu
'Key: Q7MENU

Private gbToAsk As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
  On Error GoTo ErrHandler
    
    
  If Control.id > 1000 And Control.id <= 5000 And Control.Parameter <> "" Then
    gbPodeGravar = (CLng(Control.Parameter) >= 10)
    gbPodeApagar = (CLng(Control.Parameter) = 11)
  End If
    
  Select Case Control.id
    '----------------------------------------------------------------------------
    'Menu Arquivo
    '----------------------------------------------------------------------------
    Case ID_ITEM_ARQUIVO_ESTACOES_CONECTADAS
      frmUsers.Show
      frmUsers.cmdRefresh.Value = True
      frmUsers.Refresh

    Case ID_ITEM_ARQUIVO_LOGON
      MenuArquivoLogon

    Case ID_ITEM_ARQUIVO_COMPACTAR_BASE
      MenuArquivoCompactarBase

    Case ID_ITEM_ARQUIVO_REPARAR_BASE
      MenuArquivoRepararBase

    Case ID_ITEM_ARQUIVO_EXPORTAR_BASE
      MenuArquivoExportarBase

    Case ID_ITEM_ARQUIVO_BACKUP
      MenuArquivoBackup
    
    Case ID_APP_EXIT
      gbToAsk = True
      Unload Me

    '----------------------------------------------------------------------------
    'Menu Ajuda
    '----------------------------------------------------------------------------
    Case ID_ITEM_AJUDA_CONTEUDO
      HTMLHelpContents 1, "hhlpMain"

    Case ID_ITEM_AJUDA_PESQUISA
      '''HTMLHelpSearch 1, "hhlpMain"
      Dim strfile As String
      Dim objHelp As clsGeral
      Set objHelp = New clsGeral
      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
      Call objHelp.Show(strfile, "QuickStore10Help", 10034)
      Set objHelp = Nothing

    Case ID_ITEM_AJUDA_SOBRE
      Call frmAbout.About(Me, gsNomeEmpresa, gsCGCCPF)

    Case ID_ITEM_AJUDA_REGISTRO
      If Not IsProdutoRegistrado() Then
        frmDeveRegistrar.gsPrefix = "QS"
        frmDeveRegistrar.Show vbModal
      Else
        Call gbConsoleLicencas("QS")
      End If

    Case ID_ITEM_AJUDA_INSTITUCIONAL
      HTMLHelpContents 2, "hhlpMain"
      
    Case ID_ITEM_AJUDA_AGENDA
        Dim rstParametros As Recordset
        Set rstParametros = db.OpenRecordset(" SELECT [Verifica Agenda] FROM " & _
                                             " [Par�metros Filial] WHERE Filial = " & gnCodFilial, dbOpenSnapshot)
        With rstParametros
          If .Fields("Verifica Agenda") = True Then
            Call Verifica_Pend�ncias
            If frmAgenda.lstPend.ListCount > 0 Then
              frmAgenda.Show vbModal
            End If
          End If
          .Close
          Set rstParametros = Nothing
        End With

    '----------------------------------------------------------------------------
    'TAB In�cio
    '----------------------------------------------------------------------------
'''    Case ID_ITEM_INICIO_COLAR
'''      Call EditPaste(Me)
'''
'''    Case ID_ITEM_INICIO_RECORTAR
'''      Call EditCut(Me)
'''
'''    Case ID_ITEM_INICIO_COPIAR
'''      Call EditCopy(Me)

    '19/11/2009 - mpdea
    Case ID_ITEM_INICIO_COCKPIT
      Shell gsCockpitFilename, vbNormalFocus

    Case ID_ITEM_INICIO_LIVRO_PONTO
      frmPonto.Show
      
'''    Case ID_ACESSO_HELP_QUICK
'''      'Dim strfile As String
'''      'Dim objHelp As clsGeral
'''      Set objHelp = New clsGeral
'''      strfile = App.Path & "\QuickStoreHelp\QuickStoreHelp.chm"
'''      Call objHelp.Show(strfile, "QuickStore10Help", 10034)
'''      Set objHelp = Nothing
      
    Case ID_ITEM_INICIO_STANDBY
      frmLogin.Show vbModal
''      If Not IsNull(gManterAtivo) And (gManterAtivo = "Nao" Or gManterAtivo = "") Then
''          gManterAtivo = "Sim"
''          Me.Caption = " StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO StandBy ATIVADO"
''          Timer2.Enabled = True
''      Else
''          gManterAtivo = "Nao"
''          Me.Caption = " Quick Store 10"
''          Timer2.Enabled = False
''      End If
''      frm_manterAtivo.Show

    Case ID_ITEM_INICIO_PARAM_EMPRESA
      frmParametros.Show

    Case ID_ITEM_INICIO_PARAM_IMPOSTO_ESTADUAL
      frmEstados.Show

    Case ID_ITEM_INICIO_PARAM_CONFIG_IMPRESSORA
      frmConfiguraImpressora.Show
    
    '05/05/2009 - mpdea
    'Tema do menu
    Case ID_ITEM_INICIO_PARAM_TEMA_AZUL
      SetMenuTheme Control.id

    Case ID_ITEM_INICIO_PARAM_TEMA_AQUA
      SetMenuTheme Control.id

    Case ID_ITEM_INICIO_PARAM_TEMA_PRETO
      SetMenuTheme Control.id

    '07/07/2004 - Daniel
    'Adicionado Par�metro: Classifica��o de Clientes
    'Case: TV Shopping
    Case ID_ITEM_INICIO_PARAM_CLASS_CLIENTE
      frmClassificacaoClientes.Show

    '02/08/2004 - Daniel
    'Adicionado Par�metro: Faturamento Autom�tico
    'Case: STC de Caxias do Sul
    Case ID_ITEM_INICIO_PARAM_FATURAMENTO_AUTO
      frmParamFaturameAuto.Show

    '15/09/2004 - Daniel
    'Adicionado Par�metro: Configura��o de Sa�das para a Devolu��o de Materiais
    'Case: Livraria Resultado
    Case ID_ITEM_INICIO_PARAM_DEVOL_MATERIAL
      frmParametrosDevolucaoMateriais.Show

    '----------------------------------------------------------------------------
    'TAB Cadastros
    '----------------------------------------------------------------------------
    Case ID_ITEM_CADASTRO_SERVICO
      frmServicos.Show
      
    Case ID_ITEM_CADASTRO_PRODUTO_CFOP
      frmProdutosCFOP.Show
      If frmProdutosCFOP.WindowState = vbMinimized Then
        frmProdutosCFOP.WindowState = vbNormal
      End If
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE
      frmProgramaFidelidadeParametros.Show
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_OPERACOES_SAIDA
      frmProgramaFidelidadeOperacoesSaida.Show
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CLIENTES_NAO_PART
      frmProgramaFidelidadeClientesNaoParticipam.Show
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CNPJ_GRUPOS
      frmProgramaFidelidadeCNPJGrupos.Show

    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CONSULTA_GERENCIAL
      frmProgramaFidelidadeConsultaGerencial.Show

    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_RESGATE_PONTOS
      frmProgramaFidelidadeResgatePontos.Show
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CLIENTE_ENTREGA_RESGATE
      frmProgramaFidelidadeClienteEntregaResgate.Show
      
    Case ID_ITEM_CADASTRO_PRODUTO
      frmProdutos.Show
      If frmProdutos.WindowState = vbMinimized Then
        frmProdutos.WindowState = vbNormal
      End If

    Case ID_ITEM_CADASTRO_CLIENTE_FORNEC
      frmCliFor.Show

    Case ID_ITEM_CADASTRO_CARACT_CLIENTE_FORNEC
      MenuCadastroCaracteristicaClienteFornecedor

    Case ID_ITEM_CADASTRO_TRANSPORTADORA
      frmTransportadoras.Show

    Case ID_ITEM_CADASTRO_USUARIO
      frmFuncionarios.Show

    Case ID_ITEM_CADASTRO_CLASSE
      MenuCadastroClasse Control.Caption

    Case ID_ITEM_CADASTRO_SUBCLASSE
      MenuCadastroSubclasse Control.Caption

    Case ID_ITEM_CADASTRO_COR
      MenuCadastroCor Control.Caption

    Case ID_ITEM_CADASTRO_TAMANHO
      MenuCadastroTamanho Control.Caption

    Case ID_ITEM_CADASTRO_ETIQUETA_PRODUTO
      frmEtiquetas.Show

    Case ID_ITEM_FORMATAR_ETIQUETA_PRODUTO
      frmImprimeEtiq.Show

    Case ID_ITEM_CADASTRO_PESQUISA_1
      MenuCadastroPesquisa Control.Caption, 1

    Case ID_ITEM_CADASTRO_PESQUISA_2
      MenuCadastroPesquisa Control.Caption, 2

    Case ID_ITEM_CADASTRO_PESQUISA_3
      MenuCadastroPesquisa Control.Caption, 3

    Case ID_ITEM_CADASTRO_BANCO
      MenuCadastroBanco Control.Caption

    Case ID_ITEM_CADASTRO_CONTA_CORRENTE
      frmContas.Show

    Case ID_ITEM_CADASTRO_CARTAO
      frmCartoes.Show

    Case ID_ITEM_CADASTRO_CAIXA
      MenuCadastroCaixa Control.Caption

    Case ID_ITEM_CADASTRO_MOEDA
      MenuCadastroMoeda Control.Caption

    Case ID_ITEM_CADASTRO_COTACAO
      frmCotacoes.Show

    Case ID_ITEM_CADASTRO_CLASSIFICACAO_FISCAL
      MenuCadastroClassificacaoFiscal Control.Caption

    Case ID_ITEM_CADASTRO_CENTRO_CUSTO
      MenuCadastroCentroCusto Control.Caption

    '31/03/2004 - Daniel
    'Inclus�o de Cadastro de R�dio e Tipo Comercial
    'otimizado inicialmente para a STC de Caxias do Sul
    '23/07/2004 - Daniel
    'Inclus�o do Cadastro de Autoriza��es
    Case ID_ITEM_CADASTRO_RADIO
      frmRadio.Show

    Case ID_ITEM_CADASTRO_TIPO_COMERCIAL
      frmTipoComercial.Show

    Case ID_ITEM_CADASTRO_AUT_PUBLICIDADE
      frmAutorizacaoPublicidade.Show

    Case ID_ITEM_CADASTRO_SUPERVISOR
      frmSupervisores.Show

    Case ID_ITEM_CADASTRO_RETENCAO
      frmRetencoes.Show

    Case ID_ITEM_CADASTRO_CODIGO_NBM
      '20/06/2005 - Daniel
      'Solicitante...: Pneus & Cia (PE)
      'Cadastro de C�digo NBM (Codifica��o da Nomenclatura Brasileira de Mercadorias)
      'Impacto no resgistro 75 da gera��o do arquivo para o SEF
      frmCodigoNBM.Show

    Case ID_ITEM_CADASTRO_GRUPO_FISCAL
      MenuCadastroGrupoFiscal Control.Caption

    Case ID_ITEM_CADASTRO_MENSAGEM_NOTA_FISCAL
      frmMensagensNotaFiscal.Show

    Case ID_ITEM_CADASTRO_MALA_DIRETA_PREPARAR
      frmGeraMala.Show

    Case ID_ITEM_CADASTRO_MALA_DIRETA_MANUTENCAO
      frmManMalaDireta.Show

    Case ID_ITEM_CADASTRO_MALA_DIRETA_PREPARAR_REMETENTE
      frmImprimeRemetente.Show

    Case ID_ITEM_CADASTRO_MALA_DIRETA_GRUPO
      MenuCadastroMalaDiretaGrupo Control.Caption
      
    Case ID_ITEM_CADASTRO_OPERACAO_ENTRADA
      frmOpEntrada.Show

    Case ID_ITEM_CADASTRO_OPERACAO_SAIDA
      frmOpSaida.Show

    '----------------------------------------------------------------------------
    'TAB Movimento
    '----------------------------------------------------------------------------
    Case ID_ITEM_MOVIMENTO_VENDA_RAPIDA
      frmVendaRap1.Show

    Case ID_ITEM_MOVIMENTO_ENTRADAS
      frmEntrada.Show
      frmEntrada.CheckMovimentacao

    Case ID_ITEM_MOVIMENTO_SAIDAS
      frmSaidas.Show
      frmSaidas.CheckMovimentacao

    Case ID_ITEM_MOVIMENTO_DEVOLUCOES
      frmDevolucoes.Show
      'frmDevolucoes.CheckMovimentacao
      
    Case ID_ITEM_MOVIMENTO_ORDEM_SERVICO
      frmVerificaOS.Show

    Case ID_ITEM_MOVIMENTO_PEDIDOS_WEB
      frmWEB_OrderForms.Show

    '14/09/2009 - mpdea
    Case ID_ITEM_MOVIMENTO_NOTA_FISCAL_ELETRONICA
      frmNFe.Show
      
     '07/08/2014 - jean
    Case ID_IMPORTA_GESTO
      'Call Importa_Gesto
      
        'Aproveitamento de condi��o case para nova fun��o chama site...
        Dim iret As Long
        iret = ShellExecute(Me.hwnd, vbNullString, gPastaRetornoNfe, vbNullString, "c:\", 1)


    '26/02/2004 - Daniel
    'Case: PSV - Manuten��o de Reservas
    Case ID_ITEM_MOVIMENTO_MANUT_RESERVA
      If CheckSerialCaseMod("QS35552-811", "QS37705-639", "QS37825-830", "QS38933-772", "QS39369-521") Then
        frmManutencaoReservas.Show
      End If

    Case ID_ITEM_MOVIMENTO_MANUT_CONSIG_ENTRADA
      frmManConsigEntradas.Show

    '13/08/2004 - Daniel
    'Faturamento Autom�tico - Case: STC
    Case ID_ITEM_MOVIMENTO_FATUR_AUTO
      frmFaturamentoAutomatico.Show

    '17/09/2004 - Daniel
    'Presta��o de Contas com Fornecedores - Case: Resultado
    Case ID_ITEM_MOVIMENTO_PREST_FORNEC
      frmManPrestacaodeContas.Show

    '24/01/2005 - Daniel
    'Importador para a Castro Constru��es
    Case ID_ITEM_MOVIMENTO_IMPORTACAO
      frmImportacao.Show

    '28/10/2002 - mpdea
    'Inclus�o de novas telas
    Case ID_ITEM_MOVIMENTO_APAGAR_EMP_ENTRADA
      frmApagaAcertoEmpEntrada.Show

    Case ID_ITEM_MOVIMENTO_APAGAR_EMP_SAIDA
      frmApagaAcertoEmpSaida.Show

    Case ID_ITEM_MOVIMENTO_APAGAR_ENTRADA
      frmApagaEntradas.Show

    Case ID_ITEM_MOVIMENTO_APAGAR_SAIDA
      frmApagaSaidas.Show

    Case ID_ITEM_MOVIMENTO_APAGAR_MOVIMENTACAO
      frmApagaMovim.Show

    Case ID_ITEM_MOVIMENTO_MANUT_CONSIGNACAO
      frmManutencaoConsignacao.Show

    Case ID_ITEM_MOVIMENTO_MANUT_ORCAMENTO
      frmManutencaoOrcamento.Show

    Case ID_ITEM_MOVIMENTO_TRANSF_FILIAL
      frmTransfere.Show

    Case ID_ITEM_MOVIMENTO_EMPREST_ENTRADA
      frmAcertaEmpEntrada.Show

    Case ID_ITEM_MOVIMENTO_EMPREST_SAIDA
      frmAcertaEmpSaida.Show

    '----------------------------------------------------------------------------
    'TAB Pre�os
    '----------------------------------------------------------------------------
    Case ID_ITEM_PRECO_CRIAR_TAB
      frmPrecosCriaTab.Show

    Case ID_ITEM_PRECO_APAGAR_TAB
      frmPrecosResetTab.Show

    Case ID_ITEM_PRECO_LANCAR
      frmPrecosDigita.Show

    Case ID_ITEM_PRECO_ALTERAR
      frmPrecosAltera.Show

    Case ID_ITEM_PRECO_ALTERAR_CALC
      frmAlteracaoPrecoCusto.Show

    Case ID_ITEM_PRECO_CONFIG_TAB
      frmPrecosConfiguraTab.Show

    Case ID_ITEM_PRECO_COPIAR_TAB_IND
      frmPrecosCopiaIndice.Show

    Case ID_ITEM_PRECO_COPIAR_TAB_VALOR
      frmPrecosCopiaValor.Show

    Case ID_ITEM_PRECO_COPIAR_TAB_CUSTO_MEDIO
      frmPrecosCopiaCustoMedio.Show

    Case ID_ITEM_PRECO_CALC_PRECO
      frmPrecosCalculoVenda.Show

    Case ID_ITEM_PRECO_CALC_PRECO_SIMPLES
      frmPrecosCalculoVendaSimples.Show

    '----------------------------------------------------------------------------
    'TAB Estoque
    '----------------------------------------------------------------------------
    Case ID_ITEM_ESTOQUE_INFO_CONTAR
      frmInformaConta.Show

'''    Case ID_ITEM_ESTOQUE_ACERTAR_CONTAR
'''      gbAcertaGrade = False
'''      frmAcertaEstoque.Show

    Case ID_ITEM_ESTOQUE_INFO_CONTAR_GRADE
      frmInformaContaGrade.Show

    Case ID_ITEM_ESTOQUE_ACERTAR_CONTAR_GRADE
      gbAcertaGrade = True
      frmAcertaEstoque.Show

    '----------------------------------------------------------------------------
    'TAB Financeiro
    '----------------------------------------------------------------------------
    Case ID_ITEM_FINANCEIRO_MOV_MANUAL_CAIXA
      frmMovCaixa.Show

    Case ID_ITEM_FINANCEIRO_APAGA_LANC_CAIXA
      frmApagaCaixa.Show

    Case ID_ITEM_FINANCEIRO_LANC_BANC
      frmLancaContas.Show

    Case ID_ITEM_FINANCEIRO_RECAL_SALDO
      frmRecalcula.Show

    Case ID_ITEM_FINANCEIRO_APAGA_LANC_BANC
      frmApagaLancamentos.Show vbModal

    Case ID_ITEM_FINANCEIRO_CP_LANCAR
      frmLancaCPagar.Show

    Case ID_ITEM_FINANCEIRO_CP_GERAR
      frmGeraPagar.Show

    Case ID_ITEM_FINANCEIRO_CP_MANUT
      frmManContasPagar.Show

    Case ID_ITEM_FINANCEIRO_CP_APAGAR_PAGA
      frmApagaPagas.Show

    Case ID_ITEM_FINANCEIRO_CR_LANCAR
      frmLancaCReceber.Show

    '30/07/2003 - mpdea
    'Verifica altera��o personalizada
    '
    'QS32815-683 = Guarant�
    'QS39240-574 = Barro Queimado
    '
    '03/06/2005 - Daniel
    'Adicionado QS da Filial da Barro Queimado = QS39215-718
    Case ID_ITEM_FINANCEIRO_CR_MANUT
      If CheckSerialCaseMod("QS32815-683", "QS39240-574", "QS39215-718") Then
        frmManContasReceberII.Show
      Else
        frmManContasReceber.Show
      End If

    Case ID_ITEM_FINANCEIRO_CR_APAGAR_RECEBIDA
      frmApagaRecebidas.Show

    Case ID_ITEM_FINANCEIRO_CR_LANCAR_CHEQUE_PRE
      frmChequesPre.Show

    Case ID_ITEM_FINANCEIRO_CR_MANUT_CHEQUE_PRE
      frmManCheques.Show

    Case ID_ITEM_FINANCEIRO_CR_APAGAR_CHEQUE_PRE
      frmApagaCheques.Show

    Case ID_ITEM_FINANCEIRO_CR_LANCAR_CARTAO
      frmLancaCCredito.Show

    Case ID_ITEM_FINANCEIRO_CR_MANUT_CARTAO
      frmManCartoes.Show

    Case ID_ITEM_FINANCEIRO_CR_APAGAR_CARTAO
      frmApagaCartoes.Show

    Case ID_ITEM_FINANCEIRO_CR_LANCAR_CONTA_CLIENTE
      frmLancaContaCliente.Show

    Case ID_ITEM_FINANCEIRO_CR_MANUT_CONTA_CLIENTE
      frmManContas.Show

    Case ID_ITEM_FINANCEIRO_CR_APAGAR_CONTA_CLIENTE
      frmApagaContaCliente.Show

    '08/04/2004 - Daniel
    'Verifica altera��o personalizada
    'QS39823-684 = STC de Caxias do Sul
    Case ID_ITEM_FINANCEIRO_CR_AUT_PUBLICIDADE
      If CheckSerialCaseMod("QS39823-684") Then
        frmConsultaAutorizacao.Show
      End If

    '----------------------------------------------------------------------------
    'TAB Gerador
    '----------------------------------------------------------------------------
    Case ID_ITEM_GERADOR_RELATORIO
      Call ChamaGerador(gsGeradorReportFileName)

    Case ID_ITEM_GERADOR_LAYOUT_NOTA
      MenuGeradorLayout 2, "NOTA"

    Case ID_ITEM_GERADOR_LAYOUT_TICKET
      MenuGeradorLayout 2, "TICKET"

    Case ID_ITEM_GERADOR_LAYOUT_BOLETO
      MenuGeradorLayout 0

    Case ID_ITEM_GERADOR_LAYOUT_CARNET
      MenuGeradorLayout 1

    Case ID_ITEM_GERADOR_ARQ_REC_ESTADUAL
      'Par�metro para chamar o programa InfoICMS somente atrav�s do Quick Store
      Call Shell(gsGeradorRecEstadual & " QuickStore_InfoICMS", vbNormalFocus)
    
    '----------------------------------------------------------------------------
    'TAB Relat�rios
    '----------------------------------------------------------------------------
    '---------------------------------------------------------------------------- [Servi�os]
    Case ID_ITEM_REL_SERVICO_EXECUTADO
      frmRelServicos.Show

    Case ID_ITEM_REL_SERVICO_COMISSAO
      frmRelComServ.Show

    '---------------------------------------------------------------------------- [Produtos]
    Case ID_ITEM_REL_PRODUTO_GERAL
      frmRelProdutos.Show

    Case ID_ITEM_REL_PRODUTO_GRADE
      frmRelGrade.Show

    '---------------------------------------------------------------------------- [Estoque]
    Case ID_ITEM_REL_ESTOQUE_GERAL
      frmRelEstoque2.Show

    Case ID_ITEM_REL_ESTOQUE_GRADE
      frmRelProdGrade.Show

    Case ID_ITEM_REL_ESTOQUE_ANALITICO
      frmRelEstoqueAna.Show

    Case ID_ITEM_REL_ESTOQUE_POR_FILIAL
      frmRelEstoqueFiliais.Show

    '19/01/2005 - mpdea
    'Relat�rio de Estoque das Filiais e Pre�o (Personalizado)
    'Solicitante: Cliente Kilou�a (QS71271-970)
    Case ID_ITEM_REL_ESTOQUE_FILIAL_PRECO
      frmRelEstoquePreco.Show

    Case ID_ITEM_REL_ESTOQUE_PRODUTO_COMPRAR
      frmRelProdComprar.Show

    Case ID_ITEM_REL_ESTOQUE_ACOMPANHAMENTO_PRODUTO
      frmRelAcompaProd.Show

    Case ID_ITEM_REL_ESTOQUE_ACOMPANHAMENTO_ESTOQUE
      frmRelAcompaEstoq.Show

    Case ID_ITEM_REL_ESTOQUE_REGISTRO_INVENTARIO
      frmRegInvent.Show

    Case ID_ITEM_REL_ESTOQUE_CONTAGEM
      frmRelContagem.Show

    Case ID_ITEM_REL_ESTOQUE_CONTAGEM_GRADE
      frmRelContagemGrade.Show

    '---------------------------------------------------------------------------- [Compras e Vendas]
    Case ID_ITEM_REL_CV_VENDA
      frmRelVendas.Show

    '16/01/2008 - Anderson
    'Customiza��o de Relat�rio de Vendas para LL Com�rcio de Ferramentas LTDA.
    Case ID_ITEM_REL_CV_VENDA_2
      frmRelVendasII.Show

    Case ID_ITEM_REL_CV_COMISSAO
      frmRelComissoes.Show

    '26/04/2005 - Daniel
    'Adicionado relat�rio de comiss�es que exibe �s
    'reten��es contidas sobre cart�es
    'Solicitante: Bem Me Quer
    Case ID_ITEM_REL_CV_COMISSAO_RETENCAO
      frmRelComissaoComRetencao.Show

    Case ID_ITEM_REL_CV_COMISSAO_VENDA_VENDEDOR
      frmRelVendaComissao.Show

    Case ID_ITEM_REL_CV_COMPRAS
      frmRelCompras.Show

    '30/10/2007 - Anderson
    'Relat�rio de Produtos a comprar
    'Solicitante: Kings Cross
    Case ID_ITEM_REL_CV_PRODUTO_COMPRAR_FATOR
      frmRelProdutosAComprar.Show

    Case ID_ITEM_REL_CV_VENDA_CLIENTE
      frmRelVendasPorCliente.Show

    '30/03/2005 - Daniel
    'Relat�rio de Vendedores e Comiss�es (Sint�tico)
    'Solicita��o: Bem Me Quer
    'Dispon�vel para todos clientes do Quick Store
    Case ID_ITEM_REL_CV_VENDA_VENDEDOR_COMISSAO
      frmRelVendedorComissoesSintetico.Show

    Case ID_ITEM_REL_CV_VENDA_TAMANHO
      frmRelVendasTamanho.Show

    Case ID_ITEM_REL_CV_VENDA_EDITORA
      frmRelEditoras.Show

    '17/08/2007 - Anderson
    'Altera��o realizada para atender solicita��o da Nutricare (QS73086-490)
    Case ID_ITEM_REL_CV_VENDA_FORNECEDOR
      'frmRelVendasPorFornecedor.Show
      If CheckSerialCaseMod("QS73032-694", "QS73086-490") Then
        frmRelVendasFornecedor.Show
      Else
        frmRelVendasPorFornecedor.Show
      End If

    Case ID_ITEM_REL_CV_VENDA_PRODUTO_CONSIGNADO
      frmRelPrestacaoDeContasComFornecedores.Show

    Case ID_ITEM_REL_CV_PRESTACAO_CONTA
      frmRelPrestacaoContas.Show
    
    Case ID_ITEM_REL_CV_VENDA_POR_VENDEDOR
      frmRelVendasPorVendedor.Show
    '---------------------------------------------------------------------------- [Movimento]
    Case ID_ITEM_REL_MOV_ENTRADA
      frmRelEntradas.Show

    Case ID_ITEM_REL_MOV_SAIDA
      frmRelSaidas.Show
    
    Case ID_ITEM_REL_SAIDAS_ENTRADAS
      frmRelSaidasEntradas.Show

    Case ID_ITEM_REL_MOV_ACERTA_EMPREST_ENTRADA
      frmRelEmpEntrada.Show

    Case ID_ITEM_REL_MOV_ACERTA_EMPREST_SAIDA
      frmRelEmpSaida.Show

    Case ID_ITEM_REL_MOV_ENTRADA_CONSIGNADA
      frmRelEntradasConsignadas.Show

    '---------------------------------------------------------------------------- [Pessoas]
    Case ID_ITEM_REL_CLIENTE_FORNECEDOR
      frmRelCliFor.Show

    Case ID_ITEM_REL_CONTATO_EFETUADO
      frmRelContatos.Show

    Case ID_ITEM_REL_CONTATO_DATA_ANIVERSARIO
      frmRelAniver.Show

    Case ID_ITEM_REL_USUARIO_FUNCIONARIO
      frmRelFuncGeral.Show

    Case ID_ITEM_REL_LIVRO_PONTO
      frmRelPonto.Show

    '---------------------------------------------------------------------------- [Cadastro]
    Case ID_ITEM_REL_CADASTRO_CLASSE
      frmRelClasses.Show

    Case ID_ITEM_REL_CADASTRO_SUBCLASSE
      frmRelSubClasses.Show

    Case ID_ITEM_REL_CADASTRO_COR
      frmRelCores.Show

    Case ID_ITEM_REL_CADASTRO_TAMANHO
      frmRelTamanhos.Show

    Case ID_ITEM_REL_CADASTRO_ETIQUETA_PRODUTO
      frmImprimeEtiq.Show

    Case ID_ITEM_REL_CADASTRO_BANCO
      frmRelBancos.Show

    Case ID_ITEM_REL_CADASTRO_CARTAO
      frmRelCartoes.Show

    Case ID_ITEM_REL_CADASTRO_MOEDA
      frmRelMoedas.Show

    Case ID_ITEM_REL_CADASTRO_COTACAO
      frmRelCotacoes.Show

    Case ID_ITEM_REL_CADASTRO_CENTRO_CUSTO
      frmRelCustos.Show

    '---------------------------------------------------------------------------- [Financeiro]
    Case ID_ITEM_REL_FINANC_CAIXA
      frmRelCaixa.Show

    '24/03/2005 - Daniel
    'Adicionado Rel. de Cart�es de Cr�dito (Posi��o Di�ria)
    'Solicitante: Bem Me Quer
    Case ID_ITEM_REL_FINANC_CARTAO
      frmRelLancCartaoPosiDiaria.Show

    Case ID_ITEM_REL_FINANC_LANC_BANCARIO
      frmRelLancamentos.Show

    Case ID_ITEM_REL_FINANC_SALDO_CC
      frmRelSaldos.Show

    Case ID_ITEM_REL_FINANC_DIARIO_1
      frmRelFinanc1.Show

    Case ID_ITEM_REL_FINANC_DIARIO_2
      frmRelFinanc2.Show

    Case ID_ITEM_REL_FINANC_LUCRATIVIDADE
      frmRelLucratividade.Show

    Case ID_ITEM_REL_FINANC_GERAL
      frmRelFinGeral.Show

    '19/07/2006 - Andrea
    'Inclus�o de relat�rio Recebimentos por Forma de Pagamento
    Case ID_ITEM_REL_FINANC_RECEB_FORMA_PGTO
      frmRelRecebFormaPgto.Show

    Case ID_ITEM_REL_FINANC_FLUXO_CAIXA
      frmRelFluxo.Show

    Case ID_ITEM_REL_CP_PAGAR_DATA_VCTO
      frmRelPagar1.Show

    Case ID_ITEM_REL_CP_PAGAR_FORNECEDOR
      frmRelPagar2.Show

    Case ID_ITEM_REL_CP_PAGAR_GERAL_FILIAL
      frmRelPagar3.Show

    Case ID_ITEM_REL_CP_PAGAR_CENTRO_CUSTO
      frmRelPagar4.Show

    Case ID_ITEM_REL_CP_PAGAS_FORNECEDOR
      frmRelPagas3.Show

    Case ID_ITEM_REL_CP_PAGAS_DATA_PGTO
      frmRelPagas2.Show

    Case ID_ITEM_REL_CP_PAGAS_CENTRO_CUSTO
      frmRelPagas1.Show

    '10/05/2005 - Daniel
    'Adicionado dois novos relat�rios que
    'analisam o m�dulo de centros de custo
    Case ID_ITEM_REL_CP_CONTROLE_CENTRO_CUSTO
      'Solicitante: Carlos (OSM Consultoria)
      frmRelCentroCustoControle.Show

    Case ID_ITEM_REL_CP_CENTRO_CUSTO_COMPETENCIA
      'Solicitante: Bem Me Quer
      frmRelCentroCustoCompetencia.Show
      
    Case ID_ITEM_REL_CR_LANCAMENTOS_DATA_EMISSAO
      frmRelContasReceberPorDtEmissao.Show

    Case ID_ITEM_REL_CR_RECEBER_DATA_VCTO
      frmRelReceber1.Show

    Case ID_ITEM_REL_CR_RECEBER_CLIENTE
      frmRelReceber2.Show

    Case ID_ITEM_REL_CR_RECEBER_VENDEDOR
      frmRelRecebidas1.Show

    Case ID_ITEM_REL_CR_RECEBIDA_DATA_RECEBIMENTO
      frmRelRecebidas2.Show

    Case ID_ITEM_REL_CR_RECEBIDA_VENDEDOR
      frmRelRecebidas1.Show

    Case ID_ITEM_REL_CR_RECEBIDA_CLIENTE
      frmRelRecebidas3.Show

    Case ID_ITEM_REL_CR_CHEQUE_PRE
      frmRelCheque.Show

    Case ID_ITEM_REL_CR_CARTAO
      frmRelLancCartao.Show

    Case ID_ITEM_REL_CR_CONTA_CLIENTE
      frmRelContaCliente.Show

    Case ID_ITEM_REL_CR_EMISSAO_BOLETO
      frmImprimeBoletos.Show

    '27/09/2007 - Anderson
    'Implementado Impress�o de Carn� com C�digo de barras
    'Solicitado: Naativa
    Case ID_ITEM_REL_CR_EMISSAO_CARNET
      'frmImprimeCarnes.Show
      If g_bolCarneCodigoBarras Then
        frmImprimeCarneCodigoBarras.Show
      Else
        frmImprimeCarnes.Show
      End If

    '----------------------------------------------------------------------------
    Case ID_ITEM_REL_MALA_DIRETA
      frmImprimeMala.Show

    '20/07/2004 - Daniel
    'Altera��o: Adicionado o miRepGerarArquivoMalaDireta
    'para uso exclusivo da TV Shopping na gera��o de arquivo
    'para mala direta
    Case ID_ITEM_REL_MALA_DIRETA_GERAR_ARQUIVO
      If CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959") Then
        frmGerarArquivoMalaDireta.Show
      End If

    '24/05/2004 - Daniel
    'Case: Bic Amaz�nia
    Case ID_ITEM_REL_FOLHA_PGTO
      If CheckSerialCaseMod("QS35509-939", "QS37715-731") Then
          frmFolhaPagamento.Show
      End If

    '17/02/2004 - Daniel
    'Case: STC
    Case ID_ITEM_REL_AUTORIZACAO
      If CheckSerialCaseMod("QS39823-684") Then
          frmRelAutorizacaoPublicidade.Show
      End If

    '13/04/2004 - Daniel
    'Case: STC
    Case ID_ITEM_REL_MALA_DIRETA_AUTORIZACAO
      If CheckSerialCaseMod("QS39823-684") Then
          frmRelMalaAutorizacoes.Show
      End If

    '---------------------------------------------------------------------------- [Gr�fico]
    Case ID_ITEM_REL_GRAFICO_COMPARATIVO_CV
      frmGrafico1.Show

    Case ID_ITEM_REL_GRAFICO_VENDA_CLASSE_PERIODO
      frmGrafico2.Show

    Case ID_ITEM_REL_GRAFICO_VENDA_PRODUTO_MENSAL
      frmGrafico3.Show
      
    Case ID_ITEM_REL_GRAFICO4_VENDA_PRODUTOS
      frmGrafico4.Show
      
    Case ID_ITEM_REL_GRAFICO5_COMPRA_FORNECEDORES
      frmGrafico5.Show
      
    Case ID_ITEM_REL_GRAFICO6_VENDA_CLIENTES
      frmGrafico6.Show
      
    Case ID_ITEM_REL_EXPORTA_CLIENTES_PRODUTO
      frmPesquisaClientesProduto.Show
      
    Case ID_ITEM_REL_ESTRATEGICO_AVISO_AQUISICAO
      frmAquisicaoEstrategicoRel.Show

    '----------------------------------------------------------------------------
    '20/12/2007 - Anderson
    'Implementa��o de Relat�rio
    Case ID_ITEM_REL_NSU_CORRELACAO
      frmRelNSU.Show

    '---------------------------------------------------------------------------- [Pre�os]
    Case ID_ITEM_REL_PRECO_LISTA
      frmImprimePreco.Show

    '22/03/2004 - Daniel
    'Case: Ortociso
    'Adicionado o Rel. de Localiza��o de Produtos para otimizar
    'a rotina de busca em prateleiras ou gavetas
    Case ID_ITEM_REL_PRECO_LOCAL_PRODUTO
      frmRelLocalizacao.Show
  
    '----------------------------------------------------------------------------
  End Select
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub MDIForm_Load()
  Dim nRet As Integer
'''  Dim oQuickInfo As IQuickInfo
'''
'''  On Error GoTo ErrHandler
'''
'''  On Error Resume Next
'''
'''  Set oQuickInfo = New QuickInfoCls
'''  If Err.Number <> 0 Then
'''    gsTitle = LoadResString(201)
'''    gsMsg = "Erro: " & CStr(Err.Number) & " - Objeto QuickInfo n�o pode ser criado. Confira a instala��o do software."
'''    gnStyle = vbOKOnly + vbCritical
'''    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'''    End
'''  End If
'''  Set oQuickInfo = Nothing
'''
'''  On Error GoTo ErrHandler
  
  
  '--------------------------------------------------------------------------
  '08/08/2003 - mpdea
  'Verifica��o de inst�ncias do Quick Store
  'atrav�s de op��o em Par�metros da Filial
  '1) MDI
  '2) Login
  Me.Caption = LoadResString(5) & " " & gsMainCaption
  Call InstanceControl(Me)
  Call InstanceControl(frmLogin)
  'Descarrega form
  Unload frmLogin
  Set frmLogin = Nothing
  '--------------------------------------------------------------------------
  
  
  'Assuma com sendo Demo version. Adiante � feito check se o contr�rio
  gbDemoVersion = False
  gbProdutoRegistrado = False
  gbToAsk = False
  '
  'Nesta altura, o usu�rio atual j� est� logado e aceito pelo sistema
  'tarefa realizada pelo frmLogin...
  '
  
  'Call MountStatusBar(Me)
  
  'Abra o banco com contabiliza��o de licen�as
  'Obtenha o n�mero total delas e um dos n�meros de s�rie
  'para servir de modelo para a abertura e instala��o de arquivos
  'de conv�nios.
  nRet = gnOpenDB(gsQuickDBFileName, False, True)
  If nRet <> 0 Then
    If nRet = -1 Then
      Unload Me
      End
    Else
      If nRet = -2 Then
        gsTitle = LoadResString(201)
        gsMsg = "Este Software n�o est� registrado e houve tentativa de se estabelecer mais de uma conex�o com o Banco de Dados."
        gnStyle = vbOKOnly + vbCritical
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Unload Me
        End
      Else
        Unload Me
        End
      End If
    End If
  End If
  '
   

  '22/01/2003 - mpdea
  'Modificado chamada para InitWorld
'  Call GetGlobals
  
  
  '27/01/2009 - mpdea
  'Cria o menu
  CreateMenu CommandBars, ImageManager
  
  '29/01/2009 - mpdea
  'Seta acessos ao menu
  SetMenuAcesso
  'Call SetEnabledMenus
  
  'Call LoadBackGroundLogoX
  
  ' Flag para o comportamento duplo do bot�o "SAIR" da tela do Logon
  gbAppStarting = False
  
  Timer1.Interval = 1000
  Timer1.Enabled = True
  Timer1.Interval = 2000
   
  Exit Sub
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub MenuArquivoLogon()
  Dim nCodUser As Integer
  Dim nCodFilial As Integer
  Dim sUserName As String
  
  nCodUser = gnUserCode
  sUserName = gsUserName
  nCodFilial = gnCodFilial
  
  Call bCloseAllForms
  
  frmLogin.Show vbModal
  
  If nCodUser <> gnUserCode Or nCodFilial <> gnCodFilial Then
    frmMain.CommandBars.StatusBar.FindPane(ID_STATUSBAR_FILIAL).Text = "Filial: " & CStr(gnCodFilial)
    frmMain.CommandBars.StatusBar.FindPane(ID_STATUSBAR_USUARIO).Text = "Usu�rio: " & CStr(gnUserCode) & "-" & gsUserName
  End If
  
  '29/01/2009 - mpdea
  'Seta acessos ao menu
  SetMenuAcesso
End Sub

Private Sub MenuArquivoBackup()
  Dim nRet As Integer
  
  Screen.MousePointer = vbHourglass
  Call GetNumberOfUsers
  Screen.MousePointer = vbDefault
  
  If gnCtCurrentUsers > 1 Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "O Programa de C�pia de Seguran�a somente poder� ser rodado ap�s todas as demais esta��es em rede fecharem suas respectivas se��es."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = " Programa de C�pia de Seguran�a ser� executado a seguir. Todas as telas ser�o fechadas. Deseja prosseguir?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("")
  Call bCloseAllForms
  DoEvents
      
  db.Close
  dbFoo.Close
  ws.Close
  
  Call BackupMDB
  
  nRet = gnOpenDB(gsQuickDBFileName, False, True) + gnOpenTempDB(gsTempDBFileName, False)
  If nRet <> 0 Then
    Call bCloseAllForms
    Unload Me
  End If
  
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Reabrindo a Base de Dados...")
  Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [C�digo] <> '0' ORDER BY [C�digo Ordena��o]", dbOpenDynaset)
  Set rsProdutosNome = db.OpenRecordset("SELECT Nome, C�digo FROM Produtos WHERE [C�digo] <> '0' ORDER BY Nome", dbOpenDynaset)
  frmProdutos.Form_Load
  frmProdutos.Hide
  Call StatusMsg("")
  Screen.MousePointer = vbDefault

End Sub

Private Sub MenuArquivoExportarBase()

  Screen.MousePointer = vbHourglass
  Call GetNumberOfUsers
  Screen.MousePointer = vbDefault
  
  If gnCtCurrentUsers > 1 Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "A exporta��o da base atual somente poder� ser feita ap�s todas as demais esta��es em rede fecharem suas respectivas se��es."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "A exporta��o da base atual ser� realizada a seguir. Deseja prosseguir?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    Call StatusMsg(LoadResString(254))
    Call ExportMDB
    Call StatusMsg("")
  End If

End Sub

Private Sub MenuArquivoRepararBase()
  Dim nRet As Integer

  Screen.MousePointer = vbHourglass
  Call GetNumberOfUsers
  Screen.MousePointer = vbDefault
  
  If gnCtCurrentUsers > 1 Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "A repara��o da base atual somente poder� ser feita ap�s todas as demais esta��es em rede fecharem suas respectivas se��es."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "A repara��o da base atual ser� realizada a seguir. Todas as telas ser�o fechadas. Deseja prosseguir?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    Screen.MousePointer = vbHourglass
    Call StatusMsg("")
    Call bCloseAllForms
    DoEvents
    
    Call StatusMsg(LoadResString(251))
    db.Close
    dbFoo.Close
    ws.Close
    DBEngine.RepairDatabase gsQuickDBFileName
    nRet = gnOpenDB(gsQuickDBFileName, False, True) + gnOpenTempDB(gsTempDBFileName, False)
    If nRet <> 0 Then
      Call bCloseAllForms
      Unload Me
    End If
    Call StatusMsg("")
    Screen.MousePointer = vbDefault
    gsTitle = "Fim de Opera��o"
    gsMsg = "Base de Dados reparada com sucesso."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Screen.MousePointer = vbHourglass
    Call StatusMsg("Reabrindo a Base de Dados...")
    Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [C�digo] <> '0' ORDER BY [C�digo Ordena��o]", dbOpenDynaset)
    Set rsProdutosNome = db.OpenRecordset("SELECT Nome, C�digo FROM Produtos WHERE [C�digo] <> '0' ORDER BY Nome", dbOpenDynaset)
    frmProdutos.Form_Load
    frmProdutos.Hide
    Call StatusMsg("")
    Screen.MousePointer = vbDefault
  End If

End Sub

Private Sub MenuArquivoCompactarBase()
  Dim sTempFileName As String
  Dim nRet As Integer

  Screen.MousePointer = vbHourglass
  Call GetNumberOfUsers
  Screen.MousePointer = vbDefault
  
  If gnCtCurrentUsers > 1 Then
    Beep
    gsTitle = LoadResString(201)
    gsMsg = "A compacta��o da base atual somente poder� ser feita ap�s todas as demais esta��es em rede fecharem suas respectivas se��es."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "A compacta��o da base atual ser� realizada a seguir. Todas as telas ser�o fechadas. Deseja prosseguir?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbYes Then
    Screen.MousePointer = vbHourglass
    Call StatusMsg("")
    Call bCloseAllForms
    DoEvents
    
    Screen.MousePointer = vbHourglass
    Call StatusMsg(LoadResString(252))
    db.Close
    dbFoo.Close
    ws.Close
    sTempFileName = App.Path & "\TMP" & Format(Time, "HHMMSS") & ".MDB"
    On Error Resume Next
    Kill sTempFileName
    On Error GoTo 0
    DBEngine.CompactDatabase gsQuickDBFileName, sTempFileName, , , ";pwd=" & gsGetPValue()
    Kill gsQuickDBFileName
    Name sTempFileName As gsQuickDBFileName
    nRet = gnOpenDB(gsQuickDBFileName, False, True) + gnOpenTempDB(gsTempDBFileName, False)
    If nRet <> 0 Then
      Call bCloseAllForms
      Unload Me
    End If
    Call StatusMsg("")
    
    gsTitle = "Fim de Opera��o"
    gsMsg = "Base de Dados compactada com sucesso."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Screen.MousePointer = vbDefault
    Screen.MousePointer = vbHourglass
    Call StatusMsg("Reabrindo a Base de Dados...")
    Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [C�digo] <> '0' ORDER BY [C�digo Ordena��o]", dbOpenDynaset)
    Set rsProdutosNome = db.OpenRecordset("SELECT Nome, C�digo FROM Produtos WHERE [C�digo] <> '0' ORDER BY Nome", dbOpenDynaset)
    frmProdutos.Form_Load
    frmProdutos.Hide
    Call StatusMsg("")
    Screen.MousePointer = vbDefault
  End If
      
End Sub

Private Sub MenuCadastroClasse(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 9999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  'Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o] FROM Classes ORDER BY C�digo", dbOpenDynaset)
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o], LucroMinimoPermitido FROM Classes ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster
  '19/10/2007 - Anderson
  'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
  If g_bolLucroMinimoClasse Then
    F.grdCodDesc.Columns(2).Visible = True
    F.grdCodDesc.Columns(1).Width = 4395
  End If

End Sub

Private Sub MenuCadastroSubclasse(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 9999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  '19/10/2007 - Anderson
  'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
  F.grdCodDesc.Columns(2).Visible = False
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o] FROM [Sub Classes] ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroCor(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o] FROM Cores ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroTamanho(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o] FROM Tamanhos ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroMalaDiretaGrupo(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome FROM [Grupos Interesse] ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroPesquisa(ByVal strCaption As String, ByVal intPesquisa As Integer)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 99999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o] FROM [Pesquisa " & intPesquisa & "] ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroBanco(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome FROM Bancos ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroCaixa(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 99
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Caixa, Descri��o FROM [Caixas em Uso] ORDER BY Caixa", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroMoeda(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 99
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome FROM Moedas ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

'--------------------------------------------------------------------------
'26/01/2006 - mpdea
'Inclus�o dos Cadastros de Grupo Fiscal e Mensagens
'--------------------------------------------------------------------------
Private Sub MenuCadastroGrupoFiscal(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 9999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o] FROM GrupoFiscal ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroClassificacaoFiscal(ByVal strCaption As String)
  Dim F As Form

  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  '02/06/2008 - mpdea
  'Modificado o c�digo m�ximo de 99 para 255
  'F.gnMaxCod = 255
  '14/05/2010 - Andrea
  'Modificado o c�digo m�ximo de 255 para 32767 (inteiro)
  F.gnMaxCod = 32767
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome FROM [Classifica��o Fiscal] ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

Private Sub MenuCadastroCentroCusto(ByVal strCaption As String)
  Dim F As Form

  '05/05/2005 - Daniel
  'Adi��es de objetos que tratar�o rotinas para
  'ativa��o e desativa��o de Centro(s) de Custo
  '
  'Projeto: Melhorias para o Centro de Custo
  
  If gbShowWindow(gsSupressSpecialChars(strCaption)) Then
    Exit Sub
  End If
  
  Set F = New frmTabela
  F.gnMaxCod = 999
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  '05/05/2005 - Daniel
  'Projeto: Melhorias para o Centro de Custo
  'Adicionado a cl�usula [WHERE] no SELECT
  'e adicionamos o campo Ativo na sele��o
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o], Ativo FROM [Centros de Custo] WHERE Ativo = TRUE ORDER BY C�digo", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster
  
  '05/05/2005 - Daniel
  'Tratamento para os bot�es: miOpDesativarCentro, miOpAtivarCentros,
  'miOpReativarCentroIndividualmente e miOpRefresh vis�veis apenas para o Centro de Custo
  F.ActiveBar1.Tools("miOpDesativarCentro").Visible = True
  F.ActiveBar1.Tools("miOpAtivarCentros").Visible = True
  F.ActiveBar1.Tools("miOpReativarCentroIndividualmente").Visible = True
  F.ActiveBar1.Tools("miOpRefresh").Visible = True

End Sub

Private Sub MenuCadastroCaracteristicaClienteFornecedor()
  Dim F As Form

  If Not frmGerente.gbSenhaGerente Then
    Exit Sub
  End If
  Set F = New frmCliForCaract
  F.Show

End Sub

Private Sub MenuGeradorLayout(ByVal intTipo As Integer, Optional ByVal strTipo As String = "")
  Dim F As Form

  Set F = New frmLayoutGen
  F.gnTypeDoc = intTipo
  F.gsTypeDoc = strTipo
  F.Show

End Sub

Private Sub BackupMDB()
  Dim nWinState As FormWindowStateConstants
  
  On Error GoTo ErrHandler
  nWinState = frmMain.WindowState
  frmMain.WindowState = vbMinimized
  Call ExecCmd(gsBackupFileName)
  frmMain.WindowState = nWinState
  On Error GoTo 0
  Exit Sub
 
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Software Backup para o Quick Store n�o encontrado como: " & gsBackupFileName
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

End Sub

Private Sub ExportMDB()
  Dim nPos As Integer
  Dim sNewDBName As String
  Dim dbExport As Database
  Dim nRet As Integer
  
  On Error Resume Next
  
  Call StatusMsg("")
  Call bCloseAllForms
  
  Screen.MousePointer = vbHourglass
  DBEngine.Idle dbRefreshCache
  nPos = InStr(gsQuickDBFileName, ".")
  If nPos = 0 Then
    nPos = Len(gsQuickDBFileName)
  End If
  sNewDBName = Mid(gsQuickDBFileName, 1, nPos - 1) & "-" & Format(Date, "yyyymmdd") & ".mdb"
  On Error Resume Next
  Kill sNewDBName
  On Error GoTo 0
  db.Close
  dbFoo.Close
  ws.Close
  DoEvents
  Call FileCopy(gsQuickDBFileName, sNewDBName)
  Set ws = DBEngine.Workspaces(0)
  Call ws.BeginTrans
  Set dbExport = ws.OpenDatabase(sNewDBName, True, False, ";pwd=" & gsGetPValue())
  On Error Resume Next
  Call dbExport.NewPassword(gsGetPValue(), "")
  On Error GoTo 0
  
  Call dbExport.Execute("DROP TABLE ZZZ", dbFailOnError)
  Call dbExport.Execute("DROP TABLE Reports", dbFailOnError)
  Call dbExport.Execute("DROP TABLE ZZZProgramas", dbFailOnError)
  Call dbExport.Execute("DROP TABLE Acessos", dbFailOnError)
  '08/08/2005 - Daniel
  'Antes de executar a query "ALTER TABLE Funcion�rios DROP COLUMN Senha"
  'estaremos eliminando o Index "Acessando" para corre��o do Erro 3280
  '
  Call dbExport.Execute("DROP INDEX Acessando ON Funcion�rios", dbFailOnError)
  '
  Call dbExport.Execute("ALTER TABLE Funcion�rios DROP COLUMN Senha", dbFailOnError)
  Call dbExport.Execute("ALTER TABLE Funcion�rios DROP COLUMN ValorP", dbFailOnError)
  Call dbExport.Execute("ALTER TABLE [Par�metros Filial] DROP COLUMN [Senha Gerente]", dbFailOnError)
  
  dbExport.Close
  Set dbExport = Nothing
  Call ws.CommitTrans
  ws.Close
  
  Screen.MousePointer = vbDefault
  
  gsTitle = LoadResString(201)
  gsMsg = "Base de Dados exportada com sucesso. Nome do arquivo gerado: " & sNewDBName
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  nRet = gnOpenDB(gsQuickDBFileName, False, True)
  If nRet <> 0 Then
    If nRet = -1 Then
      Unload Me
      End
    Else
      If nRet = -2 Then
        gsTitle = LoadResString(201)
        gsMsg = "Este Software n�o est� registrado e houve tentativa de se estabelecer mais de uma conex�o com o Banco de Dados."
        gnStyle = vbOKOnly + vbCritical
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Unload Me
        End
      Else
        Unload Me
        End
      End If
    End If
  End If
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Reabrindo a Base de Dados...")
  Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [C�digo] <> '0' ORDER BY [C�digo Ordena��o]", dbOpenDynaset)
  Set rsProdutosNome = db.OpenRecordset("SELECT Nome, C�digo FROM Produtos WHERE [C�digo] <> '0' ORDER BY Nome", dbOpenDynaset)
  Call StatusMsg("")
  Screen.MousePointer = vbDefault

End Sub

Private Sub ChamaGerador(ByVal sExeFileName As String)
  Dim nWinState As FormWindowStateConstants
  
  On Error GoTo ErrHandler
  nWinState = frmMain.WindowState
  frmMain.WindowState = vbMinimized
  Call ExecCmd(sExeFileName)
  frmMain.WindowState = nWinState
  On Error GoTo 0
  Exit Sub
 
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Software Gerador n�o encontrado como: " & sExeFileName
  gsMsg = gsMsg & vbCrLf & "Revise a instala��o do software."
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call CheckPendencias
  If gbToAsk = True Or UnloadMode = vbFormControlMenu Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(200)
    gnStyle = vbYesNo + vbQuestion
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbYes Then
      Cancel = False
    Else
      Cancel = True
    End If
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde, fechando o Banco de Dados...")
  Call bCloseAllForms
  
  '07/08/2003 - mpdea
  'Fecha conex�o com a base de dados
  db.Close
  ws.Close
  Set db = Nothing
  Set ws = Nothing
  
  ' Fechar conex�o com o banco de dados SQL SERVER
  gnCloseDB_SQLSERVER
  
  Set soapclient = Nothing
  Set soapclient_NFCe = Nothing
  
  DBEngine.Idle dbRefreshCache
  
  '14/06/2006 - mpdea
  'Verifica porta COM aberta
  If MSComm1.PortOpen Then MSComm1.PortOpen = False
  
  Call WaitSeconds(3)
  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  End
End Sub

Private Sub CheckPendencias()
  Dim Texto_Entrada As String
  Dim Texto_Sa�da As String
  Dim Aux_Sequ�ncia As Long
  Dim Texto As String
  
  Dim TB_Entradas As Recordset
  Dim TB_Sa�das As Recordset
  Dim TB_Op_Entradas As Recordset
  Dim TB_Op_Sa�das As Recordset
  
  
  Texto_Entrada = ""
  Texto_Sa�da = ""
  
  Call StatusMsg("Aguarde, verificando pend�ncias...")
  
  Set TB_Entradas = db.OpenRecordset("Entradas", , dbReadOnly)
  Set TB_Sa�das = db.OpenRecordset("Sa�das", , dbReadOnly)
  Set TB_Op_Entradas = db.OpenRecordset("Opera��es Entrada", , dbReadOnly)
  Set TB_Op_Sa�das = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
  
  
  Rem Verifica se existem entradas, que n�o seja pedidos, pendentes
  Aux_Sequ�ncia = 0
  TB_Entradas.Index = "Data"
  TB_Op_Entradas.Index = "C�digo"
Lp1:
  TB_Entradas.Seek ">", gnCodFilial, Data_Atual, Aux_Sequ�ncia
  If TB_Entradas.NoMatch Then GoTo V�_Sa�das
  If TB_Entradas("Filial") <> gnCodFilial Then GoTo V�_Sa�das
  If TB_Entradas("Data") <> Data_Atual Then GoTo V�_Sa�das
  Aux_Sequ�ncia = TB_Entradas("Sequ�ncia")
  
  TB_Op_Entradas.Seek "=", TB_Entradas("Opera��o")
  If TB_Op_Entradas.NoMatch Then GoTo Lp1
  If TB_Op_Entradas("Tipo") = "P" Then GoTo Lp1
  If TB_Entradas("Efetivada") = False Then
    Texto_Entrada = Texto_Entrada + " - " + str(TB_Entradas("Sequ�ncia"))
  End If
  GoTo Lp1
  
  
V�_Sa�das:
  Rem Verifica se existem sa�das, que n�o sejam or�amentos, pendentes
  Aux_Sequ�ncia = 0
  TB_Sa�das.Index = "Data"
  TB_Op_Sa�das.Index = "C�digo"
LP2:
  TB_Sa�das.Seek ">", gnCodFilial, Data_Atual, Aux_Sequ�ncia
  If TB_Sa�das.NoMatch Then GoTo Fim
  If TB_Sa�das("Filial") <> gnCodFilial Then GoTo Fim
  If TB_Sa�das("Data") <> Data_Atual Then GoTo Fim
  Aux_Sequ�ncia = TB_Sa�das("Sequ�ncia")
  
  TB_Op_Sa�das.Seek "=", TB_Sa�das("Opera��o")
  If TB_Op_Sa�das.NoMatch Then GoTo LP2
  If TB_Op_Sa�das("Tipo") = "O" Then GoTo LP2
  If TB_Sa�das("Efetivada") = False Then
    Texto_Sa�da = Texto_Sa�da + " - " + str(TB_Sa�das("Sequ�ncia"))
  End If
  GoTo LP2
  
Fim:
  
  Call StatusMsg("")
  
  If Texto_Entrada <> "" Or Texto_Sa�da <> "" Then
    Texto = "Exitem movimenta��es de entrada e/ou sa�das que n�o foram efetivadas."
    Texto = Texto + Chr(13) + "SE ESTAS MOVIMENTA��ES N�O FOREM EFETIVADAS HOJE, SEU ESTOQUE, CAIXA, COMISS�ES E OUTRAS INFORMA��ES FICAR�O INCORRETAS."
    Texto = Texto + Chr(13) + "� RECOMEND�VEL QUE VOC� EFETIVE ESTAS MOVIMENTA��ES ANTES DE SAIR."
    Texto = Texto + Chr(13) + Chr(13)
    If Texto_Entrada <> "" Then
      Texto = Texto + "Movimenta��es de entrada n�o efetivadas " + Texto_Entrada
      Texto = Texto + Chr(13)
    End If
    If Texto_Sa�da <> "" Then
      Texto = Texto + "Movimenta��es de sa�da n�o efetivadas " + Texto_Sa�da
      Texto = Texto + Chr(13)
    End If
    
    Texto = Texto + Chr(13) + "Para efetivar estas movimenta��es siga os passos abaixo: "
    Texto = Texto + Chr(13) + Chr(13) + "Opera��es de Entrada : Use a tela de entrada, encontre as movimenta��es e grave-as novamente."
    Texto = Texto + Chr(13) + "Opera��es de Sa�da : Use a tela de sa�das ou a tela de venda r�pida, encontre as movimenta��es, grave-as novamente e FA�A O RECEBIMENTO."
    
    frmPendencias.Mensagem.Caption = Texto
    frmPendencias.Show vbModal

  End If

End Sub

Private Sub Timer1_Timer()
  Dim nMins As Integer
  Dim sTime As String
  Static nMinsAnt As Integer
  
  On Error GoTo ErrHandler
  
  If Not gbLoginDone Then
    Exit Sub
  End If
  
  'Tips
  If gsTipFile <> "" Then
    If Dir(gsTipFile) <> "" Then
      If CInt(GetSetting("QuickStore", "Options", "Show Tips", 1)) <> 0 Then
        frmTip.Show
      End If
      gsTipFile = ""
    End If
  End If
  
  If IsProdutoRegistrado() Then
    Timer1.Enabled = False
    Exit Sub
  End If
  
  If Weekday(Date) > 1 And Weekday(Date) < 7 Then
    sTime = Format(Time(), "hhmm")
    If (sTime >= Format("08:30", "hhmm") And _
        sTime <= Format("11:30", "hhmm")) Or _
       (sTime >= Format("13:30", "hhmm") And _
        sTime <= Format("18:00", "hhmm")) Then
      nMins = Minute(Now)
      If nMins <> nMinsAnt Then
        Timer1.Enabled = False
        If (nMins Mod 5) = 0 Then
          nMinsAnt = nMins
          Load frmDeveRegistrar
          frmDeveRegistrar.ZOrder 0
          frmDeveRegistrar.gsPrefix = "QS"
          frmDeveRegistrar.Show vbModal
        End If
      End If
    End If
  End If
  
  Exit Sub
  
ErrHandler:
  If Err.Number = 401 Then
    Exit Sub
  Else
    Screen.MousePointer = vbDefault
    Call StatusMsg("")
    Select Case frmErro.gnShowErr(Err.Number, "")
      Case 0 'Repetir
        Resume
      Case 1 'Prosseguir
        Resume Next
      Case 2 'Sair
        Exit Sub
      Case 3 'Encerrar
        End
    End Select
  End If
End Sub

Private Sub Timer2_Timer()
   frm_manterAtivo.Show 1
End Sub
