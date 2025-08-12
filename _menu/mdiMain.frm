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
                                             " [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenSnapshot)
        With rstParametros
          If .Fields("Verifica Agenda") = True Then
            Call Verifica_Pendências
            If frmAgenda.lstPend.ListCount > 0 Then
              frmAgenda.Show vbModal
            End If
          End If
          .Close
          Set rstParametros = Nothing
        End With

    '----------------------------------------------------------------------------
    'TAB Início
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
    'Adicionado Parâmetro: Classificação de Clientes
    'Case: TV Shopping
    Case ID_ITEM_INICIO_PARAM_CLASS_CLIENTE
      frmClassificacaoClientes.Show

    '02/08/2004 - Daniel
    'Adicionado Parâmetro: Faturamento Automático
    'Case: STC de Caxias do Sul
    Case ID_ITEM_INICIO_PARAM_FATURAMENTO_AUTO
      frmParamFaturameAuto.Show

    '15/09/2004 - Daniel
    'Adicionado Parâmetro: Configuração de Saídas para a Devolução de Materiais
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
    'Inclusão de Cadastro de Rádio e Tipo Comercial
    'otimizado inicialmente para a STC de Caxias do Sul
    '23/07/2004 - Daniel
    'Inclusão do Cadastro de Autorizações
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
      'Cadastro de Código NBM (Codificação da Nomenclatura Brasileira de Mercadorias)
      'Impacto no resgistro 75 da geração do arquivo para o SEF
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
      
        'Aproveitamento de condição case para nova função chama site...
        Dim iret As Long
        iret = ShellExecute(Me.hwnd, vbNullString, gPastaRetornoNfe, vbNullString, "c:\", 1)


    '26/02/2004 - Daniel
    'Case: PSV - Manutenção de Reservas
    Case ID_ITEM_MOVIMENTO_MANUT_RESERVA
      If CheckSerialCaseMod("QS35552-811", "QS37705-639", "QS37825-830", "QS38933-772", "QS39369-521") Then
        frmManutencaoReservas.Show
      End If

    Case ID_ITEM_MOVIMENTO_MANUT_CONSIG_ENTRADA
      frmManConsigEntradas.Show

    '13/08/2004 - Daniel
    'Faturamento Automático - Case: STC
    Case ID_ITEM_MOVIMENTO_FATUR_AUTO
      frmFaturamentoAutomatico.Show

    '17/09/2004 - Daniel
    'Prestação de Contas com Fornecedores - Case: Resultado
    Case ID_ITEM_MOVIMENTO_PREST_FORNEC
      frmManPrestacaodeContas.Show

    '24/01/2005 - Daniel
    'Importador para a Castro Construções
    Case ID_ITEM_MOVIMENTO_IMPORTACAO
      frmImportacao.Show

    '28/10/2002 - mpdea
    'Inclusão de novas telas
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
    'TAB Preços
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
    'Verifica alteração personalizada
    '
    'QS32815-683 = Guarantã
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
    'Verifica alteração personalizada
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
      'Parâmetro para chamar o programa InfoICMS somente através do Quick Store
      Call Shell(gsGeradorRecEstadual & " QuickStore_InfoICMS", vbNormalFocus)
    
    '----------------------------------------------------------------------------
    'TAB Relatórios
    '----------------------------------------------------------------------------
    '---------------------------------------------------------------------------- [Serviços]
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
    'Relatório de Estoque das Filiais e Preço (Personalizado)
    'Solicitante: Cliente Kilouça (QS71271-970)
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
    'Customização de Relatório de Vendas para LL Comércio de Ferramentas LTDA.
    Case ID_ITEM_REL_CV_VENDA_2
      frmRelVendasII.Show

    Case ID_ITEM_REL_CV_COMISSAO
      frmRelComissoes.Show

    '26/04/2005 - Daniel
    'Adicionado relatório de comissões que exibe às
    'retenções contidas sobre cartões
    'Solicitante: Bem Me Quer
    Case ID_ITEM_REL_CV_COMISSAO_RETENCAO
      frmRelComissaoComRetencao.Show

    Case ID_ITEM_REL_CV_COMISSAO_VENDA_VENDEDOR
      frmRelVendaComissao.Show

    Case ID_ITEM_REL_CV_COMPRAS
      frmRelCompras.Show

    '30/10/2007 - Anderson
    'Relatório de Produtos a comprar
    'Solicitante: Kings Cross
    Case ID_ITEM_REL_CV_PRODUTO_COMPRAR_FATOR
      frmRelProdutosAComprar.Show

    Case ID_ITEM_REL_CV_VENDA_CLIENTE
      frmRelVendasPorCliente.Show

    '30/03/2005 - Daniel
    'Relatório de Vendedores e Comissões (Sintético)
    'Solicitação: Bem Me Quer
    'Disponível para todos clientes do Quick Store
    Case ID_ITEM_REL_CV_VENDA_VENDEDOR_COMISSAO
      frmRelVendedorComissoesSintetico.Show

    Case ID_ITEM_REL_CV_VENDA_TAMANHO
      frmRelVendasTamanho.Show

    Case ID_ITEM_REL_CV_VENDA_EDITORA
      frmRelEditoras.Show

    '17/08/2007 - Anderson
    'Alteração realizada para atender solicitação da Nutricare (QS73086-490)
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
    'Adicionado Rel. de Cartões de Crédito (Posição Diária)
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
    'Inclusão de relatório Recebimentos por Forma de Pagamento
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
    'Adicionado dois novos relatórios que
    'analisam o módulo de centros de custo
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
    'Implementado Impressão de Carnê com Código de barras
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
    'Alteração: Adicionado o miRepGerarArquivoMalaDireta
    'para uso exclusivo da TV Shopping na geração de arquivo
    'para mala direta
    Case ID_ITEM_REL_MALA_DIRETA_GERAR_ARQUIVO
      If CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959") Then
        frmGerarArquivoMalaDireta.Show
      End If

    '24/05/2004 - Daniel
    'Case: Bic Amazônia
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

    '---------------------------------------------------------------------------- [Gráfico]
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
    'Implementação de Relatório
    Case ID_ITEM_REL_NSU_CORRELACAO
      frmRelNSU.Show

    '---------------------------------------------------------------------------- [Preços]
    Case ID_ITEM_REL_PRECO_LISTA
      frmImprimePreco.Show

    '22/03/2004 - Daniel
    'Case: Ortociso
    'Adicionado o Rel. de Localização de Produtos para otimizar
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
'''    gsMsg = "Erro: " & CStr(Err.Number) & " - Objeto QuickInfo não pode ser criado. Confira a instalação do software."
'''    gnStyle = vbOKOnly + vbCritical
'''    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
'''    End
'''  End If
'''  Set oQuickInfo = Nothing
'''
'''  On Error GoTo ErrHandler
  
  
  '--------------------------------------------------------------------------
  '08/08/2003 - mpdea
  'Verificação de instâncias do Quick Store
  'através de opção em Parâmetros da Filial
  '1) MDI
  '2) Login
  Me.Caption = LoadResString(5) & " " & gsMainCaption
  Call InstanceControl(Me)
  Call InstanceControl(frmLogin)
  'Descarrega form
  Unload frmLogin
  Set frmLogin = Nothing
  '--------------------------------------------------------------------------
  
  
  'Assuma com sendo Demo version. Adiante é feito check se o contrário
  gbDemoVersion = False
  gbProdutoRegistrado = False
  gbToAsk = False
  '
  'Nesta altura, o usuário atual já está logado e aceito pelo sistema
  'tarefa realizada pelo frmLogin...
  '
  
  'Call MountStatusBar(Me)
  
  'Abra o banco com contabilização de licenças
  'Obtenha o número total delas e um dos números de série
  'para servir de modelo para a abertura e instalação de arquivos
  'de convênios.
  nRet = gnOpenDB(gsQuickDBFileName, False, True)
  If nRet <> 0 Then
    If nRet = -1 Then
      Unload Me
      End
    Else
      If nRet = -2 Then
        gsTitle = LoadResString(201)
        gsMsg = "Este Software não está registrado e houve tentativa de se estabelecer mais de uma conexão com o Banco de Dados."
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
  
  ' Flag para o comportamento duplo do botão "SAIR" da tela do Logon
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
    frmMain.CommandBars.StatusBar.FindPane(ID_STATUSBAR_USUARIO).Text = "Usuário: " & CStr(gnUserCode) & "-" & gsUserName
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
    gsMsg = "O Programa de Cópia de Segurança somente poderá ser rodado após todas as demais estações em rede fecharem suas respectivas seções."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = " Programa de Cópia de Segurança será executado a seguir. Todas as telas serão fechadas. Deseja prosseguir?"
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
  Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
  Set rsProdutosNome = db.OpenRecordset("SELECT Nome, Código FROM Produtos WHERE [Código] <> '0' ORDER BY Nome", dbOpenDynaset)
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
    gsMsg = "A exportação da base atual somente poderá ser feita após todas as demais estações em rede fecharem suas respectivas seções."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "A exportação da base atual será realizada a seguir. Deseja prosseguir?"
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
    gsMsg = "A reparação da base atual somente poderá ser feita após todas as demais estações em rede fecharem suas respectivas seções."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "A reparação da base atual será realizada a seguir. Todas as telas serão fechadas. Deseja prosseguir?"
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
    gsTitle = "Fim de Operação"
    gsMsg = "Base de Dados reparada com sucesso."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Screen.MousePointer = vbHourglass
    Call StatusMsg("Reabrindo a Base de Dados...")
    Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
    Set rsProdutosNome = db.OpenRecordset("SELECT Nome, Código FROM Produtos WHERE [Código] <> '0' ORDER BY Nome", dbOpenDynaset)
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
    gsMsg = "A compactação da base atual somente poderá ser feita após todas as demais estações em rede fecharem suas respectivas seções."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "A compactação da base atual será realizada a seguir. Todas as telas serão fechadas. Deseja prosseguir?"
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
    
    gsTitle = "Fim de Operação"
    gsMsg = "Base de Dados compactada com sucesso."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Screen.MousePointer = vbDefault
    Screen.MousePointer = vbHourglass
    Call StatusMsg("Reabrindo a Base de Dados...")
    Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
    Set rsProdutosNome = db.OpenRecordset("SELECT Nome, Código FROM Produtos WHERE [Código] <> '0' ORDER BY Nome", dbOpenDynaset)
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
  'Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração] FROM Classes ORDER BY Código", dbOpenDynaset)
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração], LucroMinimoPermitido FROM Classes ORDER BY Código", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster
  '19/10/2007 - Anderson
  'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
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
  'Implementação do campo Lucro Mínimo Permitido conforme solicitação da Agrotama
  F.grdCodDesc.Columns(2).Visible = False
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração] FROM [Sub Classes] ORDER BY Código", dbOpenDynaset)
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração] FROM Cores ORDER BY Código", dbOpenDynaset)
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração] FROM Tamanhos ORDER BY Código", dbOpenDynaset)
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome FROM [Grupos Interesse] ORDER BY Código", dbOpenDynaset)
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração] FROM [Pesquisa " & intPesquisa & "] ORDER BY Código", dbOpenDynaset)
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome FROM Bancos ORDER BY Código", dbOpenDynaset)
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Caixa, Descrição FROM [Caixas em Uso] ORDER BY Caixa", dbOpenDynaset)
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome FROM Moedas ORDER BY Código", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster

End Sub

'--------------------------------------------------------------------------
'26/01/2006 - mpdea
'Inclusão dos Cadastros de Grupo Fiscal e Mensagens
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
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração] FROM GrupoFiscal ORDER BY Código", dbOpenDynaset)
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
  'Modificado o código máximo de 99 para 255
  'F.gnMaxCod = 255
  '14/05/2010 - Andrea
  'Modificado o código máximo de 255 para 32767 (inteiro)
  F.gnMaxCod = 32767
  F.Show
  F.Caption = gsSupressSpecialChars(strCaption)
  F.grdCodDesc.Caption = "Lista de " & F.Caption
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome FROM [Classificação Fiscal] ORDER BY Código", dbOpenDynaset)
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
  'Adições de objetos que tratarão rotinas para
  'ativação e desativação de Centro(s) de Custo
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
  'Adicionado a cláusula [WHERE] no SELECT
  'e adicionamos o campo Ativo na seleção
  Set F.datMaster.Recordset = db.OpenRecordset("SELECT Código, Nome, [Data Alteração], Ativo FROM [Centros de Custo] WHERE Ativo = TRUE ORDER BY Código", dbOpenDynaset)
  If Not F.datMaster.Recordset.EOF Then
    F.datMaster.Recordset.MoveLast
    F.datMaster.Recordset.MoveFirst
  End If
  F.datMaster.Refresh
  Set F.grdCodDesc.DataSource = F.datMaster
  
  '05/05/2005 - Daniel
  'Tratamento para os botões: miOpDesativarCentro, miOpAtivarCentros,
  'miOpReativarCentroIndividualmente e miOpRefresh visíveis apenas para o Centro de Custo
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
  gsMsg = "Software Backup para o Quick Store não encontrado como: " & gsBackupFileName
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
  'Antes de executar a query "ALTER TABLE Funcionários DROP COLUMN Senha"
  'estaremos eliminando o Index "Acessando" para correção do Erro 3280
  '
  Call dbExport.Execute("DROP INDEX Acessando ON Funcionários", dbFailOnError)
  '
  Call dbExport.Execute("ALTER TABLE Funcionários DROP COLUMN Senha", dbFailOnError)
  Call dbExport.Execute("ALTER TABLE Funcionários DROP COLUMN ValorP", dbFailOnError)
  Call dbExport.Execute("ALTER TABLE [Parâmetros Filial] DROP COLUMN [Senha Gerente]", dbFailOnError)
  
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
        gsMsg = "Este Software não está registrado e houve tentativa de se estabelecer mais de uma conexão com o Banco de Dados."
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
  Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
  Set rsProdutosNome = db.OpenRecordset("SELECT Nome, Código FROM Produtos WHERE [Código] <> '0' ORDER BY Nome", dbOpenDynaset)
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
  gsMsg = "Software Gerador não encontrado como: " & sExeFileName
  gsMsg = gsMsg & vbCrLf & "Revise a instalação do software."
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
  'Fecha conexão com a base de dados
  db.Close
  ws.Close
  Set db = Nothing
  Set ws = Nothing
  
  ' Fechar conexão com o banco de dados SQL SERVER
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
  Dim Texto_Saída As String
  Dim Aux_Sequência As Long
  Dim Texto As String
  
  Dim TB_Entradas As Recordset
  Dim TB_Saídas As Recordset
  Dim TB_Op_Entradas As Recordset
  Dim TB_Op_Saídas As Recordset
  
  
  Texto_Entrada = ""
  Texto_Saída = ""
  
  Call StatusMsg("Aguarde, verificando pendências...")
  
  Set TB_Entradas = db.OpenRecordset("Entradas", , dbReadOnly)
  Set TB_Saídas = db.OpenRecordset("Saídas", , dbReadOnly)
  Set TB_Op_Entradas = db.OpenRecordset("Operações Entrada", , dbReadOnly)
  Set TB_Op_Saídas = db.OpenRecordset("Operações Saída", , dbReadOnly)
  
  
  Rem Verifica se existem entradas, que não seja pedidos, pendentes
  Aux_Sequência = 0
  TB_Entradas.Index = "Data"
  TB_Op_Entradas.Index = "Código"
Lp1:
  TB_Entradas.Seek ">", gnCodFilial, Data_Atual, Aux_Sequência
  If TB_Entradas.NoMatch Then GoTo Vê_Saídas
  If TB_Entradas("Filial") <> gnCodFilial Then GoTo Vê_Saídas
  If TB_Entradas("Data") <> Data_Atual Then GoTo Vê_Saídas
  Aux_Sequência = TB_Entradas("Sequência")
  
  TB_Op_Entradas.Seek "=", TB_Entradas("Operação")
  If TB_Op_Entradas.NoMatch Then GoTo Lp1
  If TB_Op_Entradas("Tipo") = "P" Then GoTo Lp1
  If TB_Entradas("Efetivada") = False Then
    Texto_Entrada = Texto_Entrada + " - " + str(TB_Entradas("Sequência"))
  End If
  GoTo Lp1
  
  
Vê_Saídas:
  Rem Verifica se existem saídas, que não sejam orçamentos, pendentes
  Aux_Sequência = 0
  TB_Saídas.Index = "Data"
  TB_Op_Saídas.Index = "Código"
LP2:
  TB_Saídas.Seek ">", gnCodFilial, Data_Atual, Aux_Sequência
  If TB_Saídas.NoMatch Then GoTo Fim
  If TB_Saídas("Filial") <> gnCodFilial Then GoTo Fim
  If TB_Saídas("Data") <> Data_Atual Then GoTo Fim
  Aux_Sequência = TB_Saídas("Sequência")
  
  TB_Op_Saídas.Seek "=", TB_Saídas("Operação")
  If TB_Op_Saídas.NoMatch Then GoTo LP2
  If TB_Op_Saídas("Tipo") = "O" Then GoTo LP2
  If TB_Saídas("Efetivada") = False Then
    Texto_Saída = Texto_Saída + " - " + str(TB_Saídas("Sequência"))
  End If
  GoTo LP2
  
Fim:
  
  Call StatusMsg("")
  
  If Texto_Entrada <> "" Or Texto_Saída <> "" Then
    Texto = "Exitem movimentações de entrada e/ou saídas que não foram efetivadas."
    Texto = Texto + Chr(13) + "SE ESTAS MOVIMENTAÇÕES NÃO FOREM EFETIVADAS HOJE, SEU ESTOQUE, CAIXA, COMISSÕES E OUTRAS INFORMAÇÕES FICARÃO INCORRETAS."
    Texto = Texto + Chr(13) + "É RECOMENDÁVEL QUE VOCÊ EFETIVE ESTAS MOVIMENTAÇÕES ANTES DE SAIR."
    Texto = Texto + Chr(13) + Chr(13)
    If Texto_Entrada <> "" Then
      Texto = Texto + "Movimentações de entrada não efetivadas " + Texto_Entrada
      Texto = Texto + Chr(13)
    End If
    If Texto_Saída <> "" Then
      Texto = Texto + "Movimentações de saída não efetivadas " + Texto_Saída
      Texto = Texto + Chr(13)
    End If
    
    Texto = Texto + Chr(13) + "Para efetivar estas movimentações siga os passos abaixo: "
    Texto = Texto + Chr(13) + Chr(13) + "Operações de Entrada : Use a tela de entrada, encontre as movimentações e grave-as novamente."
    Texto = Texto + Chr(13) + "Operações de Saída : Use a tela de saídas ou a tela de venda rápida, encontre as movimentações, grave-as novamente e FAÇA O RECEBIMENTO."
    
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
