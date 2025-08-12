Attribute VB_Name = "modMenu"
Option Explicit

'13/01/2009 - mpdea
'Criação do menu
'
'27/01/2009 - mpdea
'Adaptado para o novo menu
'Key: Q7MENU

Private Const IMAGEBASE = 10000
Private m_obj_command_bar As CommandBars
Private m_obj_image_manager As ImageManager

Private Const PERMISSION_VISIVEL As String = "0"
Private Const PERMISSION_SEM_ACESSO As String = "-1"
Private Const PERMISSION_GRAVAR As String = "10"
Private Const PERMISSION_COMPLETO As String = "11"

Public Function MenuRibbonBar() As RibbonBar
  Set MenuRibbonBar = m_obj_command_bar.ActiveMenuBar
End Function

Public Sub CreateMenu(CommandBar As CommandBars, ImageManager As ImageManager)
  
  Set m_obj_command_bar = CommandBar
  Set m_obj_image_manager = ImageManager
  
  LoadIcons
  CreateRibbonBar
  CreateStatusBar
End Sub

Private Sub LoadIcons()
  m_obj_command_bar.Options.UseSharedImageList = False
  Set m_obj_command_bar.Icons = m_obj_image_manager.Icons
   
  'ToolTip
  Dim obj_tool_tip_context As ToolTipContext
  Set obj_tool_tip_context = m_obj_command_bar.ToolTipContext
  obj_tool_tip_context.Style = xtpToolTipOffice2007
  obj_tool_tip_context.ShowTitleAndDescription True, xtpToolTipIconNone
  obj_tool_tip_context.ShowImage True, IMAGEBASE
  obj_tool_tip_context.SetMargin 2, 2, 2, 2
  obj_tool_tip_context.MaxTipWidth = 180
  obj_tool_tip_context.ShowShadow = True
End Sub

Private Sub CreateRibbonBar()
  Dim obj_control As CommandBarControl
  Dim obj_control_pop As CommandBarPopup
  Dim obj_control_pop_n2 As CommandBarPopup
  Dim obj_tab As RibbonTab
  Dim obj_group As RibbonGroup
  Dim str_ret As String

  'RibbonBar
  Dim obj_ribbon_bar As RibbonBar
  Set obj_ribbon_bar = m_obj_command_bar.AddRibbonBar("The Ribbon")
  obj_ribbon_bar.EnableDocking xtpFlagStretched
  obj_ribbon_bar.Customizable = False
  obj_ribbon_bar.AllowQuickAccessCustomization = False
  
  '----------------------------------------------------------------------------
  'Controle principal
  Dim ControlFile As CommandBarPopup
  Set ControlFile = obj_ribbon_bar.AddSystemButton()
  ControlFile.id = ID_SYSTEM_CONTROL
  ControlFile.IconId = ID_SYSTEM_ICON
  ControlFile.CommandBar.SetIconSize 32, 32
  
  'Itens do controle principal
  ControlFile.CommandBar.Controls.Add xtpControlButton, ID_ITEM_ARQUIVO_ESTACOES_CONECTADAS, "&Estações Conectadas", False, False
  ControlFile.CommandBar.Controls.Add xtpControlButton, ID_ITEM_ARQUIVO_LOGON, "&Logon", False, False
  Set obj_control = ControlFile.CommandBar.Controls.Add(xtpControlButton, ID_ITEM_ARQUIVO_COMPACTAR_BASE, "&Compactar Base de Dados", False, False)
  obj_control.BeginGroup = True
  ControlFile.CommandBar.Controls.Add xtpControlButton, ID_ITEM_ARQUIVO_REPARAR_BASE, "&Reparar Base de Dados", False, False
  ControlFile.CommandBar.Controls.Add xtpControlButton, ID_ITEM_ARQUIVO_EXPORTAR_BASE, "&Exportar Base de Dados", False, False
  Set obj_control = ControlFile.CommandBar.Controls.Add(xtpControlButton, ID_ITEM_ARQUIVO_BACKUP, "&Backup", False, False)
  obj_control.BeginGroup = True
  Set obj_control = ControlFile.CommandBar.Controls.Add(xtpControlButton, ID_APP_EXIT, "&Sair", False, False)
  obj_control.BeginGroup = True
  
  '----------------------------------------------------------------------------
  'Exemplos
  '----------------------------------------------------------------------------
  'Exemplo de como criar um menu lateral ao principal
'  Set obj_control = m_obj_command_bar.CreateCommandBarControl("CXTPRibbonControlSystemPopupBarListCaption")
'  obj_control.Caption = "Cadastros"
'  obj_control.BeginGroup = True
'  ControlFile.CommandBar.Controls.AddControl obj_control
'
'  'Exemplo de como exibir um menu de janelas
'  Set obj_control = obj_group.Add(xtpControlPopup, ID_WINDOW_SWITCH, "Janelas", False, False)
'  obj_control.CommandBar.Controls.Add xtpControlButton, XtremeCommandBars.XTPCommandBarsSpecialCommands.XTP_ID_WINDOWLIST, "Item 1", False, False
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'Ajuda
  Set obj_control_pop = obj_ribbon_bar.Controls.Add(xtpControlButtonPopup, ID_POPUP_AJUDA_AJUDA, "     QUICK...ME AJUDE   ")
  obj_control_pop.Flags = xtpFlagRightAlign
  obj_control_pop.Caption = "     QUICK...ME AJUDE   "
  obj_control_pop.DescriptionText = "     QUICK...ME AJUDE   "
  obj_control_pop.ShortcutText = "     QUICK...ME AJUDE   "
  obj_control_pop.IconId = 1217
  obj_control_pop.SetIconSize 30, 30
  obj_control_pop.Style = xtpButtonIconAndCaption
  
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_AJUDA_PESQUISA, "Perguntas e respostas rápidas", False, False
    .Add xtpControlButton, ID_ITEM_AJUDA_CONTEUDO, "Conteúdo", False, False
    .Add xtpControlButton, ID_ITEM_AJUDA_SOBRE, "Sobre", False, False
    '.Add xtpControlButton, ID_ITEM_AJUDA_REGISTRO, "Registro", False, False
    '.Add xtpControlButton, ID_ITEM_AJUDA_INSTITUCIONAL, "Institucional", False, False
    .Add xtpControlButton, ID_ITEM_AJUDA_AGENDA, "Painel de Informações e Tips", False, False
  End With
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'TAB Início
  Set obj_tab = obj_ribbon_bar.InsertTab(0, "&Início")
  obj_tab.id = ID_TAB_INICIO
'''  'GROUP Área de Transferência
'''  Set obj_group = obj_tab.Groups.AddGroup("Área de Transferência", ID_GROUP_INICIO_AREA_TRANSF)
'''  obj_group.Add xtpControlButton, ID_ITEM_INICIO_COLAR, "Colar", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_INICIO_RECORTAR, "Recortar", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_INICIO_COPIAR, "Copiar", False, False
  'GROUP Cockpit
  Set obj_group = obj_tab.Groups.AddGroup("Gestão", ID_GROUP_INICIO_GESTAO)
  obj_group.Add xtpControlButton, ID_ITEM_INICIO_COCKPIT, "Cockpit Estratégico/Gerencial", False, False
  'GROUP Geral
  Set obj_group = obj_tab.Groups.AddGroup("Geral", ID_GROUP_INICIO_GERAL)
  obj_group.Add xtpControlButton, ID_ITEM_INICIO_LIVRO_PONTO, "Tarefas", False, False
  
  'GROUP StandBy
  Set obj_group = obj_tab.Groups.AddGroup("Status", ID_GROUP_INICIO_STANDBY)
  obj_group.Add xtpControlButton, ID_ITEM_INICIO_STANDBY, "Stand by", False, False

'''  Set obj_group = obj_tab.Groups.AddGroup("Help Quick", ID_GROUP_HELP_QUICK)
'''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Help On-line", False, False

  'GROUP Parâmetros
  Set obj_group = obj_tab.Groups.AddGroup("Parâmetros", ID_GROUP_INICIO_PARAMETROS)
  obj_group.Add xtpControlButton, ID_ITEM_INICIO_PARAM_EMPRESA, "Empresa/Filial", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_INICIO_PARAM_IMPOSTO_ESTADUAL, "Impostos Estaduais", False, False
  obj_group.Add xtpControlButton, ID_ITEM_INICIO_PARAM_CONFIG_IMPRESSORA, "Configuração Impressora", False, False
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_ITEM_INICIO_PARAM_TEMA, "Tema do Aplicativo", False, False)
  obj_control_pop.BeginGroup = True
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_INICIO_PARAM_TEMA_AQUA, "Tema Aqua", False, False
    .Add xtpControlButton, ID_ITEM_INICIO_PARAM_TEMA_PRETO, "Tema Preto", False, False
    .Add xtpControlButton, ID_ITEM_INICIO_PARAM_TEMA_AZUL, "Tema Azul", False, False
  End With
  'GROUP Suplementos
  Set obj_group = obj_tab.Groups.AddGroup("Suplementos", ID_GROUP_INICIO_SUPLEMENTO)
  'Suplementos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_SUPLEMENTO, "Suplementos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_INICIO_PARAM_CLASS_CLIENTE, "Classificação dos Clientes", False, False
    .Add xtpControlButton, ID_ITEM_INICIO_PARAM_FATURAMENTO_AUTO, "Faturamento Automático", False, False
    .Add xtpControlButton, ID_ITEM_INICIO_PARAM_DEVOL_MATERIAL, "Devolução de Materiais", False, False
  End With
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'TAB Cadastro
  Set obj_tab = obj_ribbon_bar.InsertTab(1, "&Cadastro")
  obj_tab.id = ID_TAB_CADASTRO
  'GROUP Geral
  Set obj_group = obj_tab.Groups.AddGroup("Geral", ID_GROUP_CADASTRO_GERAL)
  'Produto
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_CADASTRO_PRODUTO, "Produtos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_CADASTRO_PRODUTO, "Produtos", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_PRODUTO_CFOP, "ProdutosCFOPs", False, False
    
    Set obj_control = .Add(xtpControlButton, ID_ITEM_CADASTRO_CLASSE, "Classes", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_CADASTRO_SUBCLASSE, "Subclasses", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_CADASTRO_COR, "Cores", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_CADASTRO_TAMANHO, "Tamanhos", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_CADASTRO_ETIQUETA_PRODUTO, "Etiquetas - Criar lista de produtos e quantidade", False, False)
    Set obj_control = .Add(xtpControlButton, ID_ITEM_FORMATAR_ETIQUETA_PRODUTO, "Etiquetas - Formatar", False, False)
    obj_control.BeginGroup = True
  End With
  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_SERVICO, "Serviços", False, False
  'Pessoas
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_CADASTRO_PESSOA, "Pessoas", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_CADASTRO_CLIENTE_FORNEC, "Clientes/Fornecedores", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_CARACT_CLIENTE_FORNEC, "Características Clientes/Fornecedores", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_TRANSPORTADORA, "Transportadoras", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_USUARIO, "Usuários", False, False
  End With
  
  'Fidelidade
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_CADASTRO_PROGRAMA_FIDELIDADE, "Programa Fidelidade", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE, "Programa", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_OPERACOES_SAIDA, "Programa x Op.Saída", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CLIENTES_NAO_PART, "Programa x Clientes não participam", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CNPJ_GRUPOS, "Programa x CNPJs Participantes", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CONSULTA_GERENCIAL, "Consultas Gerenciais", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_RESGATE_PONTOS, "Resgate de Pontos", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CLIENTE_ENTREGA_RESGATE, "Cliente entrega Resgate", False, False
  End With
  
'''  'Pesquisa
'''  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_CADASTRO_PESQUISA, "Pesquisa", False, False)
'''  With obj_control_pop.CommandBar.Controls
'''    .Add xtpControlButton, ID_ITEM_CADASTRO_PESQUISA_1, "Pesquisa 1", False, False
'''    .Add xtpControlButton, ID_ITEM_CADASTRO_PESQUISA_2, "Pesquisa 2", False, False
'''    .Add xtpControlButton, ID_ITEM_CADASTRO_PESQUISA_3, "Pesquisa 3", False, False
'''  End With
  'Financeiro
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_CADASTRO_FINANCEIRO, "Financeiro", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_CADASTRO_BANCO, "Bancos", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_CONTA_CORRENTE, "Contas Correntes", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_CARTAO, "Cartões", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_CAIXA, "Caixas", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_MOEDA, "Moedas", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_COTACAO, "Cotações", False, False
  End With
  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_CENTRO_CUSTO, "Centros de Custos", False, False
  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_CODIGO_NBM, "Códigos NCM", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_CLASSIFICACAO_FISCAL, "Classificações Fiscais", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_GRUPO_FISCAL, "Grupos Fiscais", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_MENSAGEM_NOTA_FISCAL, "Mensagens para Nota Fiscal", False, False
  'GROUP Mala Direta
  Set obj_group = obj_tab.Groups.AddGroup("Mala Direta", ID_GROUP_CADASTRO_MALA_DIRETA)
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_CADASTRO_MALA_DIRETA, "Mala Direta", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_CADASTRO_MALA_DIRETA_PREPARAR, "Preparação", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_MALA_DIRETA_MANUTENCAO, "Manutenção", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_MALA_DIRETA_PREPARAR_REMETENTE, "Preparação Remetente", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_MALA_DIRETA_GRUPO, "Grupos", False, False
  End With
  'GROUP Operações
  Set obj_group = obj_tab.Groups.AddGroup("Operações", ID_GROUP_CADASTRO_OPERACAO)
  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_OPERACAO_ENTRADA, "Entradas", False, False
  obj_group.Add xtpControlButton, ID_ITEM_CADASTRO_OPERACAO_SAIDA, "Saídas", False, False
  'GROUP Suplementos
  Set obj_group = obj_tab.Groups.AddGroup("Suplementos", ID_GROUP_CADASTRO_SUPLEMENTO)
  'Suplementos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_SUPLEMENTO, "Suplementos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_CADASTRO_RADIO, "Rádios", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_TIPO_COMERCIAL, "Tipos Comerciais", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_AUT_PUBLICIDADE, "Autorizações de Publicidade", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_SUPERVISOR, "Supervisores", False, False
    .Add xtpControlButton, ID_ITEM_CADASTRO_RETENCAO, "Retenções", False, False
  End With
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'TAB Movimento
  Set obj_tab = obj_ribbon_bar.InsertTab(2, "&Movimento")
  obj_tab.id = ID_TAB_MOVIMENTO
  'GROUP Geral
  Set obj_group = obj_tab.Groups.AddGroup("Geral", ID_GROUP_MOVIMENTO_GERAL)
  obj_group.Add xtpControlButton, ID_ITEM_MOVIMENTO_VENDA_RAPIDA, "&Venda Rápida", False, False
  Set obj_control = obj_group.Add(xtpControlButton, ID_ITEM_MOVIMENTO_ENTRADAS, "&Entradas", False, False)
  obj_control.BeginGroup = True
  obj_group.Add xtpControlButton, ID_ITEM_MOVIMENTO_SAIDAS, "&Saídas", False, False
  obj_group.Add xtpControlButton, ID_ITEM_MOVIMENTO_DEVOLUCOES, "&Devoluções", False, False
  obj_group.Add xtpControlButton, ID_ITEM_REL_SAIDAS_ENTRADAS, "&Piloto - Saídas e Entradas", False, False
  obj_group.Add xtpControlButton, ID_ITEM_REL_GRAFICO4_VENDA_PRODUTOS, "&CockPit de Produtos", False, False
  Set obj_control = obj_group.Add(xtpControlButton, ID_ITEM_MOVIMENTO_ORDEM_SERVICO, "&Ordem de Serviços", False, False)
  obj_control.BeginGroup = True
  Set obj_control = obj_group.Add(xtpControlButton, ID_ITEM_MOVIMENTO_PEDIDOS_WEB, "&Pedidos da Loja Virtual", False, False)
  obj_control.BeginGroup = True
  Set obj_control = obj_group.Add(xtpControlButton, ID_ITEM_MOVIMENTO_NOTA_FISCAL_ELETRONICA, "&Nota Fiscal Eletrônica", False, False)
  obj_control.BeginGroup = True
  'colocar botão para importar vendas do gesto quando houver paf
'''  Set obj_control = obj_group.Add(xtpControlButton, ID_IMPORTA_GESTO, "&Importar vendas do gesto", False, False)
  Set obj_control = obj_group.Add(xtpControlButton, ID_IMPORTA_GESTO, "&Site Benefix", False, False)
  obj_control.BeginGroup = True
  'Transferência
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_MOVIMENTO_TRANSFERENCIA, "Transferência", False, False)
  obj_control_pop.BeginGroup = True
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_TRANSF_FILIAL, "&Transferência entre Filiais", False, False
  End With
  'Empréstimos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_MOVIMENTO_EMPRESTIMO, "Empréstimos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_EMPREST_ENTRADA, "Acerto de Empréstimos de &Entrada", False, False
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_EMPREST_SAIDA, "Acerto de Empréstimos de &Saídas", False, False
  End With
  'Manutenção
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_MOVIMENTO_MANUT, "Manutenção", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_APAGAR_EMP_ENTRADA, "Apagar Acerto de Empréstimos de Entrada", False, False
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_APAGAR_EMP_SAIDA, "Apagar Acerto de Empréstimos de Saída", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_MOVIMENTO_APAGAR_ENTRADA, "Apagar &Entradas", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_APAGAR_SAIDA, "Apagar &Saídas", False, False
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_APAGAR_MOVIMENTACAO, "&Apagar Movimentação ou Zerar Estoque de Produtos", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_MOVIMENTO_MANUT_CONSIGNACAO, "Manutenção de &Consignação", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_MANUT_ORCAMENTO, "Manutenção de &Orçamento", False, False
  End With
  'GROUP Suplementos
  Set obj_group = obj_tab.Groups.AddGroup("Relatórios/Suplementos", ID_GROUP_MOVIMENTO_SUPLEMENTO)
  'Suplementos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_SUPLEMENTO, "Suplementos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_MANUT_RESERVA, "&Manutenção de Reservas", False, False
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_MANUT_CONSIG_ENTRADA, "&Manutenção de Consignação de Entrada", False, False
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_FATUR_AUTO, "&Faturamento Automático", False, False
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_PREST_FORNEC, "&Prestação de Contas com Fornecedores", False, False
    .Add xtpControlButton, ID_ITEM_MOVIMENTO_IMPORTACAO, "&Importação", False, False
  End With
  
  'Movimento
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_MOVIMENTO, "Mov.Saída, Entrada, NFe e NFCe", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_MOV_SAIDA, "Saídas", False, False
    .Add xtpControlButton, ID_ITEM_REL_MOV_ENTRADA, "Entradas", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_MOV_ACERTA_EMPREST_ENTRADA, "Acerta Empréstimos de Entrada", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_REL_MOV_ACERTA_EMPREST_SAIDA, "Acerta Empréstimos de Saídas", False, False
  End With
  'Compras/Vendas
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_COMPRA_VENDA, "Compras/Vendas", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA, "Vendas", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_COMISSAO, "Comissões", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_COMISSAO_VENDA_VENDEDOR, "Comissões de Vendas por Vendedor", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_COMPRAS, "Compras", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_CLIENTE, "Vendas por Cliente", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_TAMANHO, "Vendas por Tamanho", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_EDITORA, "Vendas por Editora", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_POR_VENDEDOR, "Vendas por Vendedor", False, False
  End With
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'TAB Financeiro
  Set obj_tab = obj_ribbon_bar.InsertTab(3, "Financeiro")
  obj_tab.id = ID_TAB_FINANCEIRO
  'GROUP Conta Corrente
  Set obj_group = obj_tab.Groups.AddGroup("Conta Corrente", ID_GROUP_FINANCEIRO_CC)
  obj_group.Add xtpControlButton, ID_ITEM_FINANCEIRO_LANC_BANC, "Lançamentos", False, False
  obj_group.Add xtpControlButton, ID_ITEM_FINANCEIRO_RECAL_SALDO, "Recalcular Saldos", False, False
  obj_group.Add xtpControlButton, ID_ITEM_FINANCEIRO_APAGA_LANC_BANC, "&Apagar Lançamentos", False, False
  'GROUP Caixa
  Set obj_group = obj_tab.Groups.AddGroup("Caixa", ID_GROUP_FINANCEIRO_CAIXA)
  obj_group.Add xtpControlButton, ID_ITEM_FINANCEIRO_MOV_MANUAL_CAIXA, "&Movimentação de Caixa", False, False
  obj_group.Add xtpControlButton, ID_ITEM_FINANCEIRO_APAGA_LANC_CAIXA, "&Apagar Lançamentos", False, False
  'GROUP Contas a Pagar
  Set obj_group = obj_tab.Groups.AddGroup("Contas a Pagar", ID_GROUP_FINANCEIRO_CONTAS_PAGAR)
  'Movimento
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_FINANCEIRO_CP_MOVIMENTO, "Movimento", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CP_LANCAR, "Lançamentos/Manutenção", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CP_GERAR, "Lançar Parcelas de Contas a Pagar", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CP_MANUT, "Realizar baixa/Pagar", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CP_APAGAR_PAGA, "&Apagar Contas Pagas", False, False
  End With
  'GROUP Contas a Receber
  Set obj_group = obj_tab.Groups.AddGroup("Contas a Receber", ID_GROUP_FINANCEIRO_CONTAS_RECEBER)
  'Movimento
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_FINANCEIRO_CR_MOVIMENTO, "Movimento", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_LANCAR, "Lançamentos/Manutenção", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_MANUT, "Realizar baixa/Receber", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_APAGAR_RECEBIDA, "Apagar Contas Recebidas", False, False
  End With
  'Cheque Pré-Datado
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_FINANCEIRO_CR_CHEQUE_PRE, "Cheque Pré-Datado", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_LANCAR_CHEQUE_PRE, "&Lançamentos/Manutenção de Cheques Pré-Datados", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_MANUT_CHEQUE_PRE, "&Realizar baixa de Cheques Pré-Datados", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_APAGAR_CHEQUE_PRE, "&Apagar Cheques Pré-Datados", False, False
  End With
  'Cartão
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_FINANCEIRO_CR_CARTAO, "Cartão", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_LANCAR_CARTAO, "&Lançamentos/Manutenção de Cartões", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_MANUT_CARTAO, "&Realizar baixa de Cartões", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_APAGAR_CARTAO, "&Apagar Cartões", False, False
  End With
  'Conta Cliente
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_FINANCEIRO_CR_CONTA_CLIENTE, "Conta de Cliente", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_LANCAR_CONTA_CLIENTE, "&Lançamentos/Manutenção de Contas de Cliente", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_MANUT_CONTA_CLIENTE, "&Realizar Recebimento de Contas de Cliente", False, False
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_APAGAR_CONTA_CLIENTE, "&Apagar Contas de Cliente", False, False
  End With
  'GROUP Suplementos
  Set obj_group = obj_tab.Groups.AddGroup("Relatórios/Suplementos", ID_GROUP_FINANCEIRO_SUPLEMENTO)
  'Suplementos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_SUPLEMENTO, "Suplementos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_FINANCEIRO_CR_AUT_PUBLICIDADE, "&Autorizações de Publicidade", False, False
  End With
  
  'Financeiro
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_FINANCEIRO, "Relatórios", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_FINANC_CAIXA, "Caixas", False, False
    .Add xtpControlButton, ID_ITEM_REL_FINANC_CARTAO, "Cartões", False, False
    .Add xtpControlButton, ID_ITEM_REL_FINANC_LANC_BANCARIO, "Lançamentos Bancários", False, False
    .Add xtpControlButton, ID_ITEM_REL_FINANC_SALDO_CC, "Saldos das Contas", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_FINANC_DIARIO_1, "Financeiro Diário 1", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_REL_FINANC_DIARIO_2, "Financeiro Diário 2", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_FINANC_LUCRATIVIDADE, "Lucratividade", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_REL_FINANC_GERAL, "Financeiro Geral", False, False
    .Add xtpControlButton, ID_ITEM_REL_FINANC_RECEB_FORMA_PGTO, "Recebimentos por Formas de Pagamento", False, False
    'Contas a Pagar
    Set obj_control_pop_n2 = .Add(xtpControlPopup, ID_POPUP_RELATORIO_FINANC_CONTA_PAGAR, "Contas a Pagar", False, False)
    obj_control_pop_n2.BeginGroup = True
    With obj_control_pop_n2.CommandBar.Controls
      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_DATA_VCTO, "A Pagar por Data de Vencimento", False, False
      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_FORNECEDOR, "A Pagar por Fornecedor", False, False
      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_GERAL_FILIAL, "A Pagar - Todas as Filiais", False, False
      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_CENTRO_CUSTO, "A Pagar por Centro de Custo", False, False
      Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_CP_PAGAS_FORNECEDOR, "Pagas por Fornecedor", False, False)
      obj_control.BeginGroup = True
      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAS_DATA_PGTO, "Pagas por Data de Pagamento", False, False
      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAS_CENTRO_CUSTO, "Pagas por Centro de Custo", False, False
      Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_CP_CONTROLE_CENTRO_CUSTO, "Controle de Centro de Custo", False, False)
      obj_control.BeginGroup = True
      .Add xtpControlButton, ID_ITEM_REL_CP_CENTRO_CUSTO_COMPETENCIA, "Centros de Custo pela Competência", False, False
    End With
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'TAB Preços
  Set obj_tab = obj_ribbon_bar.InsertTab(4, "&Preços")
  obj_tab.id = ID_TAB_PRECO
  'GROUP Geral
  Set obj_group = obj_tab.Groups.AddGroup("Geral", ID_GROUP_PRECO_GERAL)
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_CRIAR_TAB, "Cria/Recria Tabela", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_APAGAR_TAB, "Apagar Tabela de Preços", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_LANCAR, "Lançamento de Preços", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_ALTERAR, "Alteração de Preços", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_ALTERAR_CALC, "Alteração de Preços Calculado", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_CONFIG_TAB, "Configuração da Tabela", False, False
  
  'Preços
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_PRECO, "Relatórios", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_PRECO_LISTA, "Lista de Preços", False, False
    .Add xtpControlButton, ID_ITEM_REL_PRECO_LOCAL_PRODUTO, "Localização dos Produtos", False, False
  End With
  
  'GROUP Copiar Tabela
  Set obj_group = obj_tab.Groups.AddGroup("Copiar Tabela", ID_GROUP_PRECO_COPIAR_TABELA)
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_COPIAR_TAB_IND, "Aplicar Índice", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_COPIAR_TAB_VALOR, "Aplicar Valor", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_COPIAR_TAB_CUSTO_MEDIO, "Custo Médio", False, False
  'GROUP Calcular preços
  Set obj_group = obj_tab.Groups.AddGroup("Calcular", ID_GROUP_PRECO_CALCULAR)
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_CALC_PRECO, "Preço de Venda", False, False
  obj_group.Add xtpControlButton, ID_ITEM_PRECO_CALC_PRECO_SIMPLES, "Preço de Venda Simplificado", False, False
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'TAB Estoque
  Set obj_tab = obj_ribbon_bar.InsertTab(5, "&Estoque")
  obj_tab.id = ID_TAB_ESTOQUE
  'GROUP Normal
  Set obj_group = obj_tab.Groups.AddGroup("Normal", ID_GROUP_ESTOQUE_NORMAL)
  obj_group.Add xtpControlButton, ID_ITEM_ESTOQUE_INFO_CONTAR, "Ajustar Estoque", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_ESTOQUE_ACERTAR_CONTAR, "Acertar Contagem", False, False
  'GROUP Grade
  Set obj_group = obj_tab.Groups.AddGroup("Grade", ID_GROUP_ESTOQUE_GRADE)
  obj_group.Add xtpControlButton, ID_ITEM_ESTOQUE_INFO_CONTAR_GRADE, "Ajustar Estoque", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_ESTOQUE_ACERTAR_CONTAR_GRADE, "Acertar Contagem", False, False
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  
  '----------------------------------------------------------------------------

  '----------------------------------------------------------------------------
  'TAB Gerador
  Set obj_tab = obj_ribbon_bar.InsertTab(6, "&Gerador")
  obj_tab.id = ID_TAB_GERADOR
  'GROUP Geral
  Set obj_group = obj_tab.Groups.AddGroup("Geral", ID_GROUP_GERADOR_GERAL)
  obj_group.Add xtpControlButton, ID_ITEM_GERADOR_RELATORIO, "&Relatórios", False, False
  obj_group.Add xtpControlButton, ID_ITEM_GERADOR_LAYOUT_NOTA, "Layout de &Nota Fiscal", False, False
  obj_group.Add xtpControlButton, ID_ITEM_GERADOR_LAYOUT_TICKET, "Layout de &Ticket", False, False
  obj_group.Add xtpControlButton, ID_ITEM_GERADOR_LAYOUT_BOLETO, "Layout de &Boleto Bancário", False, False
  obj_group.Add xtpControlButton, ID_ITEM_GERADOR_LAYOUT_CARNET, "Layout de &Carnê", False, False
'''  obj_group.Add xtpControlButton, ID_ITEM_GERADOR_ARQ_REC_ESTADUAL, "Arquivo Receita Estadual", False, False
  'GROUP Suplementos
  Set obj_group = obj_tab.Groups.AddGroup("Suplementos", ID_GROUP_GERADOR_SUPLEMENTO)
  'Suplementos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_SUPLEMENTO, "Suplementos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_GERADOR_EXP_BR_INFO, "Exportar Dados Brasil Informática", False, False
    .Add xtpControlButton, ID_ITEM_GERADOR_EXP_SADIG_WEB, "Exportar Dados Sadig Web", False, False
    .Add xtpControlButton, ID_ITEM_GERADOR_EXP_PEARSON, "Exportar Dados Pearson", False, False
  End With
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  'TAB Relatórios
  Set obj_tab = obj_ribbon_bar.InsertTab(7, "&Relatórios")
  obj_tab.id = ID_TAB_RELATORIO
  'GROUP Geral
  Set obj_group = obj_tab.Groups.AddGroup("Geral", ID_GROUP_RELATORIO_GERAL)
  'Serviços
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_SERVICO, "Serviços", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_SERVICO_EXECUTADO, "Serviços Executados", False, False
    .Add xtpControlButton, ID_ITEM_REL_SERVICO_COMISSAO, "Comissões", False, False
  End With
  'Produtos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_PRODUTO, "Produtos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_PRODUTO_GERAL, "Geral", False, False
    .Add xtpControlButton, ID_ITEM_REL_PRODUTO_GRADE, "Grade", False, False
  End With
  'Estoque
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_ESTOQUE, "Estoque", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_ESTOQUE_GERAL, "Estoque Geral", False, False
    .Add xtpControlButton, ID_ITEM_REL_ESTOQUE_GRADE, "Grade", False, False
    .Add xtpControlButton, ID_ITEM_REL_ESTOQUE_ANALITICO, "Analítico", False, False
    .Add xtpControlButton, ID_ITEM_REL_ESTOQUE_POR_FILIAL, "Por Filial", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_ESTOQUE_PRODUTO_COMPRAR, "Produtos a Comprar", False, False)
    obj_control.BeginGroup = True
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_ESTOQUE_ACOMPANHAMENTO_PRODUTO, "Acompanhamento de Produto", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_REL_ESTOQUE_ACOMPANHAMENTO_ESTOQUE, "Movimentação de Estoque Simplificado", False, False
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_ESTOQUE_REGISTRO_INVENTARIO, "Registro de Inventário", False, False)
    obj_control.BeginGroup = True
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_ESTOQUE_CONTAGEM, "Contagem de Estoque", False, False)
    obj_control.BeginGroup = True
    .Add xtpControlButton, ID_ITEM_REL_ESTOQUE_CONTAGEM_GRADE, "Contagem de Estoque - Grade", False, False
  End With
'''  'Compras/Vendas
'''  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_COMPRA_VENDA, "Compras/Vendas", False, False)
'''  With obj_control_pop.CommandBar.Controls
'''    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA, "Vendas", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_CV_COMISSAO, "Comissões", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_CV_COMISSAO_VENDA_VENDEDOR, "Comissões de Vendas por Vendedor", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_CV_COMPRAS, "Compras", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_CLIENTE, "Vendas por Cliente", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_TAMANHO, "Vendas por Tamanho", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_EDITORA, "Vendas por Editora", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_POR_VENDEDOR, "Vendas por Vendedor", False, False
'''  End With
'''  'Movimento
'''  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_MOVIMENTO, "Movimento/NFe/NFCe", False, False)
'''  With obj_control_pop.CommandBar.Controls
'''    .Add xtpControlButton, ID_ITEM_REL_MOV_ENTRADA, "Entradas", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_MOV_SAIDA, "Saídas", False, False
'''    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_MOV_ACERTA_EMPREST_ENTRADA, "Acerta Empréstimos de Entrada", False, False)
'''    obj_control.BeginGroup = True
'''    .Add xtpControlButton, ID_ITEM_REL_MOV_ACERTA_EMPREST_SAIDA, "Acerta Empréstimos de Saídas", False, False
'''  End With
  'Cliente/Fornecedores
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_CLIETE_FORNECEDOR, "Cliente/Fornecedores", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_CLIENTE_FORNECEDOR, "Cliente/Fornecedores", False, False
    .Add xtpControlButton, ID_ITEM_REL_CONTATO_EFETUADO, "Contatos Efetuados", False, False
    .Add xtpControlButton, ID_ITEM_REL_CONTATO_DATA_ANIVERSARIO, "Data Aniversário Contatos e Clientes", False, False
  End With
  'Usuários
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_USUARIO, "Usuários", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_USUARIO_FUNCIONARIO, "Lista de Usuários/Funcionários", False, False
    .Add xtpControlButton, ID_ITEM_REL_LIVRO_PONTO, "Tarefas", False, False
  End With
  'Cadastro
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_CADASTRO, "Cadastro", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_CLASSE, "Classe", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_SUBCLASSE, "Subclasse", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_COR, "Cor", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_TAMANHO, "Tamanho", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_ETIQUETA_PRODUTO, "Etiquetas de Produtos", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_BANCO, "Bancos", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_CARTAO, "Cartões", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_MOEDA, "Moedas", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_COTACAO, "Cotações", False, False
    .Add xtpControlButton, ID_ITEM_REL_CADASTRO_CENTRO_CUSTO, "Centros de Custos", False, False
  End With
'''  'Financeiro
'''  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_FINANCEIRO, "Financeiro", False, False)
'''  With obj_control_pop.CommandBar.Controls
'''    .Add xtpControlButton, ID_ITEM_REL_FINANC_CAIXA, "Caixas", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_FINANC_CARTAO, "Cartões", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_FINANC_LANC_BANCARIO, "Lançamentos Bancários", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_FINANC_SALDO_CC, "Saldos das Contas", False, False
'''    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_FINANC_DIARIO_1, "Financeiro Diário 1", False, False)
'''    obj_control.BeginGroup = True
'''    .Add xtpControlButton, ID_ITEM_REL_FINANC_DIARIO_2, "Financeiro Diário 2", False, False
'''    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_FINANC_LUCRATIVIDADE, "Lucratividade", False, False)
'''    obj_control.BeginGroup = True
'''    .Add xtpControlButton, ID_ITEM_REL_FINANC_GERAL, "Financeiro Geral", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_FINANC_RECEB_FORMA_PGTO, "Recebimentos por Formas de Pagamento", False, False
'''    'Contas a Pagar
'''    Set obj_control_pop_n2 = .Add(xtpControlPopup, ID_POPUP_RELATORIO_FINANC_CONTA_PAGAR, "Contas a Pagar", False, False)
'''    obj_control_pop_n2.BeginGroup = True
'''    With obj_control_pop_n2.CommandBar.Controls
'''      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_DATA_VCTO, "A Pagar por Data de Vencimento", False, False
'''      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_FORNECEDOR, "A Pagar por Fornecedor", False, False
'''      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_GERAL_FILIAL, "A Pagar - Todas as Filiais", False, False
'''      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAR_CENTRO_CUSTO, "A Pagar por Centro de Custo", False, False
'''      Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_CP_PAGAS_FORNECEDOR, "Pagas por Fornecedor", False, False)
'''      obj_control.BeginGroup = True
'''      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAS_DATA_PGTO, "Pagas por Data de Pagamento", False, False
'''      .Add xtpControlButton, ID_ITEM_REL_CP_PAGAS_CENTRO_CUSTO, "Pagas por Centro de Custo", False, False
'''      Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_CP_CONTROLE_CENTRO_CUSTO, "Controle de Centro de Custo", False, False)
'''      obj_control.BeginGroup = True
'''      .Add xtpControlButton, ID_ITEM_REL_CP_CENTRO_CUSTO_COMPETENCIA, "Centros de Custo pela Competência", False, False
'''    End With
    'Contas a Receber
    Set obj_control_pop_n2 = .Add(xtpControlPopup, ID_POPUP_RELATORIO_FINANC_CONTA_RECEBER, "Contas a Receber", False, False)
    With obj_control_pop_n2.CommandBar.Controls
      .Add xtpControlButton, ID_ITEM_REL_CR_LANCAMENTOS_DATA_EMISSAO, "Lançamentos de Contas a Receber por data de emissão", False, False
      .Add xtpControlButton, ID_ITEM_REL_CR_RECEBER_DATA_VCTO, "A Receber por Data de Vencimento", False, False
      .Add xtpControlButton, ID_ITEM_REL_CR_RECEBER_CLIENTE, "A Receber por Cliente", False, False
      .Add xtpControlButton, ID_ITEM_REL_CR_RECEBER_VENDEDOR, "A Receber por Vendedor", False, False
      Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_CR_RECEBIDA_DATA_RECEBIMENTO, "Recebidas por Data de Recebimento", False, False)
      obj_control.BeginGroup = True
      .Add xtpControlButton, ID_ITEM_REL_CR_RECEBIDA_VENDEDOR, "Recebidas por Vendedor", False, False
      .Add xtpControlButton, ID_ITEM_REL_CR_RECEBIDA_CLIENTE, "Recebidas por Cliente", False, False
      .Add xtpControlButton, ID_ITEM_REL_CR_CHEQUE_PRE, "Cheques Pré-Datados", False, False
      .Add xtpControlButton, ID_ITEM_REL_CR_CARTAO, "Cartões", False, False
      .Add xtpControlButton, ID_ITEM_REL_CR_CONTA_CLIENTE, "Contas de Cliente", False, False
      Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_CR_EMISSAO_BOLETO, "Emissão de Boletos", False, False)
      obj_control.BeginGroup = True
      .Add xtpControlButton, ID_ITEM_REL_CR_EMISSAO_CARNET, "Emissão de Carnês", False, False
    End With
    Set obj_control = .Add(xtpControlButton, ID_ITEM_REL_FINANC_FLUXO_CAIXA, "Fluxo de Caixa", False, False)
    obj_control.BeginGroup = True
  End With
  obj_group.Add xtpControlButton, ID_ITEM_REL_MALA_DIRETA, "Mala Direta", False, False
  obj_group.Add xtpControlButton, ID_ITEM_REL_NSU_CORRELACAO, "NSU - Correlação", False, False
'''  'Preços
'''  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_PRECO, "Preços", False, False)
'''  With obj_control_pop.CommandBar.Controls
'''    .Add xtpControlButton, ID_ITEM_REL_PRECO_LISTA, "Lista de Preços", False, False
'''    .Add xtpControlButton, ID_ITEM_REL_PRECO_LOCAL_PRODUTO, "Localização dos Produtos", False, False
'''  End With
  'Gráficos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_RELATORIO_GRAFICO, "ESTRATÉGICO", False, False)
  If gESTRATEGICO_Relatorios = 0 Then
    With obj_control_pop.CommandBar.Controls
      .Add xtpControlButton, ID_ITEM_REL_ESTRATEGICO_AVISO_AQUISICAO, "Como adquirir Acesso à este módulo?!", False, False
    End With
  Else
    With obj_control_pop.CommandBar.Controls
      .Add xtpControlButton, ID_ITEM_REL_GRAFICO_COMPARATIVO_CV, "Comparativo de Compras e Vendas", False, False
      .Add xtpControlButton, ID_ITEM_REL_GRAFICO_VENDA_CLASSE_PERIODO, "Vendas por Classe no Período", False, False
      .Add xtpControlButton, ID_ITEM_REL_GRAFICO_VENDA_PRODUTO_MENSAL, "Vendas de um Produto Mês a Mês", False, False
'''      .Add xtpControlButton, ID_ITEM_REL_GRAFICO4_VENDA_PRODUTOS, "* Produtos mais vendidos", False, False
      .Add xtpControlButton, ID_ITEM_REL_GRAFICO5_COMPRA_FORNECEDORES, "* Maiores Fornecedores", False, False
      .Add xtpControlButton, ID_ITEM_REL_GRAFICO6_VENDA_CLIENTES, "* Maiores Clientes", False, False
      .Add xtpControlButton, ID_ITEM_REL_EXPORTA_CLIENTES_PRODUTO, "Exporta Clientes Por Produto", False, False
    End With
  End If
  
  'GROUP Suplementos
  Set obj_group = obj_tab.Groups.AddGroup("Suplementos", ID_GROUP_RELATORIO_SUPLEMENTO)
  'Suplementos
  Set obj_control_pop = obj_group.Add(xtpControlPopup, ID_POPUP_SUPLEMENTO, "Suplementos", False, False)
  With obj_control_pop.CommandBar.Controls
    .Add xtpControlButton, ID_ITEM_REL_ESTOQUE_FILIAL_PRECO, "Estoque das Filiais e Preços", False, False
    .Add xtpControlButton, ID_ITEM_REL_FOLHA_PGTO, "Folha de Pagamento", False, False
    .Add xtpControlButton, ID_ITEM_REL_AUTORIZACAO, "Autorizações", False, False
    .Add xtpControlButton, ID_ITEM_REL_MALA_DIRETA_AUTORIZACAO, "Mala Direta para Autorizações", False, False
    .Add xtpControlButton, ID_ITEM_REL_MALA_DIRETA_GERAR_ARQUIVO, "Gerar Arquivo para Mala Direta", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_FORNECEDOR, "Vendas por Fornecedor", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_PRODUTO_CONSIGNADO, "Vendas Produtos Consignados", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_VENDEDOR_COMISSAO, "Vendas por Vendedor e Comissões", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_COMISSAO_RETENCAO, "Comissões com Retenções", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_VENDA_2, "Vendas II", False, False
    .Add xtpControlButton, ID_ITEM_REL_MOV_ENTRADA_CONSIGNADA, "Entradas Consignadas", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_PRESTACAO_CONTA, "Prestação de Contas", False, False
    .Add xtpControlButton, ID_ITEM_REL_CV_PRODUTO_COMPRAR_FATOR, "Produtos a Comprar por Fator", False, False
  End With
  
''  Set obj_group = obj_tab.Groups.AddGroup("AJUDA", ID_GROUP_HELP_QUICK)
''  obj_group.Add xtpControlButton, ID_ACESSO_HELP_QUICK, "Estou com dúvidas...", False, False
  
  '----------------------------------------------------------------------------
  
  'Opções
  m_obj_command_bar.Options.KeyboardCuesShow = xtpKeyboardCuesShowAlways
  m_obj_command_bar.Options.ShowExpandButtonAlways = False
  m_obj_command_bar.EnableCustomization False
  
  '05/05/2009 - mpdea
  'Tema do menu
  str_ret = GetSetting("QuickStore", "Menu", "Tema", ID_ITEM_INICIO_PARAM_TEMA_AZUL)
  If Not IsDataType(dtLong, str_ret) Then
    str_ret = ID_ITEM_INICIO_PARAM_TEMA_AZUL
  End If
  SetMenuTheme CLng(str_ret)
  
  'Ativa o tema com Ribbon
  obj_ribbon_bar.EnableFrameTheme
  
  'Ponto de restauração para os controles
  MenuRibbonBar.Controls.CreateOriginalControls
  
'''  'Atalhos
'''  m_obj_command_bar.KeyBindings.Add FCONTROL, Asc("X"), ID_ITEM_INICIO_RECORTAR
'''  m_obj_command_bar.KeyBindings.Add FCONTROL, Asc("C"), ID_ITEM_INICIO_COPIAR
'''  m_obj_command_bar.KeyBindings.Add FCONTROL, Asc("V"), ID_ITEM_INICIO_COLAR
  
End Sub

'05/05/2009 - mpdea
'Seta o tema do menu
Public Sub SetMenuTheme(ByVal id As Long)
  Dim str_style As String
  
  Select Case id
    Case ID_ITEM_INICIO_PARAM_TEMA_AZUL
      str_style = "" 'Vazio (padrão)
      
    Case ID_ITEM_INICIO_PARAM_TEMA_AQUA
      str_style = "Office2007Aqua.dll"
      
    Case ID_ITEM_INICIO_PARAM_TEMA_PRETO
      str_style = "Office2007Black.dll"
      
  End Select
  
  'Completa Path
  If str_style <> "" Then str_style = App.Path & "\Styles\" & str_style
  
  'Estilo (Office2007Aqua.dll, Office2007Black.dll ou vazio)
  CommandBarsGlobalSettings.Office2007Images = str_style
  m_obj_command_bar.PaintManager.RefreshMetrics
  m_obj_command_bar.RecalcLayout
  
  'Salva o tema
  SaveSetting "QuickStore", "Menu", "Tema", id
  
End Sub

Private Sub CreateStatusBar()
  Dim StatusBar As XtremeCommandBars.StatusBar
  Set StatusBar = m_obj_command_bar.StatusBar
  StatusBar.Visible = True
  StatusBar.IdleText = "Pronto"
  
  Dim obj_status_bar_pane As StatusBarPane
  
  Set obj_status_bar_pane = StatusBar.AddPane(0) 'Texto
  obj_status_bar_pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
  obj_status_bar_pane.Text = "Pronto"
  obj_status_bar_pane.Width = 0
  
  Set obj_status_bar_pane = StatusBar.AddPane(ID_STATUSBAR_FILIAL)  'Filial
  obj_status_bar_pane.Text = "Filial: " & CStr(gnCodFilial)
  obj_status_bar_pane.Width = 0
  
  Set obj_status_bar_pane = StatusBar.AddPane(ID_STATUSBAR_USUARIO)  'Usuário
  obj_status_bar_pane.Text = "Usuário: " & CStr(gnUserCode) & "-" & gsUserName
  obj_status_bar_pane.Width = 0
  
  Set obj_status_bar_pane = StatusBar.AddPane(ID_STATUSBAR_VERSAO)  'Versão
  obj_status_bar_pane.Text = gsAppVersion
  obj_status_bar_pane.Width = 0
End Sub

'28/01/2009 - mpdea
'Retorna o Id utilizado no controle ActiveBar para o seu correspondente no controle CommandBar
'Necessário devido ao controle de acessos para o usuário do sistema
'Parâmetro: lngCommandBarControlId - Id do controle CommandBar
'Retorno: Id utilizado pelo controle ActiveBar
'         0 se não houver correspondência no controle de acesso
Private Function MenuActiveBarId(ByVal lngCommandBarControlId) As Long
  Dim lngReturnId As Long
  
  '14/09/2009 - mpdea
  'Padrão
  lngReturnId = lngCommandBarControlId
  
  Select Case lngCommandBarControlId
    '----------------------------------------------------------------------------
    'Menu Arquivo
    '----------------------------------------------------------------------------
    Case ID_ITEM_ARQUIVO_ESTACOES_CONECTADAS
      lngReturnId = 0
    
    Case ID_ITEM_ARQUIVO_LOGON
      lngReturnId = 0
    
    Case ID_ITEM_ARQUIVO_COMPACTAR_BASE
      lngReturnId = 0
    
    Case ID_ITEM_ARQUIVO_REPARAR_BASE
      lngReturnId = 0
    
    Case ID_ITEM_ARQUIVO_EXPORTAR_BASE
      lngReturnId = 0
    
    Case ID_ITEM_ARQUIVO_BACKUP
      lngReturnId = 0
    
    '----------------------------------------------------------------------------
    'Menu Ajuda
    '----------------------------------------------------------------------------
    Case ID_ITEM_AJUDA_CONTEUDO
      lngReturnId = 0
    
    Case ID_ITEM_AJUDA_PESQUISA
      lngReturnId = 0
    
    Case ID_ITEM_AJUDA_SOBRE
      lngReturnId = 0
    
    Case ID_ITEM_AJUDA_REGISTRO
      lngReturnId = 0
    
    Case ID_ITEM_AJUDA_INSTITUCIONAL
      lngReturnId = 0
    
    Case ID_ITEM_AJUDA_AGENDA
      lngReturnId = 0
    
    '----------------------------------------------------------------------------
    'TAB Início
    '----------------------------------------------------------------------------
'''    Case ID_ITEM_INICIO_COLAR
'''      lngReturnId = 0
'''
'''    Case ID_ITEM_INICIO_RECORTAR
'''      lngReturnId = 0
'''
'''    Case ID_ITEM_INICIO_COPIAR
'''      lngReturnId = 0
      
    Case ID_ITEM_INICIO_LIVRO_PONTO
      lngReturnId = 10030
      
    Case ID_ITEM_INICIO_STANDBY
      lngReturnId = 10029

'''    Case ID_ACESSO_HELP_QUICK
'''        lngReturnId = 10028
    
    Case ID_ITEM_INICIO_PARAM_EMPRESA
      lngReturnId = 30010
      
    Case ID_ITEM_INICIO_PARAM_IMPOSTO_ESTADUAL
      lngReturnId = 30020
    
    Case ID_ITEM_INICIO_PARAM_CONFIG_IMPRESSORA
      lngReturnId = 30030
    
    Case ID_ITEM_INICIO_PARAM_CLASS_CLIENTE
      lngReturnId = 320058
    
    Case ID_ITEM_INICIO_PARAM_FATURAMENTO_AUTO
      lngReturnId = 0
    
    Case ID_ITEM_INICIO_PARAM_DEVOL_MATERIAL
      lngReturnId = 0
    
    '----------------------------------------------------------------------------
    'TAB Cadastros
    '----------------------------------------------------------------------------
    Case ID_ITEM_CADASTRO_SERVICO
      lngReturnId = 40010
    
    Case ID_ITEM_CADASTRO_PRODUTO
      lngReturnId = 40020
      
    Case ID_ITEM_CADASTRO_PRODUTO_CFOP
      lngReturnId = 40021
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE
      lngReturnId = 40022
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_OPERACOES_SAIDA
      lngReturnId = 40023
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CONSULTA_GERENCIAL
      lngReturnId = 40024
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_RESGATE_PONTOS
      lngReturnId = 40025
    
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CLIENTE_ENTREGA_RESGATE
      lngReturnId = 40026
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CLIENTES_NAO_PART
      lngReturnId = 40027
      
    Case ID_ITEM_CADASTRO_PROGRAMA_FIDELIDADE_CNPJ_GRUPOS
      lngReturnId = 40028
    
    Case ID_ITEM_CADASTRO_CLIENTE_FORNEC
      lngReturnId = 40030
    
    Case ID_ITEM_CADASTRO_CARACT_CLIENTE_FORNEC
      lngReturnId = 40030
    
    Case ID_ITEM_CADASTRO_TRANSPORTADORA
      lngReturnId = 40040
    
    Case ID_ITEM_CADASTRO_USUARIO
      lngReturnId = 40050
    
    Case ID_ITEM_CADASTRO_CLASSE
      lngReturnId = 40060
    
    Case ID_ITEM_CADASTRO_SUBCLASSE
      lngReturnId = 40070
    
    Case ID_ITEM_CADASTRO_COR
      lngReturnId = 40080
    
    Case ID_ITEM_CADASTRO_TAMANHO
      lngReturnId = 40090
    
    Case ID_ITEM_CADASTRO_ETIQUETA_PRODUTO
      lngReturnId = 40100
    
    Case ID_ITEM_FORMATAR_ETIQUETA_PRODUTO
      lngReturnId = 40101
    
    Case ID_ITEM_CADASTRO_PESQUISA_1
      lngReturnId = 40130
    
    Case ID_ITEM_CADASTRO_PESQUISA_2
      lngReturnId = 40140
    
    Case ID_ITEM_CADASTRO_PESQUISA_3
      lngReturnId = 40150
    
    Case ID_ITEM_CADASTRO_BANCO
      lngReturnId = 40160
    
    Case ID_ITEM_CADASTRO_CONTA_CORRENTE
      lngReturnId = 40170
    
    Case ID_ITEM_CADASTRO_CARTAO
      lngReturnId = 40180
    
    Case ID_ITEM_CADASTRO_CAIXA
      lngReturnId = 40190
    
    Case ID_ITEM_CADASTRO_MOEDA
      lngReturnId = 40200
    
    Case ID_ITEM_CADASTRO_COTACAO
      lngReturnId = 40210
    
    Case ID_ITEM_CADASTRO_CLASSIFICACAO_FISCAL
      lngReturnId = 40220
    
    Case ID_ITEM_CADASTRO_CENTRO_CUSTO
      lngReturnId = 40230
    
    Case ID_ITEM_CADASTRO_RADIO
      lngReturnId = 0
    
    Case ID_ITEM_CADASTRO_TIPO_COMERCIAL
      lngReturnId = 0
    
    Case ID_ITEM_CADASTRO_AUT_PUBLICIDADE
      lngReturnId = 0
    
    Case ID_ITEM_CADASTRO_SUPERVISOR
      lngReturnId = 0
    
    Case ID_ITEM_CADASTRO_RETENCAO
      lngReturnId = 0
    
    Case ID_ITEM_CADASTRO_CODIGO_NBM
      lngReturnId = 0
    
    Case ID_ITEM_CADASTRO_GRUPO_FISCAL
      lngReturnId = 320084
    
    Case ID_ITEM_CADASTRO_MENSAGEM_NOTA_FISCAL
      lngReturnId = 320085
    
    Case ID_ITEM_CADASTRO_MALA_DIRETA_PREPARAR
      lngReturnId = 20199
    
    Case ID_ITEM_CADASTRO_MALA_DIRETA_MANUTENCAO
      lngReturnId = 20200
    
    Case ID_ITEM_CADASTRO_MALA_DIRETA_PREPARAR_REMETENTE
      lngReturnId = 20201
    
    Case ID_ITEM_CADASTRO_MALA_DIRETA_GRUPO
      lngReturnId = 20202
    
    Case ID_ITEM_CADASTRO_OPERACAO_ENTRADA
      lngReturnId = 40122
    
    Case ID_ITEM_CADASTRO_OPERACAO_SAIDA
      lngReturnId = 40124
    
    '----------------------------------------------------------------------------
    'TAB Movimento
    '----------------------------------------------------------------------------
    Case ID_ITEM_MOVIMENTO_VENDA_RAPIDA
      lngReturnId = 50010
    
    Case ID_ITEM_MOVIMENTO_ENTRADAS
      lngReturnId = 50020
    
    Case ID_ITEM_MOVIMENTO_SAIDAS
      lngReturnId = 50030
    
    Case ID_ITEM_MOVIMENTO_DEVOLUCOES
      lngReturnId = 50041
      
    Case ID_ITEM_REL_SAIDAS_ENTRADAS
      lngReturnId = 301103
      
    Case ID_ITEM_MOVIMENTO_ORDEM_SERVICO
      lngReturnId = 50070
    
    Case ID_ITEM_MOVIMENTO_PEDIDOS_WEB
      lngReturnId = 320042
        
    Case ID_ITEM_MOVIMENTO_MANUT_RESERVA
      lngReturnId = 320050
    
    Case ID_ITEM_MOVIMENTO_MANUT_CONSIG_ENTRADA
      lngReturnId = 0
    
    Case ID_ITEM_MOVIMENTO_FATUR_AUTO
      lngReturnId = 0
    
    Case ID_ITEM_MOVIMENTO_PREST_FORNEC
      lngReturnId = 0
    
    Case ID_ITEM_MOVIMENTO_IMPORTACAO
      lngReturnId = 0
    
    Case ID_ITEM_MOVIMENTO_APAGAR_EMP_ENTRADA
      lngReturnId = 50080
    
    Case ID_ITEM_MOVIMENTO_APAGAR_EMP_SAIDA
      lngReturnId = 50090
    
    Case ID_ITEM_MOVIMENTO_APAGAR_ENTRADA
      lngReturnId = 50080
    
    Case ID_ITEM_MOVIMENTO_APAGAR_SAIDA
      lngReturnId = 50090
    
    Case ID_ITEM_MOVIMENTO_APAGAR_MOVIMENTACAO
      lngReturnId = 50090
    
    Case ID_ITEM_MOVIMENTO_MANUT_CONSIGNACAO
      lngReturnId = 320045
    
    Case ID_ITEM_MOVIMENTO_MANUT_ORCAMENTO
      lngReturnId = 320046
    
    Case ID_ITEM_MOVIMENTO_TRANSF_FILIAL
      lngReturnId = 50040
    
    Case ID_ITEM_MOVIMENTO_EMPREST_ENTRADA
      lngReturnId = 50050
    
    Case ID_ITEM_MOVIMENTO_EMPREST_SAIDA
      lngReturnId = 50060
      
    Case ID_IMPORTA_GESTO
      lngReturnId = 0
    
    '----------------------------------------------------------------------------
    'TAB Preços
    '----------------------------------------------------------------------------
    Case ID_ITEM_PRECO_CRIAR_TAB
      lngReturnId = 70010
    
    Case ID_ITEM_PRECO_APAGAR_TAB
      lngReturnId = 70090
    
    Case ID_ITEM_PRECO_LANCAR
      lngReturnId = 70020
    
    Case ID_ITEM_PRECO_ALTERAR
      lngReturnId = 70030
    
    Case ID_ITEM_PRECO_ALTERAR_CALC
      lngReturnId = 0
    
    Case ID_ITEM_PRECO_CONFIG_TAB
      lngReturnId = 70040
    
    Case ID_ITEM_PRECO_COPIAR_TAB_IND
      lngReturnId = 70050
    
    Case ID_ITEM_PRECO_COPIAR_TAB_VALOR
      lngReturnId = 70060
    
    Case ID_ITEM_PRECO_COPIAR_TAB_CUSTO_MEDIO
      lngReturnId = 70070
    
    Case ID_ITEM_PRECO_CALC_PRECO
      lngReturnId = 70080
    
    Case ID_ITEM_PRECO_CALC_PRECO_SIMPLES
      lngReturnId = 0
    
    '----------------------------------------------------------------------------
    'TAB Estoque
    '----------------------------------------------------------------------------
    Case ID_ITEM_ESTOQUE_INFO_CONTAR
      lngReturnId = 60010
    
'''    Case ID_ITEM_ESTOQUE_ACERTAR_CONTAR
'''      lngReturnId = 60020
    
    Case ID_ITEM_ESTOQUE_INFO_CONTAR_GRADE
      lngReturnId = 60030
    
    Case ID_ITEM_ESTOQUE_ACERTAR_CONTAR_GRADE
      lngReturnId = 60040
    
    '----------------------------------------------------------------------------
    'TAB Financeiro
    '----------------------------------------------------------------------------
    Case ID_ITEM_FINANCEIRO_MOV_MANUAL_CAIXA
      lngReturnId = 80010
    
    Case ID_ITEM_FINANCEIRO_APAGA_LANC_CAIXA
      lngReturnId = 0
    
    Case ID_ITEM_FINANCEIRO_LANC_BANC
      lngReturnId = 80020
    
    Case ID_ITEM_FINANCEIRO_RECAL_SALDO
      lngReturnId = 80030
    
    Case ID_ITEM_FINANCEIRO_APAGA_LANC_BANC
      lngReturnId = 80040
    
    Case ID_ITEM_FINANCEIRO_CP_LANCAR
      lngReturnId = 20181
    
    Case ID_ITEM_FINANCEIRO_CP_GERAR
      lngReturnId = 20182
    
    Case ID_ITEM_FINANCEIRO_CP_MANUT
      lngReturnId = 20183
    
    Case ID_ITEM_FINANCEIRO_CP_APAGAR_PAGA
      lngReturnId = 20184
    
    Case ID_ITEM_FINANCEIRO_CR_LANCAR
      lngReturnId = 20185
    
    Case ID_ITEM_FINANCEIRO_CR_MANUT
      lngReturnId = 20186
    
    Case ID_ITEM_FINANCEIRO_CR_APAGAR_RECEBIDA
      lngReturnId = 20187
    
    Case ID_ITEM_FINANCEIRO_CR_LANCAR_CHEQUE_PRE
      lngReturnId = 20188
    
    Case ID_ITEM_FINANCEIRO_CR_MANUT_CHEQUE_PRE
      lngReturnId = 20189
    
    Case ID_ITEM_FINANCEIRO_CR_APAGAR_CHEQUE_PRE
      lngReturnId = 20190
    
    Case ID_ITEM_FINANCEIRO_CR_LANCAR_CARTAO
      lngReturnId = 20191
    
    Case ID_ITEM_FINANCEIRO_CR_MANUT_CARTAO
      lngReturnId = 20192
    
    Case ID_ITEM_FINANCEIRO_CR_APAGAR_CARTAO
      lngReturnId = 20193
    
    Case ID_ITEM_FINANCEIRO_CR_LANCAR_CONTA_CLIENTE
      lngReturnId = 20194
    
    Case ID_ITEM_FINANCEIRO_CR_MANUT_CONTA_CLIENTE
      lngReturnId = 20195
    
    Case ID_ITEM_FINANCEIRO_CR_APAGAR_CONTA_CLIENTE
      lngReturnId = 20196
    
    Case ID_ITEM_FINANCEIRO_CR_AUT_PUBLICIDADE
      lngReturnId = 0
    
    '----------------------------------------------------------------------------
    'TAB Gerador
    '----------------------------------------------------------------------------
    Case ID_ITEM_GERADOR_RELATORIO
      lngReturnId = 90010
    
    Case ID_ITEM_GERADOR_LAYOUT_NOTA
      lngReturnId = 90020
    
    Case ID_ITEM_GERADOR_LAYOUT_TICKET
      lngReturnId = 90020
    
    Case ID_ITEM_GERADOR_LAYOUT_BOLETO
      lngReturnId = 90030
    
    Case ID_ITEM_GERADOR_LAYOUT_CARNET
      lngReturnId = 90040
    
    Case ID_ITEM_GERADOR_ARQ_REC_ESTADUAL
      lngReturnId = 0
    
    Case ID_ITEM_GERADOR_EXP_BR_INFO
      lngReturnId = 0
    
    Case ID_ITEM_GERADOR_EXP_SADIG_WEB
      lngReturnId = 0
    
    Case ID_ITEM_GERADOR_EXP_PEARSON
      lngReturnId = 0
    
    '----------------------------------------------------------------------------
    'TAB Relatórios
    '----------------------------------------------------------------------------
    '---------------------------------------------------------------------------- [Serviços]
    Case ID_ITEM_REL_SERVICO_EXECUTADO
      lngReturnId = 300210
    
    Case ID_ITEM_REL_SERVICO_COMISSAO
      lngReturnId = 300220
    
    '---------------------------------------------------------------------------- [Produtos]
    Case ID_ITEM_REL_PRODUTO_GERAL
      lngReturnId = 300410
    
    Case ID_ITEM_REL_PRODUTO_GRADE
      lngReturnId = 300420
    
    '---------------------------------------------------------------------------- [Estoque]
    Case ID_ITEM_REL_ESTOQUE_GERAL
      lngReturnId = 300810
    
    Case ID_ITEM_REL_ESTOQUE_GRADE
      lngReturnId = 300820
    
    Case ID_ITEM_REL_ESTOQUE_ANALITICO
      lngReturnId = 300830
    
    Case ID_ITEM_REL_ESTOQUE_POR_FILIAL
      lngReturnId = 320043
    
    Case ID_ITEM_REL_ESTOQUE_FILIAL_PRECO
      lngReturnId = 320083
    
    Case ID_ITEM_REL_ESTOQUE_PRODUTO_COMPRAR
      lngReturnId = 300840
    
    Case ID_ITEM_REL_ESTOQUE_ACOMPANHAMENTO_PRODUTO
      lngReturnId = 300850
    
    Case ID_ITEM_REL_ESTOQUE_ACOMPANHAMENTO_ESTOQUE
      lngReturnId = 300860
    
    Case ID_ITEM_REL_ESTOQUE_REGISTRO_INVENTARIO
      lngReturnId = 300870
    
    Case ID_ITEM_REL_ESTOQUE_CONTAGEM
      lngReturnId = 300880
    
    Case ID_ITEM_REL_ESTOQUE_CONTAGEM_GRADE
      lngReturnId = 300890
    
    '---------------------------------------------------------------------------- [Compras e Vendas]
    Case ID_ITEM_REL_CV_VENDA
      lngReturnId = 301010
    
    Case ID_ITEM_REL_CV_VENDA_2
      lngReturnId = 0
    
    Case ID_ITEM_REL_CV_COMISSAO
      lngReturnId = 301020
    
    Case ID_ITEM_REL_CV_COMISSAO_RETENCAO
      lngReturnId = 0
    
    Case ID_ITEM_REL_CV_COMISSAO_VENDA_VENDEDOR
      lngReturnId = 0
    
    Case ID_ITEM_REL_CV_COMPRAS
      lngReturnId = 301030
    
    Case ID_ITEM_REL_CV_PRODUTO_COMPRAR_FATOR
      lngReturnId = 0
    
    Case ID_ITEM_REL_CV_VENDA_CLIENTE
      lngReturnId = 320047
    
    Case ID_ITEM_REL_CV_VENDA_VENDEDOR_COMISSAO
      lngReturnId = 320076
    
    Case ID_ITEM_REL_CV_VENDA_TAMANHO
      lngReturnId = 320048
    
    Case ID_ITEM_REL_CV_VENDA_EDITORA
      lngReturnId = 320066
    
    Case ID_ITEM_REL_CV_VENDA_FORNECEDOR
      lngReturnId = 0
    
    Case ID_ITEM_REL_CV_VENDA_PRODUTO_CONSIGNADO
      lngReturnId = 0
    
    Case ID_ITEM_REL_CV_PRESTACAO_CONTA
      lngReturnId = 0
    
    Case ID_ITEM_REL_CV_VENDA_POR_VENDEDOR
      lngReturnId = 320070
    
    '---------------------------------------------------------------------------- [Movimento]
    Case ID_ITEM_REL_MOV_ENTRADA
      lngReturnId = 301101
    
    Case ID_ITEM_REL_MOV_SAIDA
      lngReturnId = 301102
    
    Case ID_ITEM_REL_MOV_ACERTA_EMPREST_ENTRADA
      lngReturnId = 320039
    
    Case ID_ITEM_REL_MOV_ACERTA_EMPREST_SAIDA
      lngReturnId = 320040
    
    Case ID_ITEM_REL_MOV_ENTRADA_CONSIGNADA
      lngReturnId = 0
    
    '---------------------------------------------------------------------------- [Pessoas]
    Case ID_ITEM_REL_CLIENTE_FORNECEDOR
      lngReturnId = 300610
    
    Case ID_ITEM_REL_CONTATO_EFETUADO
      lngReturnId = 300620
    
    Case ID_ITEM_REL_CONTATO_DATA_ANIVERSARIO
      lngReturnId = 300630
    
    Case ID_ITEM_REL_USUARIO_FUNCIONARIO
      lngReturnId = 300712
    
    Case ID_ITEM_REL_LIVRO_PONTO
      lngReturnId = 305012
    
    '---------------------------------------------------------------------------- [Cadastro]
    Case ID_ITEM_REL_CADASTRO_CLASSE
      lngReturnId = 302500
    
    Case ID_ITEM_REL_CADASTRO_SUBCLASSE
      lngReturnId = 302600
    
    Case ID_ITEM_REL_CADASTRO_COR
      lngReturnId = 302800
    
    Case ID_ITEM_REL_CADASTRO_TAMANHO
      lngReturnId = 303000
    
    Case ID_ITEM_REL_CADASTRO_ETIQUETA_PRODUTO
      lngReturnId = 303200
    
    Case ID_ITEM_REL_CADASTRO_BANCO
      lngReturnId = 301400
    
    Case ID_ITEM_REL_CADASTRO_CARTAO
      lngReturnId = 301600
    
    Case ID_ITEM_REL_CADASTRO_MOEDA
      lngReturnId = 302000
    
    Case ID_ITEM_REL_CADASTRO_COTACAO
      lngReturnId = 301200
    
    Case ID_ITEM_REL_CADASTRO_CENTRO_CUSTO
      lngReturnId = 301800
    
    '---------------------------------------------------------------------------- [Financeiro]
    Case ID_ITEM_REL_FINANC_CAIXA
      lngReturnId = 302210
    
    Case ID_ITEM_REL_FINANC_CARTAO
      lngReturnId = 302210 'lngReturnId = 0         **** JULHO/2019 atribuído igual ao ID_ITEM_REL_FINANC_CAIXA
    
    Case ID_ITEM_REL_FINANC_LANC_BANCARIO
      lngReturnId = 302220
    
    Case ID_ITEM_REL_FINANC_SALDO_CC
      lngReturnId = 302230
    
    Case ID_ITEM_REL_FINANC_DIARIO_1
      lngReturnId = 302240
    
    Case ID_ITEM_REL_FINANC_DIARIO_2
      lngReturnId = 302250
    
    Case ID_ITEM_REL_FINANC_LUCRATIVIDADE
      lngReturnId = 302260
    
    Case ID_ITEM_REL_FINANC_GERAL
      lngReturnId = 302270
    
    Case ID_ITEM_REL_FINANC_RECEB_FORMA_PGTO
      lngReturnId = 320086
    
    Case ID_ITEM_REL_FINANC_FLUXO_CAIXA
      lngReturnId = 302400
    
    Case ID_ITEM_REL_CP_PAGAR_DATA_VCTO
      lngReturnId = 302281
    
    Case ID_ITEM_REL_CP_PAGAR_FORNECEDOR
      lngReturnId = 302282
    
    Case ID_ITEM_REL_CP_PAGAR_GERAL_FILIAL
      lngReturnId = 302283
    
    Case ID_ITEM_REL_CP_PAGAR_CENTRO_CUSTO
      lngReturnId = 302284
    
    Case ID_ITEM_REL_CP_PAGAS_FORNECEDOR
      lngReturnId = 302285
    
    Case ID_ITEM_REL_CP_PAGAS_DATA_PGTO
      lngReturnId = 302286
    
    Case ID_ITEM_REL_CP_PAGAS_CENTRO_CUSTO
      lngReturnId = 302287
    
    Case ID_ITEM_REL_CP_CONTROLE_CENTRO_CUSTO
      lngReturnId = 0
    
    Case ID_ITEM_REL_CP_CENTRO_CUSTO_COMPETENCIA
      lngReturnId = 0
    
    Case ID_ITEM_REL_CR_RECEBER_DATA_VCTO
      lngReturnId = 302291
    
    Case ID_ITEM_REL_CR_RECEBER_CLIENTE
      lngReturnId = 302292
    
    Case ID_ITEM_REL_CR_RECEBER_VENDEDOR
      lngReturnId = 302293
    
    Case ID_ITEM_REL_CR_RECEBIDA_DATA_RECEBIMENTO
      lngReturnId = 302294
    
    Case ID_ITEM_REL_CR_RECEBIDA_VENDEDOR
      lngReturnId = 302295
    
    Case ID_ITEM_REL_CR_RECEBIDA_CLIENTE
      lngReturnId = 302296
    
    Case ID_ITEM_REL_CR_CHEQUE_PRE
      lngReturnId = 302297
    
    Case ID_ITEM_REL_CR_CARTAO
      lngReturnId = 302298
    
    Case ID_ITEM_REL_CR_CONTA_CLIENTE
      lngReturnId = 302299
    
    Case ID_ITEM_REL_CR_EMISSAO_BOLETO
      lngReturnId = 302300
    
    Case ID_ITEM_REL_CR_EMISSAO_CARNET
      lngReturnId = 302301
      
    Case ID_ITEM_REL_CR_LANCAMENTOS_DATA_EMISSAO
      lngReturnId = 302302
    
    '----------------------------------------------------------------------------
    Case ID_ITEM_REL_MALA_DIRETA
      lngReturnId = 303400
    
    Case ID_ITEM_REL_MALA_DIRETA_GERAR_ARQUIVO
      lngReturnId = 0
    
    Case ID_ITEM_REL_FOLHA_PGTO
      lngReturnId = 0
    
    Case ID_ITEM_REL_AUTORIZACAO
      lngReturnId = 0
    
    Case ID_ITEM_REL_MALA_DIRETA_AUTORIZACAO
      lngReturnId = 0
    
    '---------------------------------------------------------------------------- [Gráfico]
    Case ID_ITEM_REL_GRAFICO_COMPARATIVO_CV
      lngReturnId = 304410
    
    Case ID_ITEM_REL_GRAFICO_VENDA_CLASSE_PERIODO
      lngReturnId = 304420
    
    Case ID_ITEM_REL_GRAFICO_VENDA_PRODUTO_MENSAL
      lngReturnId = 304430
      
    Case ID_ITEM_REL_GRAFICO4_VENDA_PRODUTOS
      lngReturnId = 304440
      
    Case ID_ITEM_REL_GRAFICO5_COMPRA_FORNECEDORES
      lngReturnId = 304450
      
    Case ID_ITEM_REL_ESTRATEGICO_AVISO_AQUISICAO
      lngReturnId = 304460
      
    Case ID_ITEM_REL_GRAFICO6_VENDA_CLIENTES
      lngReturnId = 304470

    '----------------------------------------------------------------------------
    Case ID_ITEM_REL_NSU_CORRELACAO
      lngReturnId = 0
    
    '---------------------------------------------------------------------------- [Preços]
    Case ID_ITEM_REL_PRECO_LISTA
      lngReturnId = 304900
    
    Case ID_ITEM_REL_PRECO_LOCAL_PRODUTO
      lngReturnId = 320051
    
    '----------------------------------------------------------------------------
  End Select
  
  'Retorno da função
  MenuActiveBarId = lngReturnId
  
End Function

'29/01/2009 - mpdea
'Define permissões de acesso ao menu
Public Sub SetMenuAcesso()
  Dim obj_ribbon_bar As RibbonBar
  Dim obj_tab As RibbonTab
  Dim obj_group As RibbonGroup
  Dim obj_control As CommandBarControl
  Dim obj_control_2 As CommandBarControl
  
  Dim lng_tab As Long
  Dim lng_group As Long
  Dim lng_control As Long
  
  Dim clc_id As New Collection
  Dim rsAcessos As Recordset
  Dim rsProg As Recordset
  Dim lng_active_bar_tool_id As Long
  Dim lng_command_bar_control_id As Long
  Dim bln_visible As Boolean
    
  
  On Error GoTo ErrHandle

  
  Screen.MousePointer = vbHourglass
  
  'Obtém todos os IDs dos contoles
  Set obj_ribbon_bar = MenuRibbonBar
  With obj_ribbon_bar
    'Reseta menu
    .Reset
    'TAB
    For lng_tab = 0 To .TabCount - 1
      Set obj_tab = .Tab(lng_tab)
      'GROUP
      For lng_group = 0 To obj_tab.Groups.GroupCount - 1
        Set obj_group = obj_tab.Groups(lng_group)
        'CONTROL
        For lng_control = 0 To obj_group.Count - 1
          Set obj_control = obj_group.Item(lng_control)
          ListMenuItem obj_control, clc_id
        Next
      Next
    Next
  End With
  
  '--------------------------------------------------------------------------------
  '22/01/2003 - mpdea
  'Verifica modo limitado do Quick Store
  If gblnQuickFull Then
    Set rsProg = db.OpenRecordset("SELECT * FROM ZZZProgramas ORDER BY ToolID", dbOpenDynaset)
    Set rsAcessos = db.OpenRecordset("SELECT * FROM Acessos WHERE Usuário = " & CStr(gnUserCode), dbOpenDynaset)

    'Seta menus
    For lng_control = 1 To clc_id.Count
      lng_command_bar_control_id = clc_id(lng_control)
      lng_active_bar_tool_id = MenuActiveBarId(lng_command_bar_control_id)
      
      Dim bln_teste As Boolean
      bln_teste = lng_command_bar_control_id = 1318
      
      'Localiza controle
      Set obj_control = obj_ribbon_bar.FindControl(, lng_command_bar_control_id, , True)
      If Not obj_control Is Nothing Then
        'Verifica permissão
        If gbSuperUser Then
          obj_control.Parameter = PERMISSION_COMPLETO 'Acesso completo
        Else
          rsProg.FindFirst "ToolID = " & CStr(lng_active_bar_tool_id)
          If Not rsProg.NoMatch Then
            rsAcessos.FindFirst "Numero = " & rsProg("Número").Value
            If Not rsAcessos.NoMatch Then
              obj_control.Parameter = PERMISSION_VISIVEL 'Somente acesso
              If rsAcessos("Gravar").Value = True Then
                obj_control.Parameter = PERMISSION_GRAVAR 'Acesso para gravar
              End If
              If rsAcessos("Apagar").Value = True Then
                obj_control.Parameter = PERMISSION_COMPLETO 'Acesso completo
              End If
            Else
              obj_control.Parameter = PERMISSION_SEM_ACESSO
            End If
          Else
            obj_control.Parameter = PERMISSION_SEM_ACESSO
          End If
        End If
      End If
    Next
  
    'Fecha tabelas
    rsAcessos.Close
    Set rsAcessos = Nothing
    rsProg.Close
    Set rsProg = Nothing
    
    '----------------------------------------------------------------------------
    'Exceções
    '----------------------------------------------------------------------------

    'Menu principal
'''    SetMenuPermission ID_ITEM_CADASTRO_CODIGO_NBM, PERMISSION_VISIVEL
    SetMenuPermission ID_ITEM_ARQUIVO_ESTACOES_CONECTADAS, PERMISSION_VISIVEL
    SetMenuPermission ID_ITEM_ARQUIVO_EXPORTAR_BASE, IIf(gbSuperUser, PERMISSION_VISIVEL, PERMISSION_SEM_ACESSO)
    SetMenuPermission ID_ITEM_ARQUIVO_BACKUP, IIf(gbSuperUser, PERMISSION_VISIVEL, PERMISSION_SEM_ACESSO)
    SetMenuPermission ID_ITEM_ARQUIVO_REPARAR_BASE, IIf(gbSuperUser, PERMISSION_VISIVEL, PERMISSION_SEM_ACESSO)
    SetMenuPermission ID_ITEM_ARQUIVO_COMPACTAR_BASE, IIf(gbSuperUser, PERMISSION_VISIVEL, PERMISSION_SEM_ACESSO)
    
    '16/01/2003 - mpdea
    'Corrigido exibição da ajuda entre convênios
    bln_visible = (gnNumConvenio > 31 And Dir(gsHelpConv) <> "" And gsHelpConv <> "")
    SetMenuPermission ID_ITEM_AJUDA_INSTITUCIONAL, IIf(bln_visible, PERMISSION_VISIVEL, PERMISSION_SEM_ACESSO)
    
'    frmMain.ActiveBar1.Tools("miTips").Visible = (Dir(gsTipFile) <> "")
'    frmMain.ActiveBar1.Tools("miTips").Checked = GetSetting("QuickStore", "Options", "Show Tips", 1)
'    .Tools("miAjudaRegistro").Enabled = True
'    .Tools("miAjudaRegistro").Tag = 0
    
'''    'Área de transferência
'''    SetMenuPermission ID_ITEM_INICIO_COLAR, PERMISSION_VISIVEL
'''    SetMenuPermission ID_ITEM_INICIO_RECORTAR, PERMISSION_VISIVEL
'''    SetMenuPermission ID_ITEM_INICIO_COPIAR, PERMISSION_VISIVEL

    '08/08/2005 - Daniel
    'A configuração das impessoras ficará habilitada somente
    'liberação da autorização
'''    SetMenuPermission ID_ITEM_INICIO_PARAM_CONFIG_IMPRESSORA, PERMISSION_VISIVEL

    'Grade
    If Not gbGrade Then
      SetMenuPermission ID_ITEM_CADASTRO_COR, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_CADASTRO_TAMANHO, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_ESTOQUE_INFO_CONTAR_GRADE, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_ESTOQUE_ACERTAR_CONTAR_GRADE, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_REL_CADASTRO_COR, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_REL_CADASTRO_TAMANHO, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_REL_PRODUTO_GRADE, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_REL_ESTOQUE_GRADE, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_REL_ESTOQUE_CONTAGEM_GRADE, PERMISSION_SEM_ACESSO
    End If
    
    'Serviço
    If Not gbServico Then
      SetMenuPermission ID_ITEM_CADASTRO_SERVICO, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_MOVIMENTO_ORDEM_SERVICO, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_REL_SERVICO_EXECUTADO, PERMISSION_SEM_ACESSO
      SetMenuPermission ID_ITEM_REL_SERVICO_COMISSAO, PERMISSION_SEM_ACESSO
    End If
    
    '16/12/2009 - mpdea
    'Caixa
    If Not gbCaixas Then
      SetMenuPermission ID_ITEM_CADASTRO_CAIXA, PERMISSION_SEM_ACESSO
    End If
    
    'Pesquisa 1, 2 e 3
    If gsPesq1 <> "" Then
      Set obj_control = MenuRibbonBar.FindControl(, ID_ITEM_CADASTRO_PESQUISA_1, , True)
      If Not obj_control Is Nothing Then
        If obj_control.Parameter <> PERMISSION_SEM_ACESSO Then
          obj_control.Caption = "&1-" & gsPesq1 & "..."
        End If
      End If
    End If
    If gsPesq2 <> "" Then
      Set obj_control = MenuRibbonBar.FindControl(, ID_ITEM_CADASTRO_PESQUISA_2, , True)
      If Not obj_control Is Nothing Then
        If obj_control.Parameter <> PERMISSION_SEM_ACESSO Then
          obj_control.Caption = "&2-" & gsPesq2 & "..."
        End If
      End If
    End If
    If gsPesq3 <> "" Then
      Set obj_control = MenuRibbonBar.FindControl(, ID_ITEM_CADASTRO_PESQUISA_3, , True)
      If Not obj_control Is Nothing Then
        If obj_control.Parameter <> PERMISSION_SEM_ACESSO Then
          obj_control.Caption = "&3-" & gsPesq3 & "..."
        End If
      End If
    End If
    
    '31/07/2002 - mpdea
    'Menu de Gerenciamento dos Pedidos da Loja Virtual
    If Not gblnWorkWeb Then
      SetMenuPermission ID_ITEM_MOVIMENTO_PEDIDOS_WEB, PERMISSION_SEM_ACESSO
    End If
    
    '17/09/2009 - mpdea
    'Menu para uso de Nota Fiscal Eletrônica
    If Not gblnNFe Then
      SetMenuPermission ID_ITEM_MOVIMENTO_NOTA_FISCAL_ELETRONICA, PERMISSION_SEM_ACESSO
    End If
    
    'Menu adicionado que utiliza a mesma permissão do cálculo de preço de venda
    Set obj_control = MenuRibbonBar.FindControl(, ID_ITEM_PRECO_CALC_PRECO_SIMPLES, , True)
    If Not obj_control Is Nothing Then
      Set obj_control_2 = MenuRibbonBar.FindControl(, ID_ITEM_PRECO_CALC_PRECO, , True)
      If Not obj_control_2 Is Nothing Then
        obj_control.Parameter = obj_control_2.Parameter
      End If
    End If

    '24/01/2003 - mpdea
    'Menu adicionado que utiliza a mesma permissão da alteração de preços
    Set obj_control = MenuRibbonBar.FindControl(, ID_ITEM_PRECO_ALTERAR_CALC, , True)
    If Not obj_control Is Nothing Then
      Set obj_control_2 = MenuRibbonBar.FindControl(, ID_ITEM_PRECO_ALTERAR, , True)
      If Not obj_control_2 Is Nothing Then
        obj_control.Parameter = obj_control_2.Parameter
      End If
    End If
    
    '20/12/2007 - Anderson
    'Implementação do relatório de NSU
    bln_visible = gstrGetEstadoFilial(gnCodFilial) = "SC"
    SetMenuPermission ID_ITEM_REL_NSU_CORRELACAO, IIf(bln_visible, PERMISSION_VISIVEL, PERMISSION_SEM_ACESSO)
  Else
    Call SetMenuAcessoLimitado(clc_id)
  End If
  Set obj_ribbon_bar = Nothing

  '----------------------------------------------------------------------------
  'Customizações
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '17/02/2004 - Daniel
  'Case: STC Sistema Trídio de Comunicação (Rádio Difusora Caxias do Sul)
  '----------------------------------------------------------------------------
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS39823-684")
  'Chamada do Rel. de Autorização de Publicidade
  SetMenuPermission ID_ITEM_REL_AUTORIZACAO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '13/04/2004 - Chamada para o Rel de Mala Direta
  SetMenuPermission ID_ITEM_REL_MALA_DIRETA_AUTORIZACAO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '31/03/2004 - Daniel
  'Inclusão de Cadastro de Rádio e Tipo Comercial
  '23/07/2004 - Daniel
  'Inclusão do Cadastro de Autorizações
  SetMenuPermission ID_ITEM_CADASTRO_RADIO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_CADASTRO_TIPO_COMERCIAL, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_CADASTRO_AUT_PUBLICIDADE, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '08/04/2004 - Daniel
  'Inclusão da Tela de Consulta de Contrato
  SetMenuPermission ID_ITEM_FINANCEIRO_CR_AUT_PUBLICIDADE, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '30/07/2004 - Daniel
  'Adicionado cadastro de supervisores
  SetMenuPermission ID_ITEM_CADASTRO_SUPERVISOR, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '02/08/2004 - Daniel
  'Adicionado em parâmetros faturamento automático
  SetMenuPermission ID_ITEM_INICIO_PARAM_FATURAMENTO_AUTO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '13/08/2004 - Daniel
  'Adicionado em Movimento faturamento automático
  SetMenuPermission ID_ITEM_MOVIMENTO_FATUR_AUTO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '21/03/2005 - Daniel
  'Case..........: Bem Me Quer
  'Cadastro de Retenções (Ficou em Stand by)
  '
  '15/06/2005 - Daniel
  'Solicitação...: Adicionado o QS71147-191
  '
  '09/10/2007 - Anderson
  'Solicitação...: Adicionado o QS73173-153
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS40915-699", "QS71147-191", "QS73173-153")
  SetMenuPermission ID_ITEM_CADASTRO_RETENCAO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '24/01/2005 - Daniel
  'Case: Castro Construções
  'Importador de Clientes e Produtos
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS71115-747")
  SetMenuPermission ID_ITEM_MOVIMENTO_IMPORTACAO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '15/09/2004 - Daniel
  'Case: Livraria Resultado
  'Preparação do ambiente personalizado para o cliente
  'Livraria Resultado
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS40590-987")
  SetMenuPermission ID_ITEM_INICIO_PARAM_DEVOL_MATERIAL, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_MOVIMENTO_PREST_FORNEC, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_MOVIMENTO_MANUT_CONSIG_ENTRADA, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_REL_MOV_ENTRADA_CONSIGNADA, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '24/05/2004 - Daniel
  'Case: Bic Amazônia
  'Relatório de Geração de Arquivo para Folha de Pagamento
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS35509-939", "QS37715-731")
  SetMenuPermission ID_ITEM_REL_FOLHA_PGTO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '26/02/2004 - Daniel
  'Case: PSV
  'Chamada da tela de Manutenção de Reservas
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS35552-811", "QS37705-639", "QS37825-830", "QS38933-772", "QS39369-521")
  SetMenuPermission ID_ITEM_MOVIMENTO_MANUT_RESERVA, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '07/07/2004 - Daniel
  'Classificação de Clientes - Case: TV Shopping
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS39945-043", "QS40449-276", "QS39944-959")
  SetMenuPermission ID_ITEM_INICIO_PARAM_CLASS_CLIENTE, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_REL_MALA_DIRETA_GERAR_ARQUIVO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '09/09/2004 - Daniel
  'Case: Livraria Resultado
  'Chamada da tela do Rel. de Vendas por Fornecedores
  '14/10/2004 - Daniel
  'Adicionado a chamada para o Rel. de Prestação de Contas com Fornecedores
  '17/08/2007 - Anderson
  'Alteração realizada para customização de relatório da Nutricare (QS73086-490)
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS40590-987")
  SetMenuPermission ID_ITEM_REL_CV_VENDA_PRODUTO_CONSIGNADO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_REL_CV_PRESTACAO_CONTA, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  SetMenuPermission ID_ITEM_REL_MOV_ENTRADA_CONSIGNADA, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS40590-987", "QS73086-490", "QS73032-694")
  SetMenuPermission ID_ITEM_REL_CV_VENDA_FORNECEDOR, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '28/04/2005 - Daniel
  'Relatório de comissões com retenção sobre cartões
  'Case..........: Bem Me Quer
  '
  '15/06/2005 - Daniel
  'Solicitação...: Adicionado o QS71147-191
  '
  '09/10/2007 - Anderson
  'Solicitação...: Adicionado o QS73173-153
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS40915-699", "QS71147-191", "QS73173-153")
  SetMenuPermission ID_ITEM_REL_CV_COMISSAO_RETENCAO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '19/01/2006 - mpdea
  'Relatório de Estoque das Filiais e Preço (Personalizado)
  'Solicitante: Cliente Kilouça (QS71271-970)
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS71271-970")
  SetMenuPermission ID_ITEM_REL_ESTOQUE_FILIAL_PRECO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '23/05/2007 - Anderson
  'Customização do relatório de vendas
  'Solicitante: LL Comércio de Ferramentas LTDA (QS73022-602)
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS73022-602")
  SetMenuPermission ID_ITEM_REL_CV_VENDA_2, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '23/05/2007 - Anderson
  'Exportação de Dados para sistema da Brasil Informática
  'Solicitante: Anistex Ind. e Com. Ltda (QS31935-863)
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS31935-863")
  SetMenuPermission ID_ITEM_GERADOR_EXP_BR_INFO, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '23/05/2007 - Anderson
  'Exportação de Dados para sistema da SadigWeb
  'Solicitante: Gurgel e Leite (QS31935-863)
  bln_visible = gbSuperUser And g_blnSadigWeb
  SetMenuPermission ID_ITEM_GERADOR_EXP_SADIG_WEB, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '10/01/2008 - Anderson
  'Exportação de Dados para sistema da Pearson
  'Solicitante: Technomax
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS73234-876", "QS73235-960", "QS73236-044", "QS73237-128", "QS73238-212", "QS73239-296", "QS73240-632")
  SetMenuPermission ID_ITEM_GERADOR_EXP_PEARSON, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '08/08/2007 - Anderson
  'Relatório de comissões
  'Solicitante: Candy Clean (QS31935-863)
  bln_visible = gbSuperUser And CheckSerialCaseMod("QS37957-281")
  SetMenuPermission ID_ITEM_REL_CV_COMISSAO_VENDA_VENDEDOR, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  '----------------------------------------------------------------------------
  '30/10/2007 - Anderson
  'Relatório de Produtos a Comprar
  'Solicitante: Kings Cross (QS38714-658,QS38393-282)
  bln_visible = gbSuperUser And g_bolRelatorioCompra
  SetMenuPermission ID_ITEM_REL_CV_PRODUTO_COMPRAR_FATOR, IIf(bln_visible, PERMISSION_COMPLETO, PERMISSION_SEM_ACESSO)
  '----------------------------------------------------------------------------
  
  ' Help Online (acesso para todos)
''  SetMenuPermission ID_ACESSO_HELP_QUICK, PERMISSION_VISIVEL
  SetMenuPermission ID_ITEM_INICIO_STANDBY, PERMISSION_VISIVEL
  
  
  'Ajusta o menu de acordo com os controles visíveis
  AdjustMenu
  
  'Ativa Tab inicial
  MenuRibbonBar.Tab(0).Selected = True
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
ErrHandle:
  gsTitle = LoadResString(201)
  gsMsg = "Erro durante a habilitação de menus."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Screen.MousePointer = vbDefault
  
  '22/01/2003 - mpdea
  'Finaliza aplicação em caso de erro ao habilitar menus em modo limitado
  If Not gblnQuickFull Then
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Not ws Is Nothing Then ws.Close
    Set ws = Nothing
    End
  End If
  
End Sub

'30/01/2009 - mpdea
'Adaptado para novo menu
'
'03/10/2003 - mpdea
'Alterado para utilizar serviços
'
'16/01/2003 - mpdea
'Desabilita funções do Quick Store em modo limitado
Private Sub SetMenuAcessoLimitado(ByVal clcMenuId As Collection)
  Dim obj_ribbon_bar As RibbonBar
  Dim obj_control As CommandBarControl
  Dim lng_control As Long
  Dim lng_id As Long

  'Seta menus
  Set obj_ribbon_bar = MenuRibbonBar
  For lng_control = 1 To clcMenuId.Count
    'Localiza controle
    lng_id = clcMenuId(lng_control)
    Set obj_control = obj_ribbon_bar.FindControl(, lng_id, , True)
    If Not obj_control Is Nothing Then
      'Padrão
      obj_control.Parameter = PERMISSION_COMPLETO
    End If
  Next
  Set obj_ribbon_bar = Nothing

  'Exceções
  SetMenuPermission ID_ITEM_INICIO_LIVRO_PONTO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_INICIO_PARAM_IMPOSTO_ESTADUAL, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_INICIO_PARAM_EMPRESA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_INICIO_PARAM_CLASS_CLIENTE, PERMISSION_SEM_ACESSO
  
  SetMenuPermission ID_ITEM_CADASTRO_GRUPO_FISCAL, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_MENSAGEM_NOTA_FISCAL, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_POPUP_CADASTRO_PESQUISA, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_ITEM_CADASTRO_COR, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_TAMANHO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_MOEDA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_TRANSPORTADORA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_OPERACAO_ENTRADA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_OPERACAO_SAIDA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_BANCO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_CONTA_CORRENTE, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_COTACAO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_CLASSIFICACAO_FISCAL, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_CADASTRO_CENTRO_CUSTO, PERMISSION_SEM_ACESSO
  
  SetMenuPermission ID_ITEM_MOVIMENTO_TRANSF_FILIAL, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_POPUP_MOVIMENTO_EMPRESTIMO, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_POPUP_MOVIMENTO_MANUT, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_ITEM_MOVIMENTO_PEDIDOS_WEB, PERMISSION_SEM_ACESSO

  SetMenuPermission ID_ITEM_FINANCEIRO_LANC_BANC, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_FINANCEIRO_RECAL_SALDO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_FINANCEIRO_APAGA_LANC_BANC, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_POPUP_FINANCEIRO_CP_MOVIMENTO, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_POPUP_FINANCEIRO_CR_MOVIMENTO, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_POPUP_FINANCEIRO_CR_CHEQUE_PRE, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_POPUP_FINANCEIRO_CR_CARTAO, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_POPUP_FINANCEIRO_CR_CONTA_CLIENTE, PERMISSION_SEM_ACESSO, True
  
  SetMenuPermission ID_ITEM_PRECO_CRIAR_TAB, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_APAGAR_TAB, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_LANCAR, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_ALTERAR, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_ALTERAR_CALC, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_COPIAR_TAB_IND, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_COPIAR_TAB_VALOR, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_COPIAR_TAB_CUSTO_MEDIO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_CALC_PRECO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_PRECO_CALC_PRECO_SIMPLES, PERMISSION_SEM_ACESSO

  SetMenuPermission ID_ITEM_ESTOQUE_INFO_CONTAR, PERMISSION_SEM_ACESSO
'''  SetMenuPermission ID_ITEM_ESTOQUE_ACERTAR_CONTAR, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_ESTOQUE_INFO_CONTAR_GRADE, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_ESTOQUE_ACERTAR_CONTAR_GRADE, PERMISSION_SEM_ACESSO

  SetMenuPermission ID_ITEM_GERADOR_LAYOUT_NOTA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_GERADOR_LAYOUT_BOLETO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_GERADOR_LAYOUT_CARNET, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_GERADOR_RELATORIO, PERMISSION_SEM_ACESSO
  
  SetMenuPermission ID_ITEM_REL_SERVICO_COMISSAO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_CADASTRO_COR, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_CADASTRO_TAMANHO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_CADASTRO_BANCO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_CADASTRO_MOEDA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_CADASTRO_COTACAO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_CADASTRO_CENTRO_CUSTO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_AUTORIZACAO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_PRODUTO_GRADE, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_ESTOQUE_GRADE, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_ESTOQUE_CONTAGEM, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_ESTOQUE_CONTAGEM_GRADE, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_ESTOQUE_GERAL, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_CV_COMISSAO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_MOV_ACERTA_EMPREST_ENTRADA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_MOV_ACERTA_EMPREST_SAIDA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_LIVRO_PONTO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_FINANC_LANC_BANCARIO, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_ITEM_REL_FINANC_SALDO_CC, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_POPUP_RELATORIO_FINANC_CONTA_PAGAR, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_POPUP_RELATORIO_FINANC_CONTA_RECEBER, PERMISSION_SEM_ACESSO, True
  SetMenuPermission ID_ITEM_REL_FINANC_FLUXO_CAIXA, PERMISSION_SEM_ACESSO
  SetMenuPermission ID_POPUP_RELATORIO_GRAFICO, PERMISSION_SEM_ACESSO, True
    
End Sub

'29/01/2009 - mpdea
'Seta parâmetro de permissão para o menu ser visível ou não
'Parâmetros: lngId - Id do menu
'            strParameter - Valor da permissão (-1 = Sem acesso, 0 - Visível, 10 - Gravar, 11 - Total)
'            blnRecursive - Se verdadeiro aplicará a permissão a todos os sub itens
Private Function SetMenuPermission(ByVal lngID As Long, ByVal strParameter As String, Optional ByVal blnRecursive As Boolean = False) As CommandBarControl
  Dim obj_control As CommandBarControl
  Dim obj_control_2 As CommandBarControl
  Dim lng_control As Long
  
  Set obj_control = MenuRibbonBar.FindControl(, lngID, , True)
  If Not obj_control Is Nothing Then obj_control.Parameter = strParameter
  
  'Recursivo
  If blnRecursive Then
    If obj_control.Type = xtpControlPopup Then
      For lng_control = 1 To obj_control.CommandBar.Controls.Count
        Set obj_control_2 = obj_control.CommandBar.Controls.Item(lng_control)

        'Aplica permissão aos sub itens
        SetMenuPermission obj_control_2.id, strParameter, blnRecursive
      Next
    End If
  End If
  
  'Retorno da função
  Set SetMenuPermission = obj_control
End Function

'29/01/2009 - mpdea
'Função recursiva para obter os IDs dos controles
Private Sub ListMenuItem(ByVal objControl As CommandBarControl, ByVal objCollection As Collection)
  Dim obj_control As CommandBarControl
  Dim lng_control As Long
  
  Select Case objControl.Type
    Case xtpControlButton
      objCollection.Add objControl.id
      
    Case xtpControlPopup
      For lng_control = 1 To objControl.CommandBar.Controls.Count
        Set obj_control = objControl.CommandBar.Controls.Item(lng_control)
        
        'Lista sub itens
        ListMenuItem obj_control, objCollection
      Next
  
  End Select

End Sub

'29/01/2009 - mpdea
'Ajusta o menu de acordo com os controles visíveis
Private Sub AdjustMenu()
  Dim obj_ribbon_bar As RibbonBar
  Dim obj_tab As RibbonTab
  Dim obj_group As RibbonGroup
  Dim obj_control As CommandBarControl
  Dim obj_control_pop As CommandBarPopup
  
  Dim lng_tab As Long
  Dim lng_group As Long
  Dim lng_control As Long
  
  Dim bln_tab_visible As Boolean
  Dim bln_group_visible As Boolean
  Dim bln_control_visible As Boolean
  
  
  'Valor padrão
  bln_tab_visible = False
  bln_group_visible = False
  bln_control_visible = False
  
  'Analisa controles
  Set obj_ribbon_bar = MenuRibbonBar
  With obj_ribbon_bar
    'Menu principal
    Set obj_control = .FindControl(, ID_SYSTEM_CONTROL, , True)
    Call AdjustMenuItem(obj_control)
    
    'TAB
    For lng_tab = 0 To .TabCount - 1
      Set obj_tab = .Tab(lng_tab)
      
      'GROUP
      For lng_group = 0 To obj_tab.Groups.GroupCount - 1
        Set obj_group = obj_tab.Groups(lng_group)
        
        'CONTROL
        For lng_control = 0 To obj_group.Count - 1
          Set obj_control = obj_group.Item(lng_control)
          bln_control_visible = AdjustMenuItem(obj_control)
          
          'CONTROL
          obj_control.Visible = bln_control_visible
          obj_control.Enabled = bln_control_visible
          If bln_control_visible Then bln_group_visible = True
          bln_control_visible = False
        Next
        
        'GROUP
        obj_group.Visible = bln_group_visible
        If bln_group_visible Then bln_tab_visible = True
        bln_group_visible = False
      Next
      
      'TAB
      obj_tab.Visible = bln_tab_visible
      bln_tab_visible = False
    Next
  End With
  
End Sub

'29/01/2009 - mpdea
'Função recursiva para ajustar o menu de acordo com os controles visíveis
'Parâmetro: objControl - Controle a ser ajustado
'Retorno: Verdadeiro caso tenha controle vísivel, falso caso contrário
Private Function AdjustMenuItem(ByVal objControl As CommandBarControl) As Boolean
  Dim obj_control As CommandBarControl
  Dim lng_control As Long
  Dim bln_control_visible As Boolean
  Dim bln_return As Boolean
  
  
  'Valor padrão
  bln_return = False
  bln_control_visible = False
  
  Select Case objControl.Type
    Case xtpControlButton
      'Flag para menu visível
      'A propriedade Visible do controle só é True quando o controle é visível na tela,
      'caso esteja em outra tab é marcado como False
      bln_return = objControl.Parameter <> PERMISSION_SEM_ACESSO
      
    Case xtpControlPopup
      For lng_control = 1 To objControl.CommandBar.Controls.Count
        Set obj_control = objControl.CommandBar.Controls.Item(lng_control)
        
        'Ajusta sub itens
        bln_control_visible = AdjustMenuItem(obj_control)
        
        'CONTROL
        obj_control.Visible = bln_control_visible
        obj_control.Enabled = bln_control_visible
        If bln_control_visible Then bln_return = True
      Next
  
  End Select
  
  'Retorno da função
  AdjustMenuItem = bln_return

End Function

