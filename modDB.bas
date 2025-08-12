Attribute VB_Name = "modDB"
Option Explicit

Public Function AlteraDB() As Boolean
  Dim nPhase As Integer
  Dim intX As Integer
    
  On Error GoTo ErrHandler
  
  Screen.MousePointer = vbHourglass
  
  Call ws.BeginTrans
  
  Call AlteraDB5(nPhase)
  
  '22/05/2007 - Anderson
  'Criação de uma procedure pois a funcão AlteraDB passou do limite permitido.
  Call AlteraDB2(nPhase)
  
  nPhase = nPhase + 1
  Call AlteraDBNFe
  
  '13/11/2014 - Eduardo Franco
  'Criação de outra procedure pois a funcão AlteraDB2 passou do limite permitido.
  Call AlteraDB3(nPhase)
  
  '16/05/2007 - Anderson
  '
  '386. Informar o número do bordero gerado para o título
  '     Tabela     : CNAB_Bordero
  '     Finalidade : Informar o número do bordero gerado para o título
  '     Solicitante: Technomax - Cliente Agrotama (QS73073-894)
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CNAB_Bordero") Then
    If Not gbCreateField("Contas a Receber", "CNAB_Bordero", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '17/05/2007 - Anderson
  '
  '387. Inclusão de registro para exibição de contatos efetuados
  '     Tabela     : ZZZProgramas
  '     Finalidade : Inclusão de novo programa
  '     Solicitante: Supri Print
  nPhase = nPhase + 1
  If Not AddFileZZZProgramas17 Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
   
  'Atualiza informações
  Call ws.CommitTrans
  
  'Inicia Transação
  Call ws.BeginTrans
  
  '17/05/2007 - Anderson
  '
  '388. Campo utilizado para informar o código CDRC do funcionário, conforme solicitação do cliente Gurgel & Leite para exportação de dados
  'no sistema Sadigweb. Este campo estará disponível apenas para o cliente Gurgel & Leite.
  '     Tabela     : Funcionários
  '     Finalidade : Informar o código CDRC do funcionário para ser exportado no arquivo de texto do SadigWeb
  '     Solicitante: Gurgel & Leite Comércio de Produtos Veterinários Ltda (QS31734-765)
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "SadigWeb_CDRC") Then
    If Not gbCreateField("Funcionários", "SadigWeb_CDRC", dbText, 20) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '17/05/2007 - Anderson
  '389. Campo utilizado para informar o tipo do cliente, conforme solicitação do cliente Gurgel & Leite para exportação de dados
  'no sistema Sadigweb. Este campo estará disponível apenas para o cliente Gurgel & Leite.
  'AG -AGROPECUARIA
  'CC - CRIADOR DE CAES
  'CG - CRIADOR DE GATOS
  'CL - CLINICA VETERINARIA
  'CP - CLINICA VETERINARIA COM PETSHOP
  'CR - CRIADOR DE CAES E GATOS
  'PC - PETSHOP COM CLINICA VETERINARIA
  'PE -PETSHOP
  'VE -VETERINARIO
  '     Tabela     : Funcionários
  '     Finalidade : Informar o tipo do cliente utilizado no sistema da SadigWeb para ser exportado no arquivo de texto
  '     Solicitante: Gurgel & Leite Comércio de Produtos Veterinários Ltda (QS31734-765)
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "SadigWeb_Tipo") Then
    If Not gbCreateField("Cli_For", "SadigWeb_Tipo", dbText, 40) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '17/05/2007 - Anderson
  '
  '333. Campo utilizado para informar o tipo da operação, conforme solicitação do cliente Gurgel & Leite para exportação de dados
  'no sistema Sadigweb. Este campo estará disponível apenas para o cliente Gurgel & Leite.
  'VE -Venda
  'BO -BONIFICAÇÃO
  'OU -OUTROS
  '     Tabela     : Funcionários
  '     Finalidade : Informar o tipo da operação utilizado no sistema da SadigWeb para ser exportado no arquivo de texto
  '     Solicitante: Gurgel & Leite Comércio de Produtos Veterinários Ltda (QS31734-765)
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "SadigWeb_Tipo") Then
    If Not gbCreateField("Operações Saída", "SadigWeb_Tipo", dbText, 15) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '19/05/2007 - Anderson
  '
  '334. Campo utilizado para informar a obrigatoriedade do campo Lembrar Em na tela de Clientes/Fornecedores, guia contatos efetuados.
  '     Tabela     : Funcionários
  '     Finalidade : Obrigar a digitação da campo Lembrar Em na tela de contatos efetuados do cliente/fornecedor
  '     Solicitante: Supri Print
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "ContatosEfetuadosLembrarEm") Then
    If Not gbCreateField("Funcionários", "ContatosEfetuadosLembrarEm", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '31/07/2007 - Anderson
  '
  '335. Campo utilizado para Armazenar o preço de custo do produto no momento da venda.
  '     Tabela     : Saídas - Produtos
  '     Finalidade : Armazenar o preço de custo do produto no momento da venda para análises posteriores
  '     Solicitante: Candy Clean
  nPhase = nPhase + 1
  If Not gbGetField("Saídas - Produtos", "PrecoCusto") Then
    If Not gbCreateField("Saídas - Produtos", "PrecoCusto", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      'Atualiza o preço de custo em todas as vendas.
      db.Execute "UPDATE [Saídas - Produtos] INNER JOIN Preços ON [Saídas - Produtos].[Código sem Grade] = Preços.Produto SET [Saídas - Produtos].PrecoCusto = [Preços]![Preço] WHERE Preços.Tabela='CUSTO'", dbFailOnError
    End If
  End If
  
  '07/08/2007 - Anderson
  '
  '336. Inclusão do relatório de comissões por vendedor
  '     Tabela     : ZZZProgramas
  '     Finalidade : Inclusão de novo programa
  '     Solicitante: CandyClean
  nPhase = nPhase + 1
  If CheckSerialCaseMod("QS37957-281") Then
    If Not AddFileZZZProgramas18 Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '17/08/2007 - Anderson
  '
  '337. Alteração do tamanho do campo de cadastro de características do cliente/fornecedor
  '     Tabela     : TabCaractCliFor
  '     Finalidade : Alterar o tamanho do campo pois estava limitado a quantidade da caracteres
  '     Solicitante: Marcelo (Infopar)
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo("TabCaractCliFor", "DescCaract", dbText, 255) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""TabCaractCliFor"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '17/08/2007 - Anderson
  '
  '338. Alteração do tamanho do campo de características do cliente/fornecedor
  '     Tabela     : CliForCaract
  '     Finalidade : Alterar o tamanho do campo pois estava limitado a quantidade da caracteres
  '     Solicitante: Marcelo (Infopar)
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo("CliForCaract", "ValCaract", dbText, 255) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""CliForCaract"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '29/08/2007 - Anderson
  '
  '339. Implementação de recurso para automatizar o cálculo do custo do produto.
  '     Tabela     : Operações Entrada
  '     Finalidade : Informar se a operação realiza cálculo automático do custo do produto de acordo com as configurações na guia Cálculo no cadastro de produtos
  '     Solicitante: Candy Clean
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "PrecoCustoCalculado") Then
    If Not gbCreateField("Operações Entrada", "PrecoCustoCalculado", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '24/09/2007 - Anderson
  '
  '340. Implementação do campo Código de Barras para utilização da impressão de carnês
  '     Tabela     : Contas a Receber
  '     Finalidade : Automatizar o procedimento de emissão de carnês e de pagamento de contas.
  '     Solicitante: Naativa (QS73159-473)
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CarneCodigoBarras") Then
    If Not gbCreateField("Contas a Receber", "CarneCodigoBarras", dbText, 40) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      If Not g_blnGetIndex("Contas a Receber", "CarneCodigoBarras") Then
        If Not m_blnCreateIndexCarneCodigoBarras() Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Contas a Receber"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
        End If
      End If
    End If
  End If
  
  '19/10/2007 - Anderson
  '
  '341. Informar o lucro mínimo permitido por classe de produtos, evitando assim que o vendedor utilize descontos fora do padrão
  '     Tabela     : Classes
  '     Solicitante: Agrotama
  nPhase = nPhase + 1
  If Not gbGetField("Classes", "LucroMinimoPermitido") Then
    If Not gbCreateField("Classes", "LucroMinimoPermitido", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Classes"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE Classes SET LucroMinimoPermitido = 0 ", dbFailOnError
    End If
  End If
  
  '19/10/2007 - Anderson
  '
  '342. Informar o lucro mínimo permitido por classe de produtos, evitando assim que o vendedor utilize descontos fora do padrão
  '     Tabela     : Funcionários
  '     Solicitante: Agrotama
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "LucroMinimoPermitido") Then
    If Not gbCreateField("Funcionários", "LucroMinimoPermitido", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '30/10/2007 - Anderson
  '
  '343. Campo utilizado para informar a quantidade ocupada pelo produto no estoque
  '     Tabela     : Produtos
  '     Solicitante: Kings Cross
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "EspacoFisicoTotal") Then
    If Not gbCreateField("Produtos", "EspacoFisicoTotal", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '30/10/2007 - Celso
  '
  '344. Campo utilizado para imprimir etiqueta pequena de produto
  '     Tabela     : Etiquetas - Tempo
  '     Solicitante: Jefferson

  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "Lote") Then
    If Not gbCreateField("Etiquetas - Tempo", "Lote", dbText, 15) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '30/10/2007 - Celso
  '
  '345. Campo utilizado para imprimir etiqueta pequena de produto
  '     Tabela     : Etiquetas - Tempo
  '     Solicitante: Jefferson
  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "DataValidade") Then
    If Not gbCreateField("Etiquetas - Tempo", "DataValidade", dbDate) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '30/10/2007 - Anderson
  '
  '346. Inclusão do relatório de comissões por vendedor
  '     Tabela     : ZZZProgramas
  '     Finalidade : Inclusão de novo programa
  '     Solicitante: Kings Cross
  nPhase = nPhase + 1
  If CheckSerialCaseMod("QS38393-282", "QS38714-658") Then
    If Not AddFileZZZProgramas19 Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '05/11/2007 - Anderson
  '
  '347. Parametros de operações de saídas para somar produtos no total da nota
  '     Tabela     : Operações Saída
  '     Finalidade : Configurar se os produtos devem somar o total da nota
  '     Solicitante: Cristiano Pavinatto
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "SomarProdutosTotalNota") Then
    If Not gbCreateField("Operações Saída", "SomarProdutosTotalNota", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE [Operações Saída] SET SomarProdutosTotalNota = -1 ", dbFailOnError
    End If
  End If
  
  '04/12/2007 - Celso
  '
  '348. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : ExigeSenhaGerVndContaAtraso
  '     Finalidade : Exigir senha do gerente no caso de venda para clientes com contas em atraso
  '     Solicitante: Valdeci - Vaplak
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ExigeSenhaGerVndContaAtraso") Then
    If Not gbCreateField("Parâmetros Filial", "ExigeSenhaGerVndContaAtraso", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros da Empresa/Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "NumCasasDecimais") Then
    If Not gbCreateField("Parâmetros Filial", "NumCasasDecimais", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '04/12/2007 - Anderson
  '
  '349. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : ImprimeNotaMovEfetivada
  '     Finalidade : Somente imprimir nota fiscal para movimentações efetivadas

  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ImprimeNotaMovEfetivada") Then
    If Not gbCreateField("Parâmetros Filial", "ImprimeNotaMovEfetivada", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros da Empresa/Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '350. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : NaoPermiteDuplicarCNPJ
  '     Finalidade : Não permitir duplicidade de CNPJ e CPF em cadastro de Clientes
  '     Solicitante: SMQ
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "NaoPermiteDuplicarCNPJ") Then
    If Not gbCreateField("Parâmetros Filial", "NaoPermiteDuplicarCNPJ", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros da Empresa/Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '18/12/2007 - Anderson
  '351. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Entradas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "NSU") Then
    If Not gbCreateField("Entradas", "NSU", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '18/12/2007 - Anderson
  '352. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Entradas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "NSU_Data") Then
    If Not gbCreateField("Entradas", "NSU_Data", dbDate) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '18/12/2007 - Anderson
  '353. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Entradas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "NSU_Hora") Then
    If Not gbCreateField("Entradas", "NSU_Hora", dbDate) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '18/12/2007 - Anderson
  '354. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Saídas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NSU") Then
    If Not gbCreateField("Saídas", "NSU", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '18/12/2007 - Anderson
  '355. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Entradas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NSU_Data") Then
    If Not gbCreateField("Saídas", "NSU_Data", dbDate) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '18/12/2007 - Anderson
  '356. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Entradas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NSU_Hora") Then
    If Not gbCreateField("Saídas", "NSU_Hora", dbDate) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '18/12/2007 - Anderson
  '357. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Saídas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "NSU") Then
    If Not gbCreateField("Parâmetros Filial", "NSU", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE [Parâmetros Filial] SET NSU = 0 ", dbFailOnError
    End If
  End If
  
  '19/12/2007 - Anderson
  '358. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Saídas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetTable("NSU") Then
    If Not gbCreateTableNSU Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""NSU"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
    End If
  End If
  
  '22/02/2008 - Celso
  '359. Implementação do NSU (Número de Série Única) para receita estadual de Santa Catarina
  '     Tabela     : Saídas
  '     Finalidade : Incluir campo para armazenamento do NSU
  '     Solicitante: Infopar
  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "PrecoPrazo") Then
    If Not gbCreateField("Etiquetas - Tempo", "PrecoPrazo", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '29/04/2008 - mpdea
  '
  '360. Parametros de operações de saídas para exibir tela para preenchimento
  '     do número de documento (CPF ou CNPJ)
  '     Tabela     : Operações Saída
  '     Finalidade : Atender solicitação do programa Nota Fiscal Paulista
  '     Solicitante: Nota Fiscal Paulista
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "ExibirTelaNumeroDocumento") Then
    If Not gbCreateField("Operações Saída", "ExibirTelaNumeroDocumento", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE [Operações Saída] SET ExibirTelaNumeroDocumento = 0 ", dbFailOnError
    End If
  End If
  
  '29/04/2008 - mpdea
  '
  '361. Número do documento (CPF ou CNPJ)
  '     Tabela     : Saídas
  '     Finalidade : Atender solicitação do programa Nota Fiscal Paulista
  '     Solicitante: Nota Fiscal Paulista
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NumeroDocumentoCliente") Then
    If Not gbCreateField("Saídas", "NumeroDocumentoCliente", dbText, 20) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/06/2008 - mpdea
  '362. Valor de isenção mensal no cálculo de impostos de serviços (PIS, COFINS e CSLL)
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ValorIsencaoPisCofinsCsll") Then
    If Not gbCreateField("Parâmetros Filial", "ValorIsencaoPisCofinsCsll", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      'Valor atual de isenção mensal (5000 ou superior não isento)
      db.Execute "UPDATE [Parâmetros Filial] SET ValorIsencaoPisCofinsCsll = 4999.99 ", dbFailOnError
    End If
  End If
  
  '25/09/2008 - mpdea
  '
  '363. Inclusão de permissões ausentes para relatórios de vendas
  '     Tabela     : ZZZProgramas
  '     Finalidade : Inclusão de permissões
  '     Solicitante: Patrício (Technomax)
  nPhase = nPhase + 1
  If Not AddFileZZZProgramas20 Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
    
  '11/11/2008 - mpdea
  '
  '364. Parâmetro de operações de saídas
  '     do número de documento (CPF ou CNPJ)
  '     Tabela     : Operações Saída
  '     Finalidade : Somar icms retido ao total da nota
  '     Solicitante: Patricio (Technomax)
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "SomaIcmsRetidoTotalNota") Then
    If Not gbCreateField("Operações Saída", "SomaIcmsRetidoTotalNota", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE [Operações Saída] SET SomaIcmsRetidoTotalNota = 0 ", dbFailOnError
    End If
  End If
  
  '30/01/2009 - mpdea
  '
  '365. Tabela de configuração para envio de email
  '     Tabela     : Email
  nPhase = nPhase + 1
  If Not gbGetTable("Email") Then
    If Not gbCreateTableEmail Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Email"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
    
  '18/06/2009 - mpdea
  '
  '366. Alíquota de ICMS para aproveitamento de crédito
  '     Tabela     : Parâmetros Filial
  '     Finalidade : Exibir na nota fiscal a alíquota e seu valor sobre o total da movimentação
  '     Solicitante: Cristiano Pavinato (Ti-Brasil)
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "AliquotaAprovCreditoIcms") Then
    If Not gbCreateField("Parâmetros Filial", "AliquotaAprovCreditoIcms", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE [Parâmetros Filial] SET AliquotaAprovCreditoIcms = 0 ", dbFailOnError
    End If
  End If
  
   
  '17/08/2009 - mpdea
  '367. Alterações na base de dados para adequação a NFe (Nota Fiscal Eletrônica)
'  nPhase = nPhase + 1
'  Call AlteraDBNFe
  
  
  '17/11/2009 - mpdea
  '368. Adiciona permissão para acesso ao Quick Cockpit
  nPhase = nPhase + 1
  Call AddUserPermission("QUICK COCKPIT", "Quick Cockpit - Visões Estratégicas e Gerenciais", 183, ID_ITEM_INICIO_COCKPIT)
  
  '09/12/2009 - Andrea
  '369. Criação da tabela Movimento - Cartoes
  '     Tabela     : Movimento - Cartoes
  '     Finalidade : Armazenar os dados do recebimento feito em cartões entre a tela de recebimento e a efetivação da movimentacao (para gravar no contas a receber).
  '     Solicitante: Marcelo
'  nPhase = nPhase + 1
'  If Not gbGetTable("Movimento - Cartoes") Then
'    If Not gbCreateTableMovimentoCartoes Then
'        ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Movimento - Cartoes"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'    End If
'  End If
     
  '10/12/2009 - Andrea
  '370. Criação do Indice Ordem na tabela Movimento - Cartoes
  nPhase = nPhase + 1
  Call gbCreateIndexFieldMovimentoCartoes
    
  '08/01/2010 - Andrea
  '371. Criação do campo FornecedorCreditado na tabela de Contas a Receber
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "FornecedorCreditado") Then
    If Not gbCreateField("Contas a Receber", "FornecedorCreditado", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '08/01/2010 - Andrea
  '372. Criação do campo SequenciaEntrada na tabela de Contas a Receber
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "SequenciaEntrada") Then
    If Not gbCreateField("Contas a Receber", "SequenciaEntrada", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '08/01/2010 - Andrea
  '373. Criação do campo Troco tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "Troco") Then
    If Not gbCreateField("Entradas", "Troco", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '374. Criação do campo NumeroDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "NumeroDI") Then
    If Not gbCreateField("Entradas", "NumeroDI", dbText, 10) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '375. Criação do campo CodigoExportador tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "CodigoExportador") Then
    If Not gbCreateField("Entradas", "CodigoExportador", dbText, 60) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '376. Criação do campo DataDeRegistroDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "DataDeRegistroDI") Then
    If Not gbCreateField("Entradas", "DataDeRegistroDI", dbDate) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '377. Criação do campo UFDesembaracoDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "UFDesembaracoDI") Then
    If Not gbCreateField("Entradas", "UFDesembaracoDI", dbText, 2) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '378. Criação do campo LocalDesembaracoDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "LocalDesembaracoDI") Then
    If Not gbCreateField("Entradas", "LocalDesembaracoDI", dbText, 60) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '379. Criação do campo DataDesembaracoDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "DataDesembaracoDI") Then
    If Not gbCreateField("Entradas", "DataDesembaracoDI", dbDate) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '380. Criação do campo NumeroAdicaoDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "NumeroAdicaoDI") Then
    If Not gbCreateField("Entradas", "NumeroAdicaoDI", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '381. Criação do campo NumeroSeqItemAdicaoDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "NumeroSeqItemAdicaoDI") Then
    If Not gbCreateField("Entradas", "NumeroSeqItemAdicaoDI", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '382. Criação do campo CodigoFabricanteAdicaoDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "CodigoFabricanteAdicaoDI") Then
    If Not gbCreateField("Entradas", "CodigoFabricanteAdicaoDI", dbText, 60) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/01/2010 - Andrea
  '383. Criação do campo DescontoAdicaoDI tabela Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "DescontoAdicaoDI") Then
    If Not gbCreateField("Entradas", "DescontoAdicaoDI", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '13/05/2010 - Andrea
  '384. Alteração tipo do campo Classificação Fiscal na tabela Produtos
  nPhase = nPhase + 1
  If gbAlteraClassificacaoFiscalProduto("Produtos") = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo Classificação Fiscal na tabela ""Produtos"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '13/05/2010 - Andrea
  '385. Alteração tipo do campo Classificação Fiscal na tabela Classificação Fiscal
  nPhase = nPhase + 1
  If gbAlteraClassificacaoFiscal("Classificação Fiscal") = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Classificação Fiscal"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '13/09/2012 - mpdea
  '
  '386. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : VRUtilizarTicketModoRelatorio
  '     Finalidade : Ticket em formato de relatório (devido a incompatibilidade com impressoras USB do objeto Printer)
  '     Solicitante: Amarelinha
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VRUtilizarTicketModoRelatorio") Then
    If Not gbCreateField("Parâmetros Filial", "VRUtilizarTicketModoRelatorio", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros da Empresa/Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '13/03/2013 - Alexandre Afornali
  '387. Criação do campo FiltrarProdutosInativos tabela Produtos
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "FiltrarProdutosInativos") Then
    If Not gbCreateField("Parâmetros Filial", "FiltrarProdutosInativos", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '17/05/2013 - Alexandre Afornali
  '388. Criação do campo TrabalharComComanda tabela Produtos para case DiskEmbalagens
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "TrabalharComComanda") Then
    If Not gbCreateField("Parâmetros Filial", "TrabalharComComanda", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '-----------------------------------------------------------------------------------
  'RODAPÉ
  'Esta função deverá sempre estar por último para evitar o lock na tabela de
  'Parâmetros
  '-----------------------------------------------------------------------------------
  '26/05/2004 - Daniel & Marcelo
  '
  'XXX. Tratamento para os campos da tabela Parâmetros
  '     CSLL, COFINS, PIS, IRRF
  '     Caso o percentual esteja vazio... atualizaremos para
  '     igual a zero para não dar conflito na emissão de notas
  nPhase = nPhase + 1
  If Not UpdateRecordParametros() Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Atualização de registro na tabela ""Parâmetros"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  '-----------------------------------------------------------------------------------
  
  'Call AlteraDB3(nPhase)
  
  Call WriteOurDBVersion
  
  Call ws.CommitTrans
  
  Screen.MousePointer = vbDefault
  
  On Error GoTo 0
  Exit Function
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Manutenção na Base de Dados - Alterações Vitais na Base de Dados não foram possíveis."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gsMsg = gsMsg & vbCrLf & "Fase da Alteração: " & CStr(nPhase)
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  '12/05/2004 - mpdea
  'Em caso de erro não tratado desfaz transações pendentes
  On Error Resume Next
  Call ws.Rollback
  
  Set db = Nothing
  Set dbFoo = Nothing
  Set ws = Nothing
  End
End Function

Private Function AlteraDB2(ByRef nPhase As Integer)
  Dim intX As Integer
  
  '1. Tabela Sistema
  nPhase = 1
  If gbGetField("ZZZ", "DBVersion") = False Then
    If gbCreateField("ZZZ", "DBVersion", dbText, 10) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Sistema"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '2. Tabela de Produtos
  nPhase = nPhase + 1
  If gbGetField("Produtos", "QtdeCasasDecimais") = False Then
    If gbCreateField("Produtos", "QtdeCasasDecimais", dbInteger, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na Tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '3. Tabela ZZZ
  nPhase = nPhase + 1
  If gbGetField("ZZZ", "CGCCPF") = False Then
    If gbCreateField("ZZZ", "CGCCPF", dbText, 30) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Sistema"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '4. Tabela Acessos e ZZZProgramas
  nPhase = nPhase + 1
  If gbGetField("Acessos", "Numero") = False Then
    If gbCreateField("Acessos", "Numero", dbInteger, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Acessos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '5. Tabela Acessos e ZZZProgramas
  nPhase = nPhase + 1
  If gbGetField("ZZZProgramas", "ToolID") = False Then
    If gbCreateField("ZZZProgramas", "ToolID", dbLong, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Acessos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
    If Not gbLoadToolID() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Carga de Valores na tabela ""Acessos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '6. Arquivo de Indices na Tabela Acessos
  nPhase = nPhase + 1
  Call gbCreateIndexFieldCodigosAcesso
  '
  '7. Carga dos Numeros de Acessos via tabela ZZZProgramas
  nPhase = nPhase + 1
  If gbLoadCodigosAcesso() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Atualização de campo na tabela ""Acessos"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  '
  '8. Alteração na Tabela de Etiquetas
  nPhase = nPhase + 1
  If gbGetField("Etiquetas", "Preco2") = False Then
    If gbCreateField("Etiquetas", "Preco2", dbSingle, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Etiquetas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '9. Alteração na Tabela de Funcionarios
  nPhase = nPhase + 1
  If gbGetField("Funcionários", "ValorP") = False Then
    If gbCreateField("Funcionários", "ValorP", dbText, 30) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
    If Not gbLoadValorP() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Atualização de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '
  '10. Alteração na Tabela de Operações de Saída
  nPhase = nPhase + 1
  If gbGetField("Operações Saída", "InTelaObsTransp") = False Then
    If gbCreateField("Operações Saída", "InTelaObsTransp", dbBoolean, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '11. Criação da Tabela Reports para suporte a flag de Relatórios Zebrados
  '    Modificado para o novo suporte de cores v.6.0.35 - por mpdea
  nPhase = nPhase + 1
  If gbGetTable("Reports") = False Then
    If gbCreateTableReports() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Reports"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '
  '12. Alteração na Tabela de Cli_For para inclusão do Tipo de Frete
  nPhase = nPhase + 1
  If gbGetField("Cli_For", "CodTipoFrete") = False Then
    If gbCreateField("Cli_For", "CodTipoFrete", dbText, 1) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '13. Criação da Tabela CliForCaract contendo os valores das Caracteristicas Diversas para o Cliente
  nPhase = nPhase + 1
  If gbGetTable("CliForCaract") = False Then
    If gbCreateTableCliForCaract() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""CliForCaract"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  Else
    If gbAlterTableCliForCaract() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Alteração da tabela ""CliForCaract"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '14. Criação da Tabela CliForCaract contendo as Caracteristicas Diversas para cada tipo de Cliente
  nPhase = nPhase + 1
  If gbGetTable("TabCaractCliFor") = False Then
    If gbCreateTableTabCaractCliFor() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""TabCaractCliFor"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  '
  '15. Criação da Tabela CliForMaterial contendo bens numeraveis associados ao Cliente
  nPhase = nPhase + 1
  If gbGetTable("CliForNumeravel") = False Then
    If gbCreateTableCliForNumeravel() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""CliForNumeravel"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '16. Alteração da Tabela "Conta Cliente" para adicionar o campo "TabPrecos"
  nPhase = nPhase + 1
  If gbGetField("Conta Cliente", "TabPrecos") = False Then
    If gbCreateField("Conta Cliente", "TabPrecos", dbText, 15) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Conta Cliente"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '17. Alteração da Tabela "Parâmetros Filial" para adicionar
  '    os campos para controle do comprimento do form
  nPhase = nPhase + 1
 ' If gbGetField("Parâmetros Filial", "Cód Comprim 1") = False Then
    If gbCreateFieldComprim("Parâmetros Filial", "Cód Comprim 1", dbText, 3, "1") = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  'End If
  'If gbGetField("Parâmetros Filial", "Cód Comprim 2") = False Then
    If gbCreateFieldComprim("Parâmetros Filial", "Cód Comprim 2", dbText, 3, "2") = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  'End If
  'If gbGetField("Parâmetros Filial", "Cód Comprim 3") = False Then
    If gbCreateFieldComprim("Parâmetros Filial", "Cód Comprim 3", dbText, 3, "3") = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  'End If

  '18. Alteração da Tabela "Produtos"  'Play it again SAM
  '    Adição dos campos PESO LÍQUIDO e PESO BRUTO
  nPhase = nPhase + 1
  If gbGetField("Produtos", "PesoLiquido") = False Then
    If gbCreateField("Produtos", "PesoLiquido", dbSingle, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na Tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  nPhase = nPhase + 1
  If gbGetField("Produtos", "PesoBruto") = False Then
    If gbCreateField("Produtos", "PesoBruto", dbSingle, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na Tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '19. Alteração da Tabela "Operações Entrada"
  '    Alteração do campo "Código Fiscal", i.e. CFOP
  nPhase = nPhase + 1
  gbFirstCFOP = False
  If gbAlteraCodigoFiscal("Operações Entrada") = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Operações Entradas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  '
  '20. Alteração da Tabela "Operações Saída"
  '    Alteração do campo "Código Fiscal", i.e. CFOP
  nPhase = nPhase + 1
  If gbAlteraCodigoFiscal("Operações Saída") = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Operações Saídas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '21. Inclusão do campo "InGeradoViaConsig" na tabela de Entradas
  nPhase = nPhase + 1
  If gbGetField("Entradas - Produtos", "InGeradoViaConsig") = False Then
    If gbCreateField("Entradas - Produtos", "InGeradoViaConsig", dbBoolean, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '22. Inclusão do campo "InGeradoViaConsig" na tabela de Saídas
  nPhase = nPhase + 1
  If gbGetField("Saídas - Produtos", "InGeradoViaConsig") = False Then
    If gbCreateField("Saídas - Produtos", "InGeradoViaConsig", dbBoolean, 0) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
    
  '23. Inclusão na Tabela de Funcionarios
  '    Incluído para suporte de desconto por funcionário na v.6.0.40 - por mpdea
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "bPermiteDesconto") Then
    If Not gbCreateField("Funcionários", "bPermiteDesconto", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
    If Not gbCreateField("Funcionários", "nPercDesconto", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
    'Atualiza a permissão de desconto para verdadeira como padrão
    db.Execute "UPDATE Funcionários SET bPermiteDesconto = True, nPercDesconto = 0;", dbFailOnError
  End If
  
  '24. Inclusão na Tabela de Operações de Entrada
  '    Incluído para suporte de cálculo de IPI com ICMS na v.6.0.40 - por mpdea
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "Base ICM com IPI") Then
    If Not gbCreateField("Operações Entrada", "Base ICM com IPI", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações de Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '25. Inclusão na Tabela de Cartões
  '    Incluído para suporte do TEF na v.6.0.40 - por mpdea
  nPhase = nPhase + 1
  If Not gbGetField("Cartões", "TEF") Then
    If Not gbCreateField("Cartões", "TEF", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cartões"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '26. Inclusão na Tabela de Operações de Saída
  '    Incluído para calculo do IPI somente no total na v.6.0.42 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "IPI TOT") Then
    If Not gbCreateField("Operações Saída", "IPI TOT", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '27. Inclusão na Tabela de Operações de Entrada
  '    Incluído para calculo do IPI somente no total na v.6.0.42 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "IPI TOT") Then
    If Not gbCreateField("Operações Entrada", "IPI TOT", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
   
  
  '29 Inclusão na Tabela de Saídas - Produtos
  '   Incluído campo Situação Tributária para emissão do registro 60 na v.6.0.42 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Saídas - Produtos", "Situação Tributária") Then
    If Not gbCreateField("Saídas - Produtos", "Situação Tributária", dbText, 3) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  ' 30. Alteração da Tabela "Produtos"
  '     Criação do campo "Percentual ICM Entrada" e Copia do Icm de Saída para esse campo.
  nPhase = nPhase + 1
  If gbAlteraIcmEntra("Produtos") = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Produtos"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If

  ' 31. Alteração da Tabela "Produtos"
  '     Criação do campo "Percentual ICM ECF" e Copia do Icm de Saída para esse campo.
  nPhase = nPhase + 1
  If gbAlteraIcmSai("Produtos") = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Produtos"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
   
  
  '32. Alteração da Tabela "ZZZProgramas"
  '    Atualiza Id de alguns programas na tabela zzzProgramas V.6.0.42
  nPhase = nPhase + 1
  If gbAlteraZZZProgramas() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
   
  ' 33. Adiciona Item na ZzzProgramas"
  '     Adiciona Programa novo na zzzProgramas V.6.0.42
  
  nPhase = nPhase + 1
  If AddFileZZZProgramas() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
   
  '34 Inclusão na Tabela de Saídas - Produtos
  '   Incluído campo Unidade de Venda para emissão do registro 60 na v.6.0.42 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Saídas - Produtos", "Unidade Venda") Then
    If Not gbCreateField("Saídas - Produtos", "Unidade Venda", dbText, 5) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
   
  '35 Inclusão na Tabela de Parametros Empresa Filial
  '   Incluído campo Saida Descr Adicional na v.6.0.43 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Saida Descr Adicional") Then
    If Not gbCreateField("Parâmetros Filial", "Saida Descr Adicional", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
   
'  36 Inclusão na Tabela de Saídas - Produtos
'     Incluído campo Saida Descr Adicional na v.6.0.43 - por Leandro
    nPhase = nPhase + 1
    If Not gbGetField("Saídas - Produtos", "Descricao Adicional") Then
      If Not gbCreateFieldZeroLenght("Saídas - Produtos", "Descricao Adicional", dbText, 50) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
   
  '  37 Inclusão na Tabela de Resumo Clientes
  '     Incluído campo Saida Descr Adicional na v.6.0.43 - por Leandro
    nPhase = nPhase + 1
    If Not gbGetField("Resumo Clientes", "Descricao Adicional") Then
      If Not gbCreateFieldZeroLenght("Resumo Clientes", "Descricao Adicional", dbMemo) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Resumo Clientes"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
  '38. Inclusão na Tabela de Operações de Saída
  '    Incluído para calculo do ICM de Frete na v.6.0.43 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "Perc Icms Frete") Then
    If Not gbCreateField("Operações Saída", "Perc Icms Frete", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
   
  '39. Inclusão na Tabela de Operações de Saída
  '    Incluído para calculo do ICM de Frete na v.6.0.43 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "Calcula Icm Frete") Then
    If Not gbCreateField("Operações Saída", "Calcula Icm Frete", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
     
     
  '40. Inclusão na Tabela de Operações de Saída
  '    Incluído para calculo do ICM de Frete na v.6.0.43 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "Soma Frete") Then
    If Not gbCreateField("Operações Saída", "Soma Frete", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '41. Inclusão na Tabela de Parâmetros
  '    Incluído campo para configuração da alteração de preço na tela de saídas na v.6.0.44 - por Leandro
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Saida Altera Preco") Then
    If gbCreateField("Parâmetros Filial", "Saida Altera Preco", dbBoolean) Then
      '01/08/2002 - mpdea
      'modificado a atualização do novo campo somente em sua criação
      '
      '42. Alteração do campo Altera Sai Precos do Parametros Filial
      '
      nPhase = nPhase + 1
      If gbGravaTrueParamSaiPrecos() = False Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Parâmetros Filial"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '43. Alterações no Banco de dados para compatibilidade com Loja Virtual V.6.0.44 - por Leandro
   nPhase = nPhase + 1
   Call AlteraDBWeb
   
  
  '31/07/2002 - mpdea
  '44. Inclusão na Tabela de Parâmetros
  '    Incluído campo para utilização da Loja Virtual - v6.0.45
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "WorkWeb") Then
    If Not gbCreateField("Parâmetros Filial", "WorkWeb", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If


  '08/08/2002 - mpdea
  '45. Inclusão na Tabela de Saídas
  '    Incluído campo para informações do nr. do orçamento - v6.0.45
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "InfoNrOrcamento") Then
    If Not gbCreateField("Saídas", "InfoNrOrcamento", dbText, 255) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  

  '08/08/2002 - mpdea
  '46. Inclusão na Tabela de Parâmetros Filial
  '    Incluído campo controle do nr. do orçamento - v6.0.45
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "NrOrcamento") Then
    If Not gbCreateField("Parâmetros Filial", "NrOrcamento", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
     
     
  '15/08/2002 - mpdea
  '47. Inclusão na Tabela de Saídas
  '    Incluído campo Desconto no Sub Total - v6.0.45
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "DescontoSubTotal") Then
    If Not gbCreateField("Saídas", "DescontoSubTotal", dbCurrency) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
     
     
  '30/08/2002 - maikel
  '48. Inclusão de um campo para guardar a data de abertura do cadastro
  nPhase = nPhase + 1
  If gbGetField("Cli_For", "datAberturaCadastro") = False Then
    If gbCreateField("Cli_For", "datAberturaCadastro", dbDate) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '04/09/2002 - maikel
  '49. Campo para guardar o centro de custo descrito na tela de entradas
  nPhase = nPhase + 1
  If gbGetField("Entradas", "CentroCusto") = False Then
    If gbCreateField("Entradas", "CentroCusto", dbInteger) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Sistema"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/09/2002 - maikel
  '50. Criação de uma nova consulta para o combo nome na tela de venda rápida
  nPhase = nPhase + 1
  On Error Resume Next
  Dim sSql As String
  sSql = " SELECT Produtos.Nome, Produtos.Código, Produtos.[Código Ordenação] " & _
         " FROM Produtos " & _
         " WHERE (((Produtos.Código) <> ""0"") AND " & _
         " ((Produtos.[Desativado]) = False)) " & _
         " ORDER BY Produtos.Nome "
  db.CreateQueryDef "Con_Produto2", sSql
  On Error GoTo 0
  '------------------------------------------------------------------
   
  '04/09/2002 - maikel
  '51. Inclusão na Tabela de Parâmetros, de um campo que determina se na tela de venda rápida o usuário pode pesquisar no campo nome
    nPhase = nPhase + 1
    If Not gbGetField("Parâmetros Filial", "PesquisaCodigoENome_VR") Then   'PesquisaCodigoENome_VR
      If Not gbCreateField("Parâmetros Filial", "PesquisaCodigoENome_VR", dbBoolean) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
   
  '==============================================================
  ' Dev..: Maikel
  ' Data.: 04/09/2002 15:22
  ' 52.    Alterações no banco de dados para o módulo de verificação de pedidos
  '--------------------------------------------------------------
    nPhase = nPhase + 1
    If Not gbGetField("Operações Saída", "ControleEntregas") Then
      If Not gbCreateField("Operações Saída", "ControleEntregas", dbBoolean) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações de Saída"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
    If Not gbGetField("Operações Saída", "OpEntrega") Then
      If Not gbCreateField("Operações Saída", "OpEntrega", dbInteger) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações de Saída"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
    If Not gbGetField("Saídas - Produtos", "QtdeEntregue") Then
      If Not gbCreateField("Saídas - Produtos", "QtdeEntregue", dbDouble) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
    If Not gbGetField("Saídas", "SequênciaPai") Then
      If Not gbCreateField("Saídas", "SequênciaPai", dbLong) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  '==============================================================
   
  '53.
    nPhase = nPhase + 1
    If Not gbGetField("Funcionários", "VRVisualizarEstoque") Then
      If Not gbCreateField("Funcionários", "VRVisualizarEstoque", dbBoolean) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
      db.Execute " UPDATE Funcionários SET VRVisualizarEstoque = TRUE "
      
    End If
  
  '54.
    nPhase = nPhase + 1
    If Not gbGetField("Funcionários", "VRVisualizarPreco") Then
      If Not gbCreateField("Funcionários", "VRVisualizarPreco", dbBoolean) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
      db.Execute " UPDATE Funcionários SET VRVisualizarPreco = TRUE "
    End If
  
  '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
  
  '17/09/2002 - mpdea
  '55. Inclusão na Tabela de Parâmetros
  '    Incluído campo para ativar Traffic Light - v6.0.45
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "WorkTrafficLight") Then
    If Not gbCreateField("Parâmetros Filial", "WorkTrafficLight", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/09/2002 - mpdea
  '56. Alteração na Tabela Produtos
  '    Alterado o tamanho do campo Nome de 50 para 80 - v6.0.45
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Produtos", "Nome", dbText, 80, "Nome") Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Produtos"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '18/09/2002 - mpdea
  '57. Alteração nas Tabelas Pesquisa 1, 2 e 3
  '    Alterado o tamanho do campo Nome de 30 para 80 - v6.0.45
  nPhase = nPhase + 1
  For intX = 1 To 3
    If Not gbAlteraTamanhoCampo2("Pesquisa " & intX, "Nome", dbText, 80, "Nome") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Pesquisa " & _
        intX & """ - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  Next intX
  
  
  '19/09/2002 - mpdea
  '58. Alteração na Tabela Produtos
  '    Alterado o tamanho do campo Situação Tributária para 4 - v6.0.45
  '    Alteração necessária devido a BUG em alteração anterior, em alguns casos
  '    estava com tamanho 2 ou 3 (padronizado 4, mas utilizado 3)
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Produtos", "Situação Tributária", dbText, 4) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Produtos"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '07/10/2002 - mpdea
  '59. Inclusão na Tabela de Parâmetros da Filial
  '    Incluído campo para verificação de estoque em Saídas - v6.0.45
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VerificaEstoqueSaidas") Then
    If Not gbCreateField("Parâmetros Filial", "VerificaEstoqueSaidas", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '29/10/2002 - mpdea
  '60. Alteração na Tabela Caixa
  '    Alterado o tamanho do campo Descrição de 30 para 60
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Caixa", "Descrição", dbText, 60) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Caixa"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '13/11/2002 - mpdea
  '61. Inclusão na Tabela de Parâmetros da Filial
  '    Incluído campo código da operação de saída para transformação do
  '    Orçamento em Venda - v6.45.7
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "OpSaidaOrcVenda") Then
    If Not gbCreateField("Parâmetros Filial", "OpSaidaOrcVenda", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/11/2002 - mpdea
  '62. Inclusão na Tabela de Saídas
  '    Incluído campo flag para impedir que uma movimentação possa ser
  '    gravada novamente - v6.45.7
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Locked") Then
    If Not gbCreateField("Saídas", "Locked", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/11/2002 - mpdea
  '63. Alteração na Tabela Produtos
  '    Alterado o tamanho do campo Nome para nota de 50 para 80 - v6.45.7
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Produtos", "Nome Nota", dbText, 80) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Produtos"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '08/04/2003 - mpdea
  '64. Inclusão na Tabela de Saídas
  '    Incluído campo data da emissão da nota fiscal - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "DataEmissaoNota") Then
    If Not gbCreateField("Saídas", "DataEmissaoNota", dbDate) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '08/04/2003 - mpdea
  '65. Inclusão na Tabela de Conta a Receber
  '    Incluído campo Nosso Número (boleto) - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CNAB_NossoNumero") Then
    If Not gbCreateField("Contas a Receber", "CNAB_NossoNumero", dbText, 20) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If


  '08/04/2003 - mpdea
  '66. Inclusão na Tabela de Conta a Receber
  '    Incluído campo de código de instrução do arquivo de retorno - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CNAB_CodMovRet") Then
    If Not gbCreateField("Contas a Receber", "CNAB_CodMovRet", dbByte) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '09/04/2003 - mpdea
  '67. Inclusão na Tabela de Conta a Receber
  '    Incluído campo descrição do campo CNAB_CodMovRet - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CNAB_DescrMovRet") Then
    If Not gbCreateField("Contas a Receber", "CNAB_DescrMovRet", dbText, 255) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '30/04/2003 - mpdea
  '68. Inclusão na Tabela de Parâmetro Filial
  '    Incluído campo Desconto no Sub Total rateado para Venda Rápida e Saídas - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "DescSubTotalRateado") Then
    If Not gbCreateField("Parâmetros Filial", "DescSubTotalRateado", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '30/04/2003 - maikel
  '69. Criado campo que diz a forma de ordenação da combo de código na tela de venda rápida
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VROrdenacaoCombo") Then
    If Not gbCreateField("Parâmetros Filial", "VROrdenacaoCombo", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '05/05/2003 - maikel
  '70. Inclusão na Tabela de Conta a Receber
  '    Incluído campo descrição do campo CNAB_CodMovRet - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CNAB_CodIdComplementar") Then
    If Not gbCreateField("Contas a Receber", "CNAB_CodIdComplementar", dbByte) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '02/06/2003 - maikel
  '71. Inclusão na Tabela de Cliente e Fornecedores
  '    Incluído campo DiaBaseConsignacao que guarda o dia base para acerto de consignações - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "DiaBaseConsignacao") Then
    If Not gbCreateField("Cli_For", "DiaBaseConsignacao", dbByte) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '02/06/2003 - maikel
  '72. Inclusão na Tabela de Usuários/ Funcionários
  '    Incluído campo MargemLimiteCredito que guarda a margem excedente ao limite de crédito - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "MargemLimiteCredito") Then
    If Not gbCreateField("Funcionários", "MargemLimiteCredito", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/06/2003 - maikel
  '73. Inclusão na Tabela de Clientes/ Fornecedores
  '    Data do próximo acerto da consignação - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "DataProxAcertoConsignacao") Then
    If Not gbCreateField("Cli_For", "DataProxAcertoConsignacao", dbDate) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Clientes/ Fornecedores"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/06/2003 - maikel
  '74. Inclusão na Tabela de Clientes/ Fornecedores
  '    Número da última sequência de consignação - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "UltimaConsignacao") Then
    If Not gbCreateField("Cli_For", "UltimaConsignacao", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Clientes/ Fornecedores"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '23/06/2003 - maikel
  '75. Criado campo que diz a forma de ordenação da combo de código na tela de venda rápida
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "UltimaConsignacao") Then
    If Not gbCreateField("Parâmetros Filial", "UltimaConsignacao", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '10/07/2003 - Maikel
  '76. Criado campo guarda o número da consignação mestre na tabela saídas
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "ConsignacaoMestre") Then
    If Not gbCreateField("Saídas", "ConsignacaoMestre", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '10/07/2003 - Maikel
  '77. Criado campo guarda o número da consignação mestre na tabela entradas
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "ConsignacaoMestre") Then
    If Not gbCreateField("Entradas", "ConsignacaoMestre", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/06/2003 - maikel
  '78. Criado campo que diz a operação de entrada para consignação
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Consignacao_OpEntrada") Then
    If Not gbCreateField("Parâmetros Filial", "Consignacao_OpEntrada", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/06/2003 - maikel
  '79. Criado campo que diz a operação de saída para consignação
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Consignacao_OpSaida") Then
    If Not gbCreateField("Parâmetros Filial", "Consignacao_OpSaida", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/06/2003 - maikel
  '80. Criado campo que diz o caixa a ser usado na consignação
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Consignacao_Caixa") Then
    If Not gbCreateField("Parâmetros Filial", "Consignacao_Caixa", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/06/2003 - maikel
  '81. Criado campo que diz a tabela de preços a ser usado na consignação
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Consignacao_TabelaPrecos") Then
    If Not gbCreateField("Parâmetros Filial", "Consignacao_TabelaPrecos", dbText, 15) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '14/06/2003 - maikel
  '82. Inclusão na Tabela de Clientes/ Fornecedores
  '    campo que diz se a ultima consignacao está fechada - v6.45.8
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "ConsignacaoFechada") Then
    If gbCreateField("Cli_For", "ConsignacaoFechada", dbBoolean) Then
      '15/06/2003 - maikel
      db.Execute "UPDATE Cli_For SET ConsignacaoFechada = TRUE WHERE ISNULL(UltimaConsignacao) = TRUE"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Clientes/ Fornecedores"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '11/06/2003 - maikel
  '83. Criado campo que diz a operação de fechamento para consignação
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Consignacao_OpFechamento") Then
    If Not gbCreateField("Parâmetros Filial", "Consignacao_OpFechamento", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '07/08/2003 - mpdea
  '84. Parâmetro para não permitir executar mais de uma vez
  '    o sistema na estação
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "CheckInstance") Then
    If Not gbCreateField("Parâmetros Filial", "CheckInstance", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  Call CreateFieldsOBS(nPhase)
  
  '15/08/2003 - maikel
  '102. Preço do produto
  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "Preco") Then
    If Not gbCreateField("Etiquetas - Tempo", "Preco", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '19/08/2003 - mpdea
  '103. Modificado nome do campo VerificaEstoqueSaidas para
  '     Venda Sem Estoque Saidas em Parâmetros Filial
  nPhase = nPhase + 1
  If gbGetField("Parâmetros Filial", "VerificaEstoqueSaidas") And Not gbGetField("Parâmetros Filial", "Venda Sem Estoque Saidas") Then
    db.TableDefs("Parâmetros Filial").Fields("VerificaEstoqueSaidas").Name = "Venda Sem Estoque Saidas"
  End If
  
  
  '29/08/2003 - mpdea
  '104. Inclusão na Tabela de Usuários/ Funcionários
  '     Incluído campo PermiteAcharVenda para controle de permissão da função
  '     Achar Venda
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "PermiteAcharVenda") Then
    If gbCreateField("Funcionários", "PermiteAcharVenda", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET PermiteAcharVenda = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '03/09/2003 - Maikel
  '105.         Adicionado campo que dá permissão ao operador de caixa de visualizar o limite de crédito do cliente.
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "VR_PermiteVisualizarLimiteCredito") Then
    If gbCreateField("Funcionários", "VR_PermiteVisualizarLimiteCredito", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET VR_PermiteVisualizarLimiteCredito = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '24/09/2003 - mpdea
  '106. Incluído índice na tabela de saídas
  nPhase = nPhase + 1
  If Not g_blnGetIndex("Saídas", "VrAchaVenda") Then
    If Not m_blnCreateIndexVrAchaVenda() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '22/10/2003 - Maikel
  '107. Adicionado campo que guarda a tabela de preços padrão
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "TabelaPrecoPadrao") Then
    If Not gbCreateField("Cli_For", "TabelaPrecoPadrao", dbText, 15) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '22/10/2003 - Maikel
  '108. Adicionado campo que guarda o percentual a diminuar da comissão, caso seja aplicado algum desconto.
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "DiasBloqueioVenda") Then
    If Not gbCreateField("Parâmetros Filial", "DiasBloqueioVenda", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Tabela de Preços"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '22/10/2003 - Maikel
  '108. Adicionado campo que guarda o percentual a diminuar da comissão, caso seja aplicado algum desconto.
  nPhase = nPhase + 1
  If Not gbGetField("Tabela de Preços", "PercentualComissaoDesconto") Then
    If Not gbCreateField("Tabela de Preços", "PercentualComissaoDesconto", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Tabela de Preços"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '29/10/2003 - Maikel
  '109. Adicionado campo que diz se o orçamento foi ou não aprovado
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "OrcamentoAprovado") Then
    If Not gbCreateField("Saídas", "OrcamentoAprovado", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '29/10/2003 - Maikel
  '110. Adicionado campo que guarda a observação sobre a liberação ou bloqueio do orçamento
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "ComentariosSobreOrcamento") Then
    If Not gbCreateField("Saídas", "ComentariosSobreOrcamento", dbMemo) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '30/10/2003 - Maikel
  '111. Adicionado campo que diz se o orçamento deve ser aprovado para que seja transformado em venda
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "ExigeAprovacaoOrcamento") Then
    If Not gbCreateField("Operações Saída", "ExigeAprovacaoOrcamento", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/12/2003 - Daniel
  'Case: F. Linhares
  '
  '112. Inclusão na Tabela de Funcionários
  '     Incluído campo SenhaConfirmarCRDiff para controle de baixas
  '     com datas ou valores diferentes dos previstos
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "SenhaConfirmarCRDiff") Then
    If gbCreateField("Funcionários", "SenhaConfirmarCRDiff", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET SenhaConfirmarCRDiff = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '29/12/2003 - mpdea
  '113. Inclusão de campo
  '     Tabela    : Parâmetros da Filial
  '     Campo     : VR_GravarExigeSenhaVend
  '     Descrição : Flag para a exigência da senha do vendedor de caixa
  '                 sempre que gravar venda
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VR_GravarExigeSenhaVend") Then
    If Not gbCreateField("Parâmetros Filial", "VR_GravarExigeSenhaVend", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '29/12/2003 - mpdea
  '114. Inclusão de campo
  '     Tabela    : Produtos
  '     Campo     : DontAllowDesc
  '     Descrição : Flag proibindo desconto no produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "DontAllowDesc") Then
    If Not gbCreateField("Produtos", "DontAllowDesc", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '05/01/2004 - Daniel
  '115. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Valor Recebido
  '     Descrição : Finalidade na impressão de ticket's e recibo
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Valor Recebido") Then
    If Not gbCreateField("Saídas", "Valor Recebido", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '05/01/2004 - Daniel
  '116. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Troco
  '     Descrição : Finalidade na impressão de ticket's e recibo
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Troco") Then
    If Not gbCreateField("Saídas", "Troco", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
    
    
  '20/01/2004 - Daniel
  '117. Criação da Tabela Contrato para atender inicialmente à STC
  'de Caxias do Sul - RS
  nPhase = nPhase + 1
  If gbGetTable("Contrato") = False Then
    If gbCreateTableContrato() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Contrato"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '30/01/2004 - Daniel
  'A seguir temos a Criação de 04 campos para a tabela [Parâmetros Filial]
  'Impostos sobre Serviços: CSLL, COFINS, PIS, IRRF
  '
  '118. Inclusão de campo
  '     Tabela    : [Parâmetros Filial]
  '     Campo     : CSLL
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "CSLL") Then
    If Not gbCreateField("Parâmetros Filial", "CSLL", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '119. Inclusão de campo
  '     Tabela    : [Parâmetros Filial]
  '     Campo     : COFINS
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "COFINS") Then
    If Not gbCreateField("Parâmetros Filial", "COFINS", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '120. Inclusão de campo
  '     Tabela    : [Parâmetros Filial]
  '     Campo     : PIS
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "PIS") Then
    If Not gbCreateField("Parâmetros Filial", "PIS", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '121. Inclusão de campo
  '     Tabela    : [Parâmetros Filial]
  '     Campo     : IRRF
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "IRRF") Then
    If Not gbCreateField("Parâmetros Filial", "IRRF", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '02/02/2004 - Daniel
  'A seguir temos a Criação de 04 campos para a tabela Saídas
  'com a Finalidade de armazenar os percentuais (históricos) dos
  'Impostos sobre Serviços: CSLL, COFINS, PIS, IRRF
  '
  '122. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Percentual CSLL
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Percentual CSLL") Then
    If Not gbCreateField("Saídas", "Percentual CSLL", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '123. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Percentual COFINS
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Percentual COFINS") Then
    If Not gbCreateField("Saídas", "Percentual COFINS", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '124. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Percentual PIS
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Percentual PIS") Then
    If Not gbCreateField("Saídas", "Percentual PIS", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '125. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Percentual IRRF
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Percentual IRRF") Then
    If Not gbCreateField("Saídas", "Percentual IRRF", dbSingle) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '19/02/2004 - Daniel
  '
  '126. Inclusão de campo
  '     Tabela    : Operações Saída
  '     Campo     : Validade
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "Validade") Then
    If Not gbCreateField("Operações Saída", "Validade", dbBoolean, 0) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '19/02/2004 - Daniel
  '
  '127. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Data Validade
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Data Validade") = False Then
    If Not gbCreateField("Saídas", "Data Validade", dbDate) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '27/02/2004 - Daniel
  '
  '128. Inclusão de campo
  '     Tabela    : Operações Entrada
  '     Campo     : Estorno
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "Estorno") Then
    If Not gbCreateField("Operações Entrada", "Estorno", dbBoolean, 0) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/03/2004 - mpdea
  '
  '129. Incluído índice na tabela de saídas
  nPhase = nPhase + 1
  If Not g_blnGetIndex("Saídas", "DataMov") Then
    If Not m_blnCreateIndexSaidasDataMov() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '11/03/2004 - Daniel
  'Case: F. Linhares
  '
  '130. Inclusão na Tabela de Funcionários
  '     Adicionado o campo ImprimirTicket para controle de
  '     impressões no Manutenções de Contas a Receber
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "ImprimirTicket") Then
    If gbCreateField("Funcionários", "ImprimirTicket", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET ImprimirTicket = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '31/03/2004 - Daniel
  'Case: STC
  '
  '131. Criação da Tabela Radio para atender inicialmente à STC
  'de Caxias do Sul - RS
  nPhase = nPhase + 1
  If gbGetTable("Radio") = False Then
    If gbCreateTableRadio() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Radio"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '31/03/2004 - Daniel
  'Case: STC
  '
  '132. Criação da Tabela TipoComercial para atender inicialmente à STC
  'de Caxias do Sul - RS
  nPhase = nPhase + 1
  If gbGetTable("TipoComercial") = False Then
    If gbCreateTableTipoComercial() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""TipoComercial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '06/04/2004 - Daniel
  '133. Criação da Tabela Programacao para atender inicialmente à STC
  'de Caxias do Sul - RS
  'Nota: Esta Tabela é Filha da Tabela Contrato [Um Contrato para n Programacoes]
  nPhase = nPhase + 1
  If gbGetTable("Programacao") = False Then
    If gbCreateTableProgramacao() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Programacao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '07/04/2004 - Daniel
  'Case: STC de Caxias do Sul - RS
  '
  '134. Inclusão na Tabela de Serviços
  '     Adicionado o campo Publicidade
  nPhase = nPhase + 1
  If Not gbGetField("Serviços", "Publicidade") Then
    If gbCreateField("Serviços", "Publicidade", dbBoolean) Then
      db.Execute "UPDATE Serviços SET Publicidade = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Serviços"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '12/04/2004 - Daniel
  '
  '135. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : Num Autorizacao
  '     Case      : STC de Caxias do Sul
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Num Autorizacao") Then
    If Not gbCreateField("Saídas", "Num Autorizacao", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '12/04/2004 - Daniel
  '
  '136. Inclusão de campo
  '     Tabela    : Saídas
  '     Campo     : MesX
  '     Case      : STC de Caxias do Sul
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "MesX") Then
    If Not gbCreateField("Saídas", "MesX", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/04/2004 - Daniel
  '
  '137. Inclusão de campo
  '     Tabela    : Saídas
  '     Campos    : Total CSLL
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Total CSLL") Then
    If Not gbCreateField("Saídas", "Total CSLL", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/04/2004 - Daniel
  '
  '138. Inclusão de campo
  '     Tabela    : Saídas
  '     Campos    : Total COFINS
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Total COFINS") Then
    If Not gbCreateField("Saídas", "Total COFINS", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/04/2004 - Daniel
  '
  '139. Inclusão de campo
  '     Tabela    : Saídas
  '     Campos    : Total PIS
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Total PIS") Then
    If Not gbCreateField("Saídas", "Total PIS", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/04/2004 - Daniel
  '
  '140. Inclusão de campo
  '     Tabela    : Saídas
  '     Campos    : Total IRRF
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Total IRRF") Then
    If Not gbCreateField("Saídas", "Total IRRF", dbDouble) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '23/04/2004 - Daniel
  'Case: PSV
  '
  '141. Inclusão na Tabela de Saídas
  '     Incluído campo FaturaSourceReserva toda vez que é incluída
  '     uma Saída este campo será setado para False
  '     na tela de manutenção de reservas quando esta saída gerar
  '     uma saída para venda este campo será atualizado para True
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "FaturaSourceReserva") Then
    If gbCreateField("Saídas", "FaturaSourceReserva", dbBoolean) Then
      db.Execute "UPDATE Saídas SET FaturaSourceReserva = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '05/05/2004 - Daniel
  'Case: Embalavi
  '
  '143. Inclusão na Tabela de Cli_For
  '     Incluído campo IsentoIPI para controle se
  '     o Cli_For é isento deste Imposto, aproveitado
  '     para os demais clientes do Quick
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "IsentoIPI") Then
    If gbCreateField("Cli_For", "IsentoIPI", dbBoolean) Then
      db.Execute "UPDATE Cli_For SET IsentoIPI = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '06/05/2004 - Daniel
  'Case: Embalavi
  '
  '144. Inclusão na Tabela de Cli_For
  '     Incluído campo ObsIsentoIPI para impressão
  '     personalizada por Cli_For na NF aproveitado
  '     para os demais clientes do Quick
  nPhase = nPhase + 1
  If gbGetField("Cli_For", "ObsIsentoIPI") = False Then
    If gbCreateField("Cli_For", "ObsIsentoIPI", dbText, 100) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '11/05/2004 - Daniel
  '
  '145. Adiciona dois registros novos na ZzzProgramas
  '     São eles: Rel. de Estoque por filiais e Rel. Localização de Produtos
  nPhase = nPhase + 1
  If AddFileZZZProgramas2() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
    
  
  '12/05/2004 - Daniel
  '
  '146. Inclusão na Tabela de Saídas
  '     Incluído campo TotalMenosServ que sempre será:
  '     TotalMenosServ = Total em Serviços - (CSLL, COFINS, PIS, IRRF)
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "TotalMenosServ") Then
    If gbCreateField("Saídas", "TotalMenosServ", dbDouble) Then
      db.Execute "UPDATE Saídas SET TotalMenosServ = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '13/05/2004 - Daniel
  '
  '147. Inclusão na Tabela de [Operações Saída]
  '     Incluído campo ComissaoServicos, quando ele estiver Verdadeiro
  '     esta Op. de Saída suspenderá o Cálculo de alguns impostos sobre
  '     serviços: (CSLL, COFINS, PIS) mas o IRRF continuará calculando
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "ComissaoServicos") Then
    If gbCreateField("Operações Saída", "ComissaoServicos", dbBoolean) Then
      db.Execute "UPDATE [Operações Saída] SET ComissaoServicos = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  
  '14/05/2004 - Daniel
  'Case: Embalavi
  '
  '148. Criação da Tabela Diferimento para atender inicialmente à Embalavi
  nPhase = nPhase + 1
  If gbGetTable("Diferimento") = False Then
    If gbCreateTableDiferimento() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Diferimento"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '21/05/2004 - Daniel
  '
  'Case: Bic Amazônia
  '
  '149. Inclusão na Tabela de Saídas
  '     Incluído campo [Codigo Func Comprador]. A finalidade deste campo é de
  '     armazenar o código do funcionário
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Codigo Func Comprador") Then
    If gbCreateField("Saídas", "Codigo Func Comprador", dbInteger) Then
      db.Execute "UPDATE Saídas SET [Codigo Func Comprador] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '21/05/2004 - Daniel
  '
  'Case: Bic Amazônia
  '
  '150. Inclusão na Tabela de Saídas
  '     Incluído campo [Status Venda Func] quando gerarmos o arquivo de lay out
  '     de exportação este campo ficará true, pois a baixa será dada pelo RH da Bic
  '     este campo nada mais é do que um flag que indicará que o valor foi encaminhado
  '     para o RH
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Status Venda Func") Then
    If gbCreateField("Saídas", "Status Venda Func", dbBoolean) Then
      db.Execute "UPDATE Saídas SET [Status Venda Func] = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '14/07/2006 - mpdea
  'Modificado o valor padrão do novo campo para que não altere a situação padrão do cliente,
  'ficando assim, necessário a configuração conforme necessidade
  '
  '25/05/2004 - Daniel
  '
  '151. Inclusão na Tabela de Parâmetros Filial
  '     Incluído campo [VR_RecalcularPreço]
  '     Realiza ou não o recalculo dos preços (na grid de venda rápida) devido
  '     a possíveis modificações de desconto
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VR_RecalcularPreço") Then
    If gbCreateField("Parâmetros Filial", "VR_RecalcularPreço", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET VR_RecalcularPreço = False;", dbFailOnError
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/05/2004 - Daniel
  '
  'Case: Cia. do Aquário - RJ
  '
  '152. Inclusão na Tabela de Parâmetros Filial
  '     Incluído campo [Zero a Esquerda]
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Zero a Esquerda") Then
    If gbCreateField("Parâmetros Filial", "Zero a Esquerda", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET [Zero a Esquerda] = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/05/2004 - Daniel
  '
  '153. Adiciona um registro novo na ZzzProgramas
  '     Manutenção de Reservas, case PSV
  nPhase = nPhase + 1
  If AddFileZZZProgramas3() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '21/06/2004 - Daniel
  '
  'Case: Coneg Campos
  '
  '154. Inclusão na Tabela de Funcionários
  '     Incluído campo SenhaClear
  '     Senha para impedir que funcionários não limpem a
  '     tela de venda rápida e saídas
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "SenhaClear") Then
    If gbCreateField("Funcionários", "SenhaClear", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET SenhaClear = TRUE;" '02/07/2004 - Daniel - Alterado de False para True
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '30/06/2004 - Daniel
  '
  'Case: Nazareno
  '
  '155. Inclusão na Tabela de Cli_For - Crédito
  '     Incluído campo PercentualLimite
  '     Este campo será o percentual sobre o salário
  '     para o preenchimento do campo limite
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For - Crédito", "PercentualLimite") Then
    If gbCreateField("Cli_For - Crédito", "PercentualLimite", dbDouble) Then
      db.Execute "UPDATE [Cli_For - Crédito] SET PercentualLimite = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For - Crédito"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '07/07/2004 - Daniel
  '156. Esta tabela foi desenvolvida para a TV Shopping
  '     Curva A B C D
  '     Table: GruposDeClientes
  nPhase = nPhase + 1
  If gbGetTable("GruposDeClientes") = False Then
    If gbCreateTableGruposDeClientes() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""GruposDeClientes"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '12/07/2004 - Daniel
  '
  'Case: TV Shopping
  '
  '157. Inclusão na Tabela de Funcionários
  '     Incluído campo Marketing
  '     Quando o user que possuir este campo habilitado
  '     entrar no Quick, será disparado uma rotina para
  '     atualização dos clientes dentro do Grupo de Classificação
  '     Curva A B C D
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "Marketing") Then
    If gbCreateField("Funcionários", "Marketing", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET Marketing = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '13/07/2004 - Daniel
  '
  'Case: TV Shopping
  '
  '158. Inclusão na Tabela Cli_For
  '     Incluído campo CodGrupo
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "CodGrupo") Then
    If gbCreateField("Cli_For", "CodGrupo", dbByte) Then
      db.Execute "UPDATE Cli_For SET CodGrupo = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '13/07/2004 - Daniel
  '
  'Case: TV Shopping
  '
  '159. Inclusão na Tabela Cli_For
  '     Incluído campo TotDinheiroBoletos
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "TotDinheiroBoletos") Then
    If gbCreateField("Cli_For", "TotDinheiroBoletos", dbDouble) Then
      db.Execute "UPDATE Cli_For SET TotDinheiroBoletos = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '13/07/2004 - Daniel
  '
  'Case: TV Shopping
  '
  '160. Inclusão na Tabela Cli_For
  '     Incluído campo TotCheques
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "TotCheques") Then
    If gbCreateField("Cli_For", "TotCheques", dbDouble) Then
      db.Execute "UPDATE Cli_For SET TotCheques = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '13/07/2004 - Daniel
  '
  'Case: TV Shopping
  '
  '161. Inclusão na Tabela Cli_For
  '     Incluído campo TotCartoes
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "TotCartoes") Then
    If gbCreateField("Cli_For", "TotCartoes", dbDouble) Then
      db.Execute "UPDATE Cli_For SET TotCartoes = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '13/07/2004 - Daniel
  '
  'Case: TV Shopping
  '
  '162. Inclusão na Tabela Cli_For
  '     Incluído campo TotRecebido
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "TotRecebido") Then
    If gbCreateField("Cli_For", "TotRecebido", dbDouble) Then
      db.Execute "UPDATE Cli_For SET TotRecebido = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '14/07/2004 - Daniel
  '
  '163. Adiciona um registro novo na ZzzProgramas
  '     Classificação dos Clientes, case TV Shopping
  nPhase = nPhase + 1
  If AddFileZZZProgramas4() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '26/07/2004 - Daniel
  '
  'Case: STC - Sistema Trídio de Comunicação
  'Alteração: A seguir temos a criação 04 campos para a tabela Programacao
  '           Cancel1, Cancel2, Cancel3, Cancel4 que possuem a finalidade de
  '           monitorar se a parcela da programação foi ou não cancelada
  '
  '164. Inclusão na Tabela Programacao
  '     do Campo Cancel1
  nPhase = nPhase + 1
  If Not gbGetField("Programacao", "Cancel1") Then
    If gbCreateField("Programacao", "Cancel1", dbBoolean) Then
      db.Execute "UPDATE Programacao SET Cancel1 = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Programacao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '165. Inclusão na Tabela Programacao
  '     do Campo Cancel2
  nPhase = nPhase + 1
  If Not gbGetField("Programacao", "Cancel2") Then
    If gbCreateField("Programacao", "Cancel2", dbBoolean) Then
      db.Execute "UPDATE Programacao SET Cancel2 = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Programacao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '166. Inclusão na Tabela Programacao
  '     do Campo Cancel3
  nPhase = nPhase + 1
  If Not gbGetField("Programacao", "Cancel3") Then
    If gbCreateField("Programacao", "Cancel3", dbBoolean) Then
      db.Execute "UPDATE Programacao SET Cancel3 = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Programacao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '167. Inclusão na Tabela Programacao
  '     do Campo Cancel4
  nPhase = nPhase + 1
  If Not gbGetField("Programacao", "Cancel4") Then
    If gbCreateField("Programacao", "Cancel4", dbBoolean) Then
      db.Execute "UPDATE Programacao SET Cancel4 = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Programacao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '168. Inclusão na Tabela Cli_For
  '     do Campo AgenciaPublicidade
  '     Case: STC
  '     Histórico: Este campo boleano fará a diferença entre
  '     Fornecedor e Agência de Publicidade
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "AgenciaPublicidade") Then
    If gbCreateField("Cli_For", "AgenciaPublicidade", dbBoolean) Then
      db.Execute "UPDATE Cli_For SET AgenciaPublicidade = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '169. Inclusão na tabela Entradas de campo pra dizer se a consignação está fechada
  '     Case: Julio Sampaio
  '     Maikel Cordeiro
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "ConsignacaoFechada") Then
    If gbCreateField("Entradas", "ConsignacaoFechada", dbBoolean) Then
      db.Execute "UPDATE Entradas SET ConsignacaoFechada = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '170. Inclusão na tabela [Entradas - Produtos] de campo pra dizer se a consignação está fechada
  '     Case: Julio Sampaio
  '     Maikel Cordeiro
  nPhase = nPhase + 1
  If Not gbGetField("Entradas - Produtos", "ConsignacaoFechada") Then
    If gbCreateField("Entradas - Produtos", "ConsignacaoFechada", dbBoolean) Then
      db.Execute "UPDATE [Entradas - Produtos] SET ConsignacaoFechada = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '29/07/2004 - Daniel
  '
  '171. Criação da Tabela Supervisores
  '     Esta table foi criada para atender inicialmente a necessidade
  '     da STC
  nPhase = nPhase + 1
  If gbGetTable("Supervisores") = False Then
    If gbCreateTableSupervisores() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Supervisores"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '172. Inclusão na Tabela Funcionários
  '     do Campo Supervisor
  '
  '     Histórico: Este campo integer estará
  '     fazendo relacionamento com a table Funcionários
  '     Um Supervisor para 'n' Funcionários
  '     Supervisor.Código 1 |---| n Funcionários.Supervisor
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "Supervisor") Then
    If gbCreateField("Funcionários", "Supervisor", dbInteger) Then
      db.Execute "UPDATE Funcionários SET Supervisor = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '02/08/2004 - Daniel
  '
  '173. Criação da Tabela ParamFaturameAuto
  '     Esta table foi criada para atender inicialmente a necessidade
  '     da STC. Nela setaremos valores padrões para faturamento sobre os
  '     'Serviços' da Rádio
  nPhase = nPhase + 1
  If gbGetTable("ParamFaturameAuto") = False Then
    If gbCreateTableParamFaturameAuto() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ParamFaturameAuto"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '10/08/2004 - Maikel
  '
  '174. Criação da Tabela Acerto Consignações
  '     Esta table foi criada para atender inicialmente a necessidade
  '     da STC
  nPhase = nPhase + 1
  If gbGetTable("AcertoConsignacaoEntrada") = False Then
    If gbCreateTableAcertoConsignacaoEntrada() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""AcertoConsignacaoEntrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '175. Inclusão na Tabela Contrato
  '     do Campo VlTotContrato
  '
  '     Case: STC
  nPhase = nPhase + 1
  If Not gbGetField("Contrato", "VlTotContrato") Then
    If gbCreateField("Contrato", "VlTotContrato", dbDouble) Then
      db.Execute "UPDATE Contrato SET VlTotContrato = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Contrato"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '17/08/2004 - Daniel
  '
  '176. Inclusão na Tabela Parâmetros Filial
  '     do Campo TaxaDesconto
  '
  '     Case: Nazareno - liberado para todos
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "TaxaDesconto") Then
    If gbCreateField("Parâmetros Filial", "TaxaDesconto", dbDouble) Then
      db.Execute "UPDATE [Parâmetros Filial] SET TaxaDesconto = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '17/08/2004 - Daniel
  '
  '177. Inclusão na Tabela Parâmetros Filial
  '     do Campo BoletoPadrao
  '
  '     Case: Nazareno - liberado para todos
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "BoletoPadrao") Then
    If gbCreateField("Parâmetros Filial", "BoletoPadrao", dbText, 30) Then
      db.Execute "UPDATE [Parâmetros Filial] SET BoletoPadrao = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/08/2004 - Daniel
  '
  '178. Inclusão na Tabela Programacao
  '     do Campo ImpressoNF
  '
  '     Case: STC
  nPhase = nPhase + 1
  If Not gbGetField("Programacao", "ImpressoNF") Then
    If gbCreateField("Programacao", "ImpressoNF", dbBoolean) Then
      db.Execute "UPDATE Programacao SET ImpressoNF = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Programacao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/08/2004 - Daniel
  '
  '179. Inclusão na Tabela Parâmetros Filial
  '     do Campo TicketPadrao
  '
  '     Case: STC
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "TicketPadrao") Then
    If gbCreateField("Parâmetros Filial", "TicketPadrao", dbText, 30) Then
      db.Execute "UPDATE [Parâmetros Filial] SET TicketPadrao = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/08/2004 - Daniel
  '
  '180. Inclusão na Tabela Programacao
  '     do Campo SomaCancelamento
  '
  '     Case: STC
  nPhase = nPhase + 1
  If Not gbGetField("Programacao", "SomaCancelamento") Then
    If gbCreateField("Programacao", "SomaCancelamento", dbDouble) Then
      db.Execute "UPDATE Programacao SET SomaCancelamento = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Programacao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
   
   
  '20/08/2004 - Daniel
  '
  '181. Inclusão na Tabela Produtos
  '     do Campo UsaDescrAdic
  '
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "UsaDescrAdic") Then
    If gbCreateField("Produtos", "UsaDescrAdic", dbBoolean) Then
      db.Execute "UPDATE Produtos SET UsaDescrAdic = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
   
   
  '23/08/2004 - Daniel
  '
  '182. Inclusão na Tabela Operações Entrada
  '     do Campo Tabela
  '
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "Tabela") Then
    If gbCreateField("Operações Entrada", "Tabela", dbText, 15) Then
      db.Execute "UPDATE [Operações Entrada] SET Tabela = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
   
   
  '20/08/2004 - Daniel
  '
  '183. Inclusão na Tabela Produtos
  '     do Campo IndiceFinanceiro
  '
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "IndiceFinanceiro") Then
    If gbCreateField("Produtos", "IndiceFinanceiro", dbDouble) Then
      db.Execute "UPDATE Produtos SET IndiceFinanceiro = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If


  '24/08/2004 - Daniel
  '
  '184. Inclusão na Tabela [Entradas - Produtos]
  '     do Campo IndiceFinanceiro
  '
  nPhase = nPhase + 1
  If Not gbGetField("Entradas - Produtos", "IndiceFinanceiro") Then
    If gbCreateField("Entradas - Produtos", "IndiceFinanceiro", dbDouble) Then
      db.Execute "UPDATE [Entradas - Produtos] SET IndiceFinanceiro = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If


  '25/08/2004 - Daniel
  '
  '185. Inclusão na Tabela Operações Entrada
  '     do Campo PermitirAlterPreco
  '
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "PermitirAlterPreco") Then
    If gbCreateField("Operações Entrada", "PermitirAlterPreco", dbBoolean) Then
      db.Execute "UPDATE [Operações Entrada] SET PermitirAlterPreco = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  
  '27/08/2004 - Daniel
  '
  '186. Inclusão na Tabela Operações Saída
  '     do Campo AcertaEmprestimoEntrada
  '
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "AcertaEmprestimoEntrada") Then
    If gbCreateField("Operações Saída", "AcertaEmprestimoEntrada", dbBoolean) Then
      db.Execute "UPDATE [Operações Saída] SET AcertaEmprestimoEntrada = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '31/08/2004 - Daniel
  '
  '187. Inclusão na Tabela [Entradas - Produtos]
  '     do Campo QtdeAtual
  '
  nPhase = nPhase + 1
  If Not gbGetField("Entradas - Produtos", "QtdeAtual") Then
    If gbCreateField("Entradas - Produtos", "QtdeAtual", dbSingle) Then
      db.Execute "UPDATE [Entradas - Produtos] SET QtdeAtual = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '15/09/2004 - Daniel
  '
  '188. Criação da Tabela ParamDevoMat
  '     Esta table foi criada para atender inicialmente a necessidade
  '     da Livraria Resultado
  nPhase = nPhase + 1
  If gbGetTable("ParamDevoMat") = False Then
    If gbCreateTableParamDevoMat() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ParamDevoMat"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '16/09/2004 - Daniel
  '
  '189. Inclusão na Tabela [Entradas - Produtos]
  '     do Campo Selecionado
  '
  nPhase = nPhase + 1
  If Not gbGetField("Entradas - Produtos", "Selecionado") Then
    If gbCreateField("Entradas - Produtos", "Selecionado", dbBoolean) Then
      db.Execute "UPDATE [Entradas - Produtos] SET Selecionado = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '16/09/2004 - Daniel
  '
  '190. Inclusão na Tabela [Entradas - Produtos]
  '     do Campo Acertado
  '
  nPhase = nPhase + 1
  If Not gbGetField("Entradas - Produtos", "Acertado") Then
    If gbCreateField("Entradas - Produtos", "Acertado", dbBoolean) Then
      db.Execute "UPDATE [Entradas - Produtos] SET Acertado = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '17/09/2004 - Daniel
  '
  '191. Criação da Tabela PrestacaoContas
  '     Esta table foi criada para atender inicialmente a necessidade
  '     da Livraria Resultado
  nPhase = nPhase + 1
  If gbGetTable("PrestacaoContas") = False Then
    If gbCreateTablePrestacaoContas() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""PrestacaoContas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '22/09/2004 - Daniel
  '
  '192. Inclusão na Tabela [Entradas - Produtos]
  '     do Campo EntradaConsignada
  '
  '     Este campo identifica que a [Entradas - Produtos] é
  '     uma Consignação
  '
  nPhase = nPhase + 1
  If Not gbGetField("Entradas - Produtos", "EntradaConsignada") Then
    If gbCreateField("Entradas - Produtos", "EntradaConsignada", dbBoolean) Then
      db.Execute "UPDATE [Entradas - Produtos] SET Acertado = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/09/2004 - mpdea
  '
  'Número     : 193
  'Case       : Embalavi
  'Descrição  : Inclusão na tabela Produtos do campo Volumagem
  'Finalidade : Permitir impressão do cálculo de volumes por quantidade
  '             de itens da movimentação
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "Volumagem") Then
    If gbCreateField("Produtos", "Volumagem", dbInteger) Then
      db.Execute "UPDATE Produtos SET Volumagem = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '01/10/2004 - Daniel
  '
  '194. Adiciona um registro novo na ZzzProgramas
  '     Gerenciador de Loja Virtual, Case Resultado
  nPhase = nPhase + 1
  If AddFileZZZProgramas5() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '04/10/2004 - Daniel
  '
  '195. Inclusão na Tabela PrestacaoContas
  '     do Campo PrestacaoFechada
  '
  nPhase = nPhase + 1
  If Not gbGetField("PrestacaoContas", "PrestacaoFechada") Then
    If gbCreateField("PrestacaoContas", "PrestacaoFechada", dbBoolean) Then
      db.Execute "UPDATE PrestacaoContas SET PrestacaoFechada = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""PrestacaoContas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/10/2004 - Daniel
  '
  '196. Inclusão na Tabela PrestacaoContas
  '     do Campo CompraFechada
  '
  nPhase = nPhase + 1
  If Not gbGetField("PrestacaoContas", "CompraFechada") Then
    If gbCreateField("PrestacaoContas", "CompraFechada", dbBoolean) Then
      db.Execute "UPDATE PrestacaoContas SET CompraFechada = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""PrestacaoContas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '08/10/2004 - Daniel
  '
  '197. Inclusão na Tabela PrestacaoContas
  '     do Campo PeriodoVenda
  '
  nPhase = nPhase + 1
  If Not gbGetField("PrestacaoContas", "PeriodoVenda") Then
    If gbCreateField("PrestacaoContas", "PeriodoVenda", dbDate) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""PrestacaoContas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '08/10/2004 - Daniel
  '
  '198. Inclusão na Tabela PrestacaoContas
  '     do Campo NotaFiscal
  '
  nPhase = nPhase + 1
  If Not gbGetField("PrestacaoContas", "NotaFiscal") Then
    If gbCreateField("PrestacaoContas", "NotaFiscal", dbLong) Then
      db.Execute "UPDATE PrestacaoContas SET NotaFiscal = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""PrestacaoContas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '08/10/2004 - Daniel
  '
  '199. Inclusão na Tabela PrestacaoContas
  '     do Campo QtdeAcertada
  '
  nPhase = nPhase + 1
  If Not gbGetField("PrestacaoContas", "QtdeAcertada") Then
    If gbCreateField("PrestacaoContas", "QtdeAcertada", dbDouble) Then
      db.Execute "UPDATE PrestacaoContas SET QtdeAcertada = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""PrestacaoContas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '14/10/2004 - Daniel
  '
  '200. Inclusão na Tabela AcertoConsignacaoEntrada
  '     do Campo Fornecedor
  '
  nPhase = nPhase + 1
  If Not gbGetField("AcertoConsignacaoEntrada", "PrecoCusto") Then
    If gbCreateField("AcertoConsignacaoEntrada", "PrecoCusto", dbDouble) Then
      db.Execute "UPDATE AcertoConsignacaoEntrada SET PrecoCusto = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""AcertoConsignacaoEntrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '16/10/2004 - Daniel
  '
  '201. Inclusão na Tabela AcertoConsignacaoEntrada
  '     do Campo PrecoVenda
  '
  nPhase = nPhase + 1
  If Not gbGetField("AcertoConsignacaoEntrada", "PrecoVenda") Then
    If gbCreateField("AcertoConsignacaoEntrada", "PrecoVenda", dbDouble) Then
      db.Execute "UPDATE AcertoConsignacaoEntrada SET PrecoVenda = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""AcertoConsignacaoEntrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '01/11/2004 - Daniel
  '
  '202. Inclusão na Tabela Produtos
  '     do Campo Cubagem
  '
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "Cubagem") Then
    If gbCreateField("Produtos", "Cubagem", dbDouble) Then
      db.Execute "UPDATE Produtos SET Cubagem = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '03/11/2004 - Daniel
  '
  '203. Inclusão na Tabela [Preços - Tempo]
  '     do [PreçoNacional 1]
  '
  nPhase = nPhase + 1
  If Not gbGetField("Preços - Tempo", "PreçoNacional 1") Then
    If gbCreateField("Preços - Tempo", "PreçoNacional 1", dbDouble) Then
      db.Execute "UPDATE [Preços - Tempo] SET [PreçoNacional 1] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Preços - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '03/11/2004 - Daniel
  '
  '204. Inclusão na Tabela [Preços - Tempo]
  '     do [PreçoNacional 2]
  '
  nPhase = nPhase + 1
  If Not gbGetField("Preços - Tempo", "PreçoNacional 2") Then
    If gbCreateField("Preços - Tempo", "PreçoNacional 2", dbDouble) Then
      db.Execute "UPDATE [Preços - Tempo] SET [PreçoNacional 2] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Preços - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '03/11/2004 - Daniel
  '
  '205. Inclusão na Tabela [Preços - Tempo]
  '     do [PreçoNacional 3]
  '
  nPhase = nPhase + 1
  If Not gbGetField("Preços - Tempo", "PreçoNacional 3") Then
    If gbCreateField("Preços - Tempo", "PreçoNacional 3", dbDouble) Then
      db.Execute "UPDATE [Preços - Tempo] SET [PreçoNacional 3] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Preços - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '03/11/2004 - Daniel
  '
  '206. Inclusão na Tabela [Preços - Tempo]
  '     do [PreçoNacional 4]
  '
  nPhase = nPhase + 1
  If Not gbGetField("Preços - Tempo", "PreçoNacional 4") Then
    If gbCreateField("Preços - Tempo", "PreçoNacional 4", dbDouble) Then
      db.Execute "UPDATE [Preços - Tempo] SET [PreçoNacional 4] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Preços - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '03/11/2004 - Daniel
  '
  '207. Inclusão na Tabela [Preços - Tempo]
  '     do [PreçoNacional 5]
  '
  nPhase = nPhase + 1
  If Not gbGetField("Preços - Tempo", "PreçoNacional 5") Then
    If gbCreateField("Preços - Tempo", "PreçoNacional 5", dbDouble) Then
      db.Execute "UPDATE [Preços - Tempo] SET [PreçoNacional 5] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Preços - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '03/11/2004 - Daniel
  '
  '208. Inclusão na Tabela [Preços - Tempo]
  '     do [PreçoNacional 6]
  '
  nPhase = nPhase + 1
  If Not gbGetField("Preços - Tempo", "PreçoNacional 6") Then
    If gbCreateField("Preços - Tempo", "PreçoNacional 6", dbDouble) Then
      db.Execute "UPDATE [Preços - Tempo] SET [PreçoNacional 6] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Preços - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '10/11/2004 - Daniel
  '
  '209. Inclusão na Tabela Produtos
  '     do Campo Lote
  '
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "Lote") Then
    If gbCreateField("Produtos", "Lote", dbText, 15) Then
      db.Execute "UPDATE Produtos SET Lote = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '10/11/2004 - Daniel
  '
  '210. Inclusão na Tabela Produtos
  '     do Campo DataValidade
  '
  nPhase = nPhase + 1
  If gbGetField("Produtos", "DataValidade") = False Then
    If gbCreateField("Produtos", "DataValidade", dbDate) = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '11/11/2004 - Daniel
  '
  '211. Novo registro em ZzzProgramas
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas6() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '11/11/2004 - Daniel
  '
  '212. Novo registro em ZzzProgramas
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas7() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '11/11/2004 - Daniel
  '
  '213. Novo registro em ZzzProgramas
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas8() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '11/11/2004 - Daniel
  '
  '214. Novo registro em ZzzProgramas
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas9() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '11/11/2004 - Daniel
  '
  '215. Novo registro em ZzzProgramas
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas10() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '11/11/2004 - Daniel
  '
  '216. Novo registro em ZzzProgramas
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas11() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '29/11/2004 - Daniel
  '
  '217. Inclusão de campo
  '     Tabela    : [Parâmetros Filial]
  '     Campo     : Permitir5Casas
  '     Finalidade: Permitir 5 casas após a vírgula no preço unitário
  '                 em Entradas
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Permitir5Casas") Then
    If gbCreateField("Parâmetros Filial", "Permitir5Casas", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET Permitir5Casas = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '13/12/2004 - Daniel
  '
  '218. Criação da Tabela BooksVendidos
  '     Esta table foi criada para atender inicialmente a necessidade
  '     da Livraria Resultado
  nPhase = nPhase + 1
  If gbGetTable("BooksVendidos") = False Then
    If gbCreateTableBooksVendidos() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""BooksVendidos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '14/01/2005 - Daniel
  '
  '219. Inclusão de campo
  '     Tabela     : [Consignação Saída]
  '     Campo      : QtdeVendidaAcumulada
  '     Solicitante: Aura Prata
  '     Finalidade : Armazenar a Qtde Vendida Acumulada para cada registro
  '
  nPhase = nPhase + 1
  If Not gbGetField("Consignação Saída", "QtdeVendidaAcumulada") Then
    If gbCreateField("Consignação Saída", "QtdeVendidaAcumulada", dbDouble) Then
      db.Execute "UPDATE [Consignação Saída] SET QtdeVendidaAcumulada = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Consignação Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/01/2005 - Daniel
  '
  '220. Alteração na Tabela [Contas a Receber] do Campo CNAB_NossoNumero
  '     Alterado o tamanho do campo de 20 para 21 caracteres, em 09/01/2005
  '     passou de 21 para 40
  '
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Contas a Receber", "CNAB_NossoNumero", dbText, 40, "CNAB_NossoNumero") Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Contas a Receber"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '28/01/2005 - Daniel
  '
  '221. Alteração na Tabela Caixa do Campo Descrição
  '     Alterado o tamanho do campo de 60 para 120 caracteres
  '
  '     Solicitante: Taupys - ES
  '
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Caixa", "Descrição", dbText, 120, "Descrição") Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Caixa"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '17/02/2005 - Daniel
  '
  '222. Inclusão de campo
  '     Tabela     : [Etiquetas - Tempo]
  '     Campo      : Descricao2
  '     Solicitante: Mozart (Hello Kyt)
  '     Finalidade : Armazenar apenas o nome do produto
  '
  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "Descricao2") Then
    If gbCreateField("Etiquetas - Tempo", "Descricao2", dbText, 70) Then
      db.Execute "UPDATE [Etiquetas - Tempo] SET Descricao2 = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/02/2005 - Daniel
  '
  '223. Inclusão na Tabela [Operações Saída]
  '     do Campo InformanteProprio
  '
  'Solicitante: Agrofarm - RS
  '
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "InformanteProprio") Then
    If gbCreateField("Operações Saída", "InformanteProprio", dbBoolean) Then
      db.Execute "UPDATE [Operações Saída] SET InformanteProprio = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/02/2005 - Daniel
  '
  '224. Inclusão na Tabela [Operações Entrada]
  '     do Campo InformanteProprio
  '
  'Solicitante: Agrofarm - RS
  '
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "InformanteProprio") Then
    If gbCreateField("Operações Entrada", "InformanteProprio", dbBoolean) Then
      db.Execute "UPDATE [Operações Entrada] SET InformanteProprio = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '01/03/2005 - Daniel
  '
  '225. Inclusão de índice na tabela de saídas
  '
  'Solicitação: Red Line - RJ
  '
  nPhase = nPhase + 1
  If Not g_blnGetIndex("Saídas", "Nota") Then
    If Not m_blnCreateIndexNota() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '01/03/2005 - Daniel
  '
  '226. Inclusão de índice na tabela de funcionários
  '
  'Solicitação: Red Line - RJ
  '
  nPhase = nPhase + 1
  If Not g_blnGetIndex("Funcionários", "Acessando") Then
    If Not m_blnCreateIndexAcessando() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '15/03/2005 - Daniel
  '
  '227. Inclusão de índice na tabela de [Contas a Receber]
  '
  'Solicitação: TV Shopping
  '
  nPhase = nPhase + 1
  If Not g_blnGetIndex("Contas a Receber", "CNAB") Then
    If Not m_blnCreateIndexCNAB() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '21/03/2005 - Daniel
  '
  '228. Criação da Tabela Retencao
  '     Esta table foi criada para atender inicialmente a necessidade
  '     da Enxovais Bem Me Quer
  nPhase = nPhase + 1
  If gbGetTable("Retencao") = False Then
    If gbCreateTableRetencao() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Retencao"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '22/03/2005 - Daniel
  '
  '229. Inclusão na Tabela de Saídas
  '     do Campo CodigoRetencao
  '
  'Solicitante: Bem Me Quer
  '
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "CodigoRetencao") Then
    If gbCreateField("Saídas", "CodigoRetencao", dbInteger) Then
      db.Execute "UPDATE Saídas SET CodigoRetencao = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '29/03/2005 - Daniel
  '
  '230. Inclusão de índice na tabela de Produtos
  '     Para otimizar a busca do fabricante
  '
  'Solicitação: Elétrica Leal
  '
  nPhase = nPhase + 1
  If Not g_blnGetIndex("Produtos", "Fabricante") Then
    If Not m_blnCreateIndexFabricante() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '01/04/2005 - Daniel
  '
  '231. Inclusão na Tabela de Saídas
  '     do Campo Seguro
  '
  'Solicitante: Aura Prata
  '
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Seguro") Then
    If gbCreateField("Saídas", "Seguro", dbDouble) Then
      db.Execute "UPDATE Saídas SET Seguro = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '12/04/2005 - Daniel
  '
  '232. Inclusão na Tabela de Operações Saída
  '     do Campo SomarSeguro
  '
  'Solicitante: Aura Prata
  '
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "SomarSeguro") Then
    If gbCreateField("Operações Saída", "SomarSeguro", dbBoolean) Then
      db.Execute "UPDATE [Operações Saída] SET SomarSeguro = False;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '19/04/2005 - Daniel
  '
  '233. Inclusão na Tabela de Parâmetros Filial
  '     do Campo CliWebComprarPrazo
  '
  'Solicitante.: Aura Prata
  '
  'Finalidade..: No momento do recebimento quando ocorrer a mensagem que o cliente não é
  '              habilitado para para fazer compras a prazo, se este flag estiver como
  '              true evitará do usuário voltar na tela de clientes para habilitar o campo
  '              comprar a prazo e depois fazer o recebimento.
  '
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "CliWebComprarPrazo") Then
    If gbCreateField("Parâmetros Filial", "CliWebComprarPrazo", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET CliWebComprarPrazo = False;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '25/04/2005 - Daniel
  '
  '234. Inclusão na Tabela de Comissão
  '     do Campo VlPagoEmCartao
  '
  'Solicitante: Bem Me Quer
  '
  nPhase = nPhase + 1
  If Not gbGetField("Comissão", "VlPagoEmCartao") Then
    If gbCreateField("Comissão", "VlPagoEmCartao", dbDouble) Then
      db.Execute "UPDATE Comissão SET VlPagoEmCartao = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Comissão"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '25/04/2005 - Daniel
  '
  '235. Inclusão na Tabela de Comissão
  '     do Campo VlPagoEmCartaoComRetencao
  '
  'Solicitante: Bem Me Quer
  '
  nPhase = nPhase + 1
  If Not gbGetField("Comissão", "VlPagoEmCartaoComRetencao") Then
    If gbCreateField("Comissão", "VlPagoEmCartaoComRetencao", dbDouble) Then
      db.Execute "UPDATE Comissão SET VlPagoEmCartaoComRetencao = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Comissão"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '25/04/2005 - Daniel
  '
  '236. Inclusão na Tabela de Comissão
  '     do Campo TaxaRetencao
  '
  'Solicitante: Bem Me Quer
  '
  nPhase = nPhase + 1
  If Not gbGetField("Comissão", "TaxaRetencao") Then
    If gbCreateField("Comissão", "TaxaRetencao", dbSingle) Then
      db.Execute "UPDATE Comissão SET TaxaRetencao = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Comissão"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/04/2005 - Daniel
  '
  '237. Inclusão na Tabela de [Contas a Pagar]
  '     do Campo OrigemDinheiro
  '
  'Solicitante: Bem Me Quer
  '
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Pagar", "OrigemDinheiro") Then
    If gbCreateField("Contas a Pagar", "OrigemDinheiro", dbText) Then
      db.Execute "UPDATE [Contas a Pagar] SET OrigemDinheiro = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Contas a Pagar"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/04/2005 - Daniel
  '
  '238. Inclusão na Tabela de [Parâmetros Filial]
  '     do Campo VerificaLimiteCli
  '
  'Solicitante..: Jorge Marcos - PSI MT
  '
  'Finalidade...: Verificar o limite de crédito do cliente antes da gravação
  '               Isto é essencial para todas as empresas que trabalham
  '               com pronta entrega
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VerificaLimiteCli") Then
    If gbCreateField("Parâmetros Filial", "VerificaLimiteCli", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET VerificaLimiteCli = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '05/05/2005 - Daniel
  '
  '239. Inclusão na Tabela de [Centros de Custo]
  '     do Campo Ativo
  '
  'Solicitante..: Carlos
  '
  'Finalidade...: Desativar Centros de Custo não mais utilizados ou
  '               habilitá-los através deste campo
  nPhase = nPhase + 1
  If Not gbGetField("Centros de Custo", "Ativo") Then
    If gbCreateField("Centros de Custo", "Ativo", dbBoolean) Then
      db.Execute "UPDATE [Centros de Custo] SET Ativo = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Centros de Custo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '06/05/2005 - Daniel
  '
  '240. Inclusão na Tabela de [Parâmetros Filial]
  '     do Campo UtilizarCodFornec
  '
  'Solicitante..: Cristiano Pavinato - PSI RS
  '
  'Finalidade...: Nas telas de Entrada, Saída e Venda Rápida ao entrar
  '               no campo código do produto com o código p/ fornecedor
  '               cadastrado na tela de produtos, este código p/ fornecedor
  '               fará a busca do código do produto que estiver amarrado nele
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "UtilizarCodFornec") Then
    If gbCreateField("Parâmetros Filial", "UtilizarCodFornec", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET UtilizarCodFornec = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '12/05/2005 - Daniel
  '
  '241. Inclusão na Tabela de [Parâmetros Filial]
  '     do Campo ExibirFabricante
  '
  'Solicitante..: Info Social
  '
  'Finalidade...: Deixamos configurável à exibição nas telas de
  '               Saída e Venda Rápida da coluna Fabricante nos
  '               dropdowns de pesquisas
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ExibirFabricante") Then
    If gbCreateField("Parâmetros Filial", "ExibirFabricante", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET ExibirFabricante = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '17/05/2005 - Daniel
  '
  '242. Inclusão na Tabela de Produtos
  '     do Campo ImprimirUmaEtiq
  '
  'Solicitante..: Miss Nuvem
  '
  'Finalidade...: Este campo gerenciará a impressão de uma ou duas
  '               etiquetas no momento da impressão na térmica
  '               O default de etiquetas para a Miss Nuvem é duas
  '               em cada linha de impressão
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "ImprimirUmaEtiq") Then
    If gbCreateField("Produtos", "ImprimirUmaEtiq", dbBoolean) Then
      db.Execute "UPDATE Produtos SET ImprimirUmaEtiq = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '17/05/2005 - Daniel
  '
  '243. Inclusão na Tabela de Produtos
  '     do Campo ImprimirPrecoEtiq
  '
  'Solicitante..: Miss Nuvem
  '
  'Finalidade...: Este campo gerenciará a impressão do preço
  '               na etiqueta ou não
  '               Default inicial TRUE
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "ImprimirPrecoEtiq") Then
    If gbCreateField("Produtos", "ImprimirPrecoEtiq", dbBoolean) Then
      db.Execute "UPDATE Produtos SET ImprimirPrecoEtiq = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/05/2005 - Daniel
  '
  '244. Inclusão na Tabela de Saídas
  '     do Campo [Nota Fiscal]
  '
  'Solicitante..: Pedágio
  '
  'Finalidade...: Neste campo será armazenado o valor que o usuário
  '               digitar para a nota de saída quando ocorrer criação
  '               de nota manual monitorada pela operação de saída
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Nota Fiscal") Then
    If gbCreateField("Saídas", "Nota Fiscal", dbLong) Then
      db.Execute "UPDATE Saídas SET [Nota Fiscal] = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '18/05/2005 - Daniel
  '
  '245. Inclusão na Tabela de Saídas
  '     do Campo SerieNF
  '
  'Solicitante..: Pedágio
  '
  'Finalidade...: Neste campo será armazenado a Série da NF
  '               que o usuário digitar para a nota de saída
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "SerieNF") Then
    If gbCreateField("Saídas", "SerieNF", dbText, 3) Then
      db.Execute "UPDATE Saídas SET SerieNF = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '19/05/2005 - Daniel
  '
  '246. Inclusão na Tabela de Entradas
  '     do Campo SerieNF
  '
  'Solicitante..: Pedágio
  '
  'Finalidade...: Neste campo será armazenado a Série da NF
  '               que o usuário digitar para a nota de entrada
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "SerieNF") Then
    If gbCreateField("Entradas", "SerieNF", dbText, 3) Then
      db.Execute "UPDATE Entradas SET SerieNF = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '19/05/2005 - Daniel
  '
  '247. Inclusão na Tabela de [Operações Saída]
  '     do Campo EmitirNFManualmente
  '
  'Solicitante..: Pedágio
  '
  'Finalidade...: Quando esta operação estiver atrelado na saída
  '               o contador do Quick estará em pause e o número
  '               da nota será inserido em Saídas.[Nota Fiscal]
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "EmitirNFManualmente") Then
    If gbCreateField("Operações Saída", "EmitirNFManualmente", dbBoolean) Then
      db.Execute "UPDATE [Operações Saída] SET EmitirNFManualmente = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '19/05/2005 - Daniel
  '
  '248. Inclusão na Tabela de [Operações Entrada]
  '     do Campo EmitirNFManualmente
  '
  'Solicitante..: Pedágio
  '
  'Finalidade...: Quando esta operação estiver atrelado na entrada
  '               o contador do Quick estará em pause e o número
  '               da nota será inserido em Entradas.[Nota Fiscal]
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "EmitirNFManualmente") Then
    If gbCreateField("Operações Entrada", "EmitirNFManualmente", dbBoolean) Then
      db.Execute "UPDATE [Operações Entrada] SET EmitirNFManualmente = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/05/2005 - Daniel
  '
  '249. Inclusão de campo
  '     Tabela     : [Etiquetas - Tempo]
  '     Campo      : ImprimirUmaEtiq
  '     Solicitante: Miss Nuvem
  '     Finalidade : Interação com a table Produtos
  '
  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "ImprimirUmaEtiq") Then
    If gbCreateField("Etiquetas - Tempo", "ImprimirUmaEtiq", dbBoolean) Then
      db.Execute "UPDATE [Etiquetas - Tempo] SET ImprimirUmaEtiq = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/05/2005 - Daniel
  '
  '250. Inclusão de campo
  '     Tabela     : [Etiquetas - Tempo]
  '     Campo      : ImprimirPrecoEtiq
  '     Solicitante: Miss Nuvem
  '     Finalidade : Interação com a table Produtos
  '
  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "ImprimirPrecoEtiq") Then
    If gbCreateField("Etiquetas - Tempo", "ImprimirPrecoEtiq", dbBoolean) Then
      db.Execute "UPDATE [Etiquetas - Tempo] SET ImprimirPrecoEtiq = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/05/2005 - Daniel
  '
  '251. Inclusão de campo
  '     Tabela     : Entradas
  '     Campo      : InfoICMSporUF
  '     Solicitante: Cristiano Pavinato - PSI RS
  '     Finalidade : Atender os usuários do Rio Grande do Sul
  '                  que geram um arquivo do sintegra por estado
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "InfoICMSporUF") Then
    If gbCreateField("Entradas", "InfoICMSporUF", dbBoolean) Then
      db.Execute "UPDATE Entradas SET InfoICMSporUF = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/05/2005 - Daniel
  '
  '252. Inclusão de campo
  '     Tabela     : Saídas
  '     Campo      : InfoICMSporUF
  '     Solicitante: Cristiano Pavinato - PSI RS
  '     Finalidade : Atender os usuários do Rio Grande do Sul
  '                  que geram um arquivo do sintegra por estado
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "InfoICMSporUF") Then
    If gbCreateField("Saídas", "InfoICMSporUF", dbBoolean) Then
      db.Execute "UPDATE Saídas SET InfoICMSporUF = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '02/06/2005 - Daniel
  '
  '253. Inclusão de campo
  '     Tabela     : Funcionários
  '     Campo      : AllowDescProd
  '     Solicitante: Suporte Infopar
  '     Finalidade : Permitir o funcionário conceder desconto para produtos
  '                  que não estejam habilitados para conceder descontos (VR)
  '                  ------[ Detalhe ]------
  '                  Antes da beta 6.52.0.47 o desconto ocorria em VR independente
  '                  se o produto era habilitado ou não para conceder desconto
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "AllowDescProd") Then
    If gbCreateField("Funcionários", "AllowDescProd", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET AllowDescProd = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '06/06/2005 - Daniel
  '
  '254. Novo registro em ZzzProgramas (Relatório de Usuários/Funcionários)
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas12() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '06/06/2005 - Daniel
  '
  '255. Inclusão de campo
  '     Tabela     : Funcionários
  '     Campo      : Ativo
  '     Solicitante: Carlos - OSM Consultoria
  '     Finalidade : Ativar / Desativar funcionários do
  '                  cadastro de funcionários
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "Ativo") Then
    If gbCreateField("Funcionários", "Ativo", dbBoolean) Then
      db.Execute "UPDATE Funcionários SET Ativo = TRUE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '08/06/2005 - Daniel
  '
  '256. Inclusão de campo
  '     Tabela     : Etiquetas - Tempo
  '     Campo      : Funcionario
  '     Solicitante: Miss Nuvem
  '     Finalidade : Através do aplicativo de etiquetas para a térmica
  '                  limpar às etiquetas do funcionário
  nPhase = nPhase + 1
  If Not gbGetField("Etiquetas - Tempo", "Funcionario") Then
    If gbCreateField("Etiquetas - Tempo", "Funcionario", dbInteger) Then
      db.Execute "UPDATE [Etiquetas - Tempo] SET Funcionario = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '15/06/2005 - Daniel
  '
  '257. Inclusão de campo
  '     Tabela     : Cli_For
  '     Campo      : NomeSacadorAvalista
  '     Solicitante: Infopar
  '     Finalidade : Correspondência com o Quick CNAB
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "NomeSacadorAvalista") Then
    If gbCreateField("Cli_For", "NomeSacadorAvalista", dbText, 40) Then
      db.Execute "UPDATE Cli_For SET NomeSacadorAvalista = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '15/06/2005 - Daniel
  '
  '258. Inclusão de campo
  '     Tabela     : Cli_For
  '     Campo      : CPFSacadorAvalista
  '     Solicitante: Infopar
  '     Finalidade : Correspondência com o Quick CNAB
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "CPFSacadorAvalista") Then
    If gbCreateField("Cli_For", "CPFSacadorAvalista", dbText, 20) Then
      db.Execute "UPDATE Cli_For SET CPFSacadorAvalista = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '17/06/2005 - Daniel
  '
  '259. Inclusão de campo
  '     Tabela     : Cli_For
  '     Campo      : CPFCNPJSacadorAvalista
  '     Solicitante: Infopar
  '     Finalidade : Correspondência com o Quick CNAB
  nPhase = nPhase + 1
  If Not gbGetField("Cli_For", "CPFCNPJSacadorAvalista") Then
    If gbCreateField("Cli_For", "CPFCNPJSacadorAvalista", dbText, 4) Then
      db.Execute "UPDATE Cli_For SET CPFCNPJSacadorAvalista = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '20/06/2005 - Daniel
  '
  '260. Criação da Tabela CodigoNBM
  '     Finalidade: Atender os usuários do Nordeste para à correta geração
  '     do arquivo SEF (Registro 75) onde cada produto deverá ter o Código NBM;
  '     Obrigatório para empresas contribuintes do IPI
  nPhase = nPhase + 1
  If gbGetTable("CodigoNBM") = False Then
    If gbCreateTableCodigoNBM() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""CodigoNBM"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '17/06/2005 - Daniel
  '
  '261. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : CodigoNBM
  '     Solicitante: Pneus & Cia (PE)
  '     Finalidade : Informar no registro 75 do arquivo Sintegra/SEF
  '                  o Código NBM para cada Produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "CodigoNBM") Then
    If gbCreateField("Produtos", "CodigoNBM", dbText, 8) Then
      db.Execute "UPDATE Produtos SET CodigoNBM = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '24/06/2005 - Daniel
  '
  '262. Inclusão de índice na tabela de Produtos (Código + CodigoNBM)
  '
  'Solicitação: Pneus & Cia (PE)
  '
  nPhase = nPhase + 1
  If Not g_blnGetIndex("Produtos", "CodigoNBM") Then
    If Not m_blnCreateIndexCodigoNBM() Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de índice na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '27/06/2005 - mpdea
  '263. Alteração na Tabela CliForNumeravel
  '     Alterado o tamanho do campo CodNumer de 15 para 25 - v6.52.60
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("CliForNumeravel", "CodNumer", dbText, 25, "PrimaryKey") Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Clientes / Fornecedores - Bens Numeráveis"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '28/06/2005 - Daniel
  '
  '264. Inclusão de campo
  '     Tabela     : [Operações Saída]
  '     Campo      : AlteraStatusPedidoWeb
  '     Solicitante: Livraria Osório
  '     Finalidade : Alterar status do pedido web para recebido
  '
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "AlteraStatusPedidoWeb") Then
    If gbCreateField("Operações Saída", "AlteraStatusPedidoWeb", dbBoolean) Then
      db.Execute "UPDATE [Operações Saída] SET AlteraStatusPedidoWeb = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '265. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : ConsumoDeTecido
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "ConsumoDeTecido") Then
    If gbCreateField("Produtos", "ConsumoDeTecido", dbDouble) Then
      db.Execute "UPDATE Produtos SET ConsumoDeTecido = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '266. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : PrecoDoMetroLinear
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "PrecoDoMetroLinear") Then
    If gbCreateField("Produtos", "PrecoDoMetroLinear", dbDouble) Then
      db.Execute "UPDATE Produtos SET PrecoDoMetroLinear = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '267. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : CustoDoTecido
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "CustoDoTecido") Then
    If gbCreateField("Produtos", "CustoDoTecido", dbDouble) Then
      db.Execute "UPDATE Produtos SET CustoDoTecido = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '268. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : VlMaoDeObraFaccao
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "VlMaoDeObraFaccao") Then
    If gbCreateField("Produtos", "VlMaoDeObraFaccao", dbDouble) Then
      db.Execute "UPDATE Produtos SET VlMaoDeObraFaccao = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '269. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : VlLavanderia
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "VlLavanderia") Then
    If gbCreateField("Produtos", "VlLavanderia", dbDouble) Then
      db.Execute "UPDATE Produtos SET VlLavanderia = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '270. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : VlBordado
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "VlBordado") Then
    If gbCreateField("Produtos", "VlBordado", dbDouble) Then
      db.Execute "UPDATE Produtos SET VlBordado = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '271. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : VlEstamparia
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "VlEstamparia") Then
    If gbCreateField("Produtos", "VlEstamparia", dbDouble) Then
      db.Execute "UPDATE Produtos SET VlEstamparia = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '272. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : VlAviamentos
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "VlAviamentos") Then
    If gbCreateField("Produtos", "VlAviamentos", dbDouble) Then
      db.Execute "UPDATE Produtos SET VlAviamentos = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '04/07/2005 - Daniel
  '
  '273. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : OutrosCustos
  '     Solicitante: Zue
  '     Finalidade : Identificação do Custo do produto
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "OutrosCustos") Then
    If gbCreateField("Produtos", "OutrosCustos", dbDouble) Then
      db.Execute "UPDATE Produtos SET OutrosCustos = 0;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/07/2005 - Daniel
  '
  '274. Inclusão de campo
  '     Tabela     : [Saídas - Serviços]
  '     Campo      : CST (Código de Situação Tributária)
  '     Solicitante: W.V. Hidroanálise Ltda (J.R. Hidroquímica )
  '     Finalidade : No momento da impressão de nota cada serviço informado na grid de serviços (saídas)
  '                  terá o seu respectivo CST exibido
  nPhase = nPhase + 1
  If Not gbGetField("Saídas - Serviços", "CST") Then
    If gbCreateField("Saídas - Serviços", "CST", dbText, 1) Then
      db.Execute "UPDATE [Saídas - Serviços] SET CST = '';"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas - Serviços"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '01/08/2005 - Daniel
  '
  '275. Inclusão de campo
  '     Tabela     : Saídas
  '     Campo      : DataEmissaoNotaManual
  '     Solicitante: Pedágio Calçados e Confecções
  '     Projeto    : Impressão de Notas Manuais
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "DataEmissaoNotaManual") Then
    If Not gbCreateField("Saídas", "DataEmissaoNotaManual", dbDate) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '08/08/2005 - Daniel
  '
  '276. Novo registro em ZzzProgramas (Configuração de Impressoras)
  '
  nPhase = nPhase + 1
  If AddFileZZZProgramas13() = False Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '09/08/2005 - Daniel
  '
  '277. Inclusão de campo
  '     Tabela     : [Parâmetros Filial]
  '     Campo      : AlterVendedorCliFor
  '     Solicitante: Konrad Comercial Ltda
  '     Finalidade : Apenas o Superusuário pode alterar o campo Vendedor
  '                  na tela de Cli / For
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "AlterVendedorCliFor") Then
    If gbCreateField("Parâmetros Filial", "AlterVendedorCliFor", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET AlterVendedorCliFor = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '15/09/2005 - mpdea
  '
  '278. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : IndicePrecoEntrada
  '     Solicitante: Pavinato
  '     Finalidade : Índice para cálculo do preço do produto na entrada
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "IndicePrecoEntrada") Then
    If gbCreateField("Produtos", "IndicePrecoEntrada", dbDouble) Then
      db.Execute "UPDATE Produtos SET IndicePrecoEntrada = 1;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '22/09/2005 - mpdea
  '
  '279. Inclusão de campo
  '     Tabela     : Operações Entrada
  '     Campo      : GravaCustoPrecoListaSemIPI
  '     Solicitante: Pavinato
  '     Finalidade : Gravar o preço de Custo no campo Preço de Lista sem IPI
  '                  utilizado na pasta Cálculos no Cadastro de Produtos
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "GravaCustoPrecoListaSemIPI") Then
    If Not gbCreateField("Operações Entrada", "GravaCustoPrecoListaSemIPI", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '16/01/2006 - mpdea
  '
  '280. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : VR_Tela_2
  '     Solicitante: Technomax - Rodrigo
  '     Finalidade : Utilização da tela de Venda Rápida em tela cheia
  '                  (estilo CheckOut)
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VR_Tela_CheckOut") Then
    If gbCreateField("Parâmetros Filial", "VR_Tela_CheckOut", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET VR_Tela_CheckOut = FALSE;"
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '25/01/2006 - mpdea
  '
  '281. Inclusão de registro
  '     Tabela     : ZZZProgramas
  '     Finalidade : Permissão para relatório
  '     Solicitante: Cliente Kilouça (QS71271-970)
  If CheckSerialCaseMod("QS71271-970") Then
    nPhase = nPhase + 1
    If Not AddFileZZZProgramas14 Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/01/2006 - mpdea
  '
  '282. Inclusão de tabela
  '     Tabela     : GrupoFiscal
  '     Finalidade : Servir como classificação no cadastro de Produtos e
  '                  Operações de Saídas para a criação das regras de Mensagens
  '                  utilizadas na impressão de Nota Fiscal
  '     Solicitante: Technomax - Rodrigo
  nPhase = nPhase + 1
  If Not gbGetTable("GrupoFiscal") Then
    If Not m_blnCreateTableGrupoFiscal Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""GrupoFiscal"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/01/2006 - mpdea
  '
  '283. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : GrupoFiscal
  '     Finalidade : Informar o Grupo Fiscal do Produto para fins classificatórios,
  '                  conforme item 282 desta função
  '     Solicitante: Technomax - Rodrigo
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "GrupoFiscal") Then
    If Not gbCreateField("Produtos", "GrupoFiscal", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/01/2006 - mpdea
  '
  '284. Inclusão de campo
  '     Tabela     : Operações Saída
  '     Campo      : GrupoFiscal
  '     Finalidade : Informar o Grupo Fiscal da Operação de Saída para fins
  '                  classificatórios, conforme item 282 desta função
  '     Solicitante: Technomax - Rodrigo
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "GrupoFiscal") Then
    If Not gbCreateField("Operações Saída", "GrupoFiscal", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações de Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/01/2006 - mpdea
  '
  '285. Inclusão de registro
  '     Tabela     : ZZZProgramas
  '     Finalidade : Permissões de acesso, conforme item 282 desta função
  '     Solicitante: Technomax - Rodrigo
  nPhase = nPhase + 1
  If Not AddFileZZZProgramas15 Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  
  '26/01/2006 - mpdea
  '
  '286. Inclusão de tabela
  '     Tabela     : MensagensNotaFiscal
  '     Finalidade : Cadastrar mensagens a serem utilizadas na impressão de
  '                  Nota Fiscal de acordo com as regras estipuladas
  '     Solicitante: Technomax - Rodrigo
  nPhase = nPhase + 1
  If Not gbGetTable("MensagensNotaFiscal") Then
    If Not m_blnCreateTableMensagensNotaFiscal Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Mensagens para Nota Fiscal"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '16/05/2006 - mpdea
  '
  '287. Inclusão de campo
  '     Tabela     : Operações Entrada
  '     Campo      : SomarFreteCustoProduto
  '     Finalidade : Informar se soma frete no custo dos produtos
  '     Solicitante: PSI TI Via Brasil - Pavinato
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "SomarFreteCustoProduto") Then
    If Not gbCreateField("Operações Entrada", "SomarFreteCustoProduto", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Operações de Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '16/05/2006 - mpdea
  '
  '288. Inclusão de campo
  '     Tabela     : Produtos
  '     Campo      : IndiceIcmsRetido
  '     Finalidade : Índice para aplicação no cálculo da base de ICMS Retido
  '     Solicitante: PSI TI Via Brasil - Pavinato
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "IndiceIcmsRetido") Then
    If gbCreateField("Produtos", "IndiceIcmsRetido", dbDouble) Then
      db.Execute "UPDATE Produtos SET IndiceIcmsRetido = 1;"
    Else
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '16/05/2006 - mpdea
  '
  '289. Inclusão de campo
  '     Tabela     : Entradas - Produtos
  '     Campo      : ValorIcmsRetido
  '     Finalidade : Valor do imposto de ICMS Retido a ser pago
  '     Solicitante: PSI TI Via Brasil - Pavinato
  nPhase = nPhase + 1
  If Not gbGetField("Entradas - Produtos", "ValorIcmsRetido") Then
    If Not gbCreateField("Entradas - Produtos", "ValorIcmsRetido", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '03/07/2006 - mpdea
  '
  '290. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : ImprimeTicketMovEfetivada
  '     Finalidade : Somente imprimir o ticket para movimentações efetivadas
  '     Solicitante: Bem me quer
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ImprimeTicketMovEfetivada") Then
    If Not gbCreateField("Parâmetros Filial", "ImprimeTicketMovEfetivada", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros da Empresa/Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '19/07/2006 - Andrea
  '
  '291. Inclusão de registro
  '     Tabela     : ZZZProgramas
  '     Finalidade : Inclusão de novo relatório
  '     Solicitante: Be Star (Marisol)
  nPhase = nPhase + 1
  If Not AddFileZZZProgramas16 Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""ZZZProgramas"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '24/07/2006 - Andrea
  '
  '292. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : ExigeSenhaGerReimpTicket
  '     Finalidade : Exigir senha do gerente no caso de reimpressão de ticket
  '     Solicitante: Rodrigo - TechnoMax
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ExigeSenhaGerReimpTicket") Then
    If Not gbCreateField("Parâmetros Filial", "ExigeSenhaGerReimpTicket", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros da Empresa/Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '24/07/2006 - Andrea
  '
  '293. Inclusão de campo
  '     Tabela     : Saídas
  '     Campo      : Ticket Impresso
  '     Finalidade : Ficar marcado quando o ticket for impresso, para que na reimpressao
  '                  seja solicitado a senha do gerente, para clientes que usam o
  '                  ticket como controle da saída de mercadorias.
  '     Solicitante: Rodrigo - TechnoMax
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Ticket Impresso") Then
    If Not gbCreateField("Saídas", "Ticket Impresso", dbBoolean) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '28/07/2006 - Andrea
  '
  '294. Inclusão de campo
  '     Tabela     : Parâmetros Filial
  '     Campo      : NumeroUltMapaECF
  '     Finalidade : Armazenar informação número do último mapa ECF
  '     Solicitante: EBS
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "NumeroUltMapaECF") Then
    If gbCreateField("Parâmetros Filial", "NumeroUltMapaECF", dbInteger) Then
      db.Execute "UPDATE [Parâmetros Filial] SET NumeroUltMapaECF = 0;"
    Else
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros da Empresa/Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '17/11/2006 - Anderson
  '295. Inclusão na Tabela de Operações de Entrada
  '     Tabela     : Operações Entrada
  '     Campo      : BaseICMSFrete
  '     Finalidade : Informar se é para ser considerado o valor do frete no cálculo do ICMS
  '     Solicitante: Technomax
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "BaseICMSFrete") Then
    If Not gbCreateField("Operações Entrada", "BaseICMSFrete", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações de Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '20/11/2006 - Anderson
  '
  '296. Considerar saldo anterior ao movimentar o caixa
  '     Tabela     : Parâmetros Filial
  '     Campo      : ConsiderarSaldoAnterior
  '     Finalidade : Identifica se deve ser considerado o saldo anterior ao movimentar o caixa
  '     Solicitante: Technomax
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ConsiderarSaldoAnterior") Then
    If Not gbCreateField("Parâmetros Filial", "ConsiderarSaldoAnterior", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
'06/12/2006 - Anderson
'O campo foi retirado da tabela por haver uma alteração na estrutura do projeto
'
'  '01/12/2006 - Anderson
'  '
'  '297. Campo para cadastro de CFOP por produto
'  '     Tabela     : Produtos
'  '     Campo      : Código Fiscal
'  '     Finalidade : Cadastrar o CFOP por produto para a emissão da nota fiscal e a geração de sintegra
'  '     Solicitante: Technomax
'  nPhase = nPhase + 1
'  If Not gbGetField("Produtos", "Código Fiscal") Then
'    If Not gbCreateField("Produtos", "Código Fiscal", dbText, 14) Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'  End If
  
  '17/11/2006 - Anderson
  '298. Inclusão na Tabela de Operações de Entrada
  '     Tabela     : Operações Entrada
  '     Campo      : ICMSSobreIPI
  '     Finalidade : Considerar cálculo do IPI sobre o ICMS
  '     Solicitante: Technomax
  nPhase = nPhase + 1
  If Not gbGetField("Operações Entrada", "ICMSSobreIPI") Then
    If Not gbCreateField("Operações Entrada", "ICMSSobreIPI", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações de Entrada"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '01/12/2006 - Anderson
  '
  '299. Campo registro de CFOP do produto no momento da venda
  '     Tabela     : Produtos
  '     Campo      : CFOP
  '     Finalidade : registrar o CFOP por produto para a emissão da nota fiscal e a geração de sintegra
  '     Solicitante: Technomax
  nPhase = nPhase + 1
  If Not gbGetField("Saídas - Produtos", "CFOP") Then
    If Not gbCreateField("Saídas - Produtos", "CFOP", dbText, 14) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas - Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '01/12/2006 - Anderson
  '
  '300. Campo registro de CFOP do serviço no momento da venda
  '     Tabela     : Saídas - Serviços
  '     Campo      : CFOP
  '     Finalidade : registrar o CFOP por serviço para a emissão da nota fiscal e a geração de sintegra
  '     Solicitante: Technomax
  nPhase = nPhase + 1
  If Not gbGetField("Saídas - Serviços", "CFOP") Then
    If Not gbCreateField("Saídas - Serviços", "CFOP", dbText, 14) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas - Serviços"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '08/12/2006 - Anderson
  '301. Criação da tabela para o registro do CFOP por produto
  nPhase = nPhase + 1
  If gbGetTable("ProdutoCFOP") = False Then
    If gbCreateTableProdutoCFOP() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ProdutoCFOP"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '15/12/2006 - Anderson
  '302. Criação da tabela para o registro do CFOP por serviço
  nPhase = nPhase + 1
  If gbGetTable("ServicoCFOP") = False Then
    If gbCreateTableServicoCFOP() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ServicoCFPO"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '18/12/2006 - Anderson
  '303. Parametro para a exibição da coluna do CFOP
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "ExibeCFOP") Then
    If Not gbCreateField("Parâmetros Filial", "ExibeCFOP", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '17/01/2006 - Anderson
  '
  '304. Solicitar senha do gerente ao alterar vendedor nas telas de clientes, venda, venda rápida e saídas.
  '     Tabela     : Parâmetros Filial
  '     Campo      : VendedorSenhaGerente
  '     Finalidade : Evitar que o vendedor seja alterado ao efetuar uma venda ou ao cadastrar um cliente
  '     Solicitante: Flávio da SMQ
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VendedorSenhaGerente") Then
    If Not gbCreateField("Parâmetros Filial", "VendedorSenhaGerente", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '26/01/2007 - Anderson
  '
  '305. Registrar os dados de recebimento da conta do cliente
  '     Tabela     : ContaClienteRecebimento
  '     Finalidade : Efetuar pagamentos de contas de clientes através dos recusos de recebimento da tela de saída
  '     Solicitante: Rodrigo - Technomax
  nPhase = nPhase + 1
  If gbGetTable("ContaClienteRecebimento") = False Then
    If m_blnCreateTableContaClienteRecebimento() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ContaClienteRecebimento"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '08/03/2007 - Anderson
  '
  '306. Alteração do ID do programa que abra o relatório de funcionário
  '     Tabela     : ZZZProgramas
  '     Finalidade : Corrigir Bug de permissões no Quick Store
  '     Solicitante: Rodrigo - Technomax
  nPhase = nPhase + 1
  db.Execute "UPDATE [ZZZProgramas] SET ToolID = 300712 WHERE Número=170", dbFailOnError
  
  '17/04/2007 - Anderson
  '
  '307. Criação do campo utilizado para dividir o preço do produto na impressão da etiqueta, atendendo as exigências do Procon
  '     Tabela     : Produtos
  '     Finalidade : Dividir o preço do produto na impressão da etiqueta
  '     Solicitante: A. M. DE FARIA E CASTRO CIA LTDA (QS38380-938)
  nPhase = nPhase + 1
  If gbGetField("Produtos", "DividirPrecoEtiqueta") = False Then
    If gbCreateField("Produtos", "DividirPrecoEtiqueta", dbInteger, 0) Then
      db.Execute "UPDATE Produtos SET DividirPrecoEtiqueta = 1;", dbFailOnError
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na Tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '19/04/2007 - Anderson
  '
  '308. Criação do campo utilizado para dividir o preço do produto na impressão da etiqueta, atendendo as exigências do Procon
  '     Tabela     : Etiquetas - Tempo
  '     Finalidade : Dividir o preço do produto na impressão da etiqueta
  '     Solicitante: A. M. DE FARIA E CASTRO CIA LTDA (QS38380-938)
  nPhase = nPhase + 1
  If gbGetField("Etiquetas - Tempo", "DividirPrecoEtiqueta") = False Then
    If Not gbCreateField("Etiquetas - Tempo", "DividirPrecoEtiqueta", dbInteger, 0) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na Tabela ""Etiquetas - Tempo"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/05/2007 - Anderson
  '
  '309. Indica se o Quick Store deve manter as observações impressas na última Nota Fiscal
  '     Tabela     : Parâmetros Filial
  '     Campo      : MantemInformacaoUltimaNotaFiscal
  '     Finalidade : Evitar que a nota fiscal seja impressa com os mesmos dados da nota fiscal anterior.
  '     Solicitante: Diego Technomax
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "MantemInformacaoUltimaNotaFiscal") Then
    If gbCreateField("Parâmetros Filial", "MantemInformacaoUltimaNotaFiscal", dbBoolean) Then
      db.Execute "UPDATE [Parâmetros Filial] SET MantemInformacaoUltimaNotaFiscal = -1;", dbFailOnError
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Parâmetros Filial"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '16/05/2007 - Anderson
  '
  '310. Criado para informar o dígito verificador do nosso número para transações com o CNAB
  '     Tabela     : Contas a Receber
  '     Campo      : CNAB_DigitoVerificador
  '     Finalidade : Armazenar o nosso número com o dígito verificador para boletos pré-impressos
  '     Solicitante: Technomax - Cliente Agrotama (QS73073-894)
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CNAB_DigitoVerificador") Then
    If Not gbCreateField("Contas a Receber", "CNAB_DigitoVerificador", dbText, 3) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '16/05/2007 - Anderson
  '
  '311. Criado para informar a carteira selecionada para boletos pré-impressos do CNAB
  '     Tabela     : Contas a Receber
  '     Campo      : CNAB_DigitoVerificador
  '     Finalidade : Armazenar a carteira selecionada para boletos pré-impressos do CNAB
  '     Solicitante: Technomax - Cliente Agrotama (QS73073-894)
  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "CNAB_Carteira") Then
    If Not gbCreateField("Contas a Receber", "CNAB_Carteira", dbText, 3) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Contas a Receber"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '16/05/2007 - Anderson
  '
  '312. Informar as carteiras disponíveis para utilização do CNAB
  '     Tabela     : CNABCarteira
  '     Finalidade : Informar as carteiras disponíveis para utilização do CNAB
  '     Solicitante: Technomax - Cliente Agrotama (QS73073-894)
  nPhase = nPhase + 1
  If gbGetTable("CNABCarteira") = False Then
    If gbCreateTableCNABCarteira() Then
      db.Execute "INSERT INTO CNABCarteira Values('9','Banco Bradesco')", dbFailOnError
    Else
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""CNABCarteira"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '15/05/2013 - Alexandre
  '313. Criação da Tabela SaidasComandas
  'Esta tabela foi criada para atender inicialmente a necessidade
  'da DiskEmbalagens
  nPhase = nPhase + 1
  If gbGetTable("SaidasComandas") = False Then
    If gbCreateTableSaidasComandas() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""SaidasComandas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
    
  '26/06/2013 - Alexandre
  '314. Criação da Tabela AliquotasNCM
  'Esta tabela foi criada para atender a lei De Olho noe Impostos
  Dim ZZZ As Recordset
  Set ZZZ = db.OpenRecordset("Select * from ZZZ")
  Dim versao As Long
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao >= 70127 And versao <= 70155 Then
    'db.Execute ("Drop Table AliquotasNCM")
    db.TableDefs.Delete "AliquotasNCM"
  End If
  nPhase = nPhase + 1
  If gbGetTable("AliquotasNCM") = False Then
    If gbCreateTableAliquotasNCM() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""AliquotasNCM"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/06/2013 - Alexandre Afornali
  '315. Criação do campo AliqNCM tabela Produtos para lei De olho nos impostos
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "AliqNCM") Then
    If Not gbCreateField("Produtos", "AliqNCM", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  '26/06/2013 - Alexandre Afornali
  '316. Criação do campo TotalNCM tabela Saídas para lei De olho nos impostos
'  nPhase = nPhase + 1
'  If Not gbGetField("Saídas", "TotalNCM") Then
'    If Not gbCreateField("Saídas", "TotalNCM", dbDouble) Then
'      ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'  End If
  
  '26/06/2013 - Alexandre Afornali
  '317. Criação do campo TotalNCM tabela Entradas para lei De olho nos impostos
'  nPhase = nPhase + 1
'  If Not gbGetField("Entradas", "TotalNCM") Then
'    If Not gbCreateField("Entradas", "TotalNCM", dbDouble) Then
'      ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'  End If
  
  '14/08/2014 - Jean
  '318. Criação do campo nrEvento tabela NFe para nova maneira de cancelar por evento
  
End Function
Private Function AlteraDB3(ByRef nPhase As Integer)
  Dim intX As Integer
  
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "TotalNCM") Then
    If Not gbCreateField("Saídas", "TotalNCM", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Entradas", "TotalNCM") Then
    If Not gbCreateField("Entradas", "TotalNCM", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '13/11/2014 - Eduardo Franco
  '1. Criação do campo Desp_Acessorias tabela Saídas para informar as despesas acessórias de cada produto
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "Total_Desp_Acessorias") Then
    If Not gbCreateField("Saídas", "Total_Desp_Acessorias", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '13/11/2014 - Eduardo Franco
  '2. Criação do campo Desp_Acessorias tabela Saídas para informar as despesas acessórias de cada produto
  nPhase = nPhase + 1
  If Not gbGetField("Saídas - Produtos", "Desp_Acessorias") Then
    If Not gbCreateField("Saídas - Produtos", "Desp_Acessorias", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  Dim ZZZ As Recordset
  Set ZZZ = db.OpenRecordset("Select * from ZZZ")
  Dim versao As Long
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70169 Then
    nPhase = nPhase + 1
    'If gbGetTableTemp("tblRelVendasPorVendedor") = False Then
      If gbCreateTableTBLRelVendasPorVendedor() = False Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição da tabela ""tblRelVendasPorFornecedor"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    'End If
  End If
  
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas", "Consumidor_Final") Then
      If Not gbCreateField("Saídas", "Consumidor_Final", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas", "Presenca_Comprador") Then
      If Not gbCreateField("Saídas", "Presenca_Comprador", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Cli_For", "IndicadorIE") Then
      If Not gbCreateField("Cli_For", "IndicadorIE", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Cli_For"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas - Produtos", "ValorICMSDesonerado") Then
      If Not gbCreateField("Saídas - Produtos", "ValorICMSDesonerado", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas - Produtos", "MotivoDesoneracaoICMS") Then
      If Not gbCreateField("Saídas - Produtos", "MotivoDesoneracaoICMS", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas - Produtos", "Percentual_Diferimento") Then
      If Not gbCreateField("Saídas - Produtos", "Percentual_Diferimento", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas - Produtos", "Valor_Aprox_Impostos") Then
      If Not gbCreateField("Saídas - Produtos", "Valor_Aprox_Impostos", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas", "TotalDesoneracaoICMS") Then
      If Not gbCreateField("Saídas", "TotalDesoneracaoICMS", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas", "Consumidor_Final") Then
      If Not gbCreateField("Entradas", "Consumidor_Final", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas", "Presenca_Comprador") Then
      If Not gbCreateField("Entradas", "Presenca_Comprador", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas", "TotalDesoneracaoICMS") Then
      If Not gbCreateField("Entradas", "TotalDesoneracaoICMS", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas - Produtos", "ValorICMSDesonerado") Then
      If Not gbCreateField("Entradas - Produtos", "ValorICMSDesonerado", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas - Produtos", "MotivoDesoneracaoICMS") Then
      If Not gbCreateField("Entradas - Produtos", "MotivoDesoneracaoICMS", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas - Produtos", "Valor_Aprox_Impostos") Then
      If Not gbCreateField("Entradas - Produtos", "Valor_Aprox_Impostos", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas - Produtos", "Percentual_Diferimento") Then
      If Not gbCreateField("Entradas - Produtos", "Percentual_Diferimento", dbDouble) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas - Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70175 Then
  nPhase = nPhase + 1
    If Not gbGetField("Produtos", "MotivoDesoneracaoICMS") Then
      If Not gbCreateField("Produtos", "MotivoDesoneracaoICMS", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70184 Then
    nPhase = nPhase + 1
    If Not gbGetField("NFeRetorno", "StatusDescricao2") Then
      If Not gbCreateField("NFeRetorno", "StatusDescricao2", dbText, 255) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""NFeRetorno"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  If versao <= 70184 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas", "FinalidadeNFe") Then
      If Not gbCreateField("Entradas", "FinalidadeNFe", dbInteger) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  '353
  If versao <= 70184 Then
  nPhase = nPhase + 1
    If Not gbGetField("Entradas", "ChaveReferenciada") Then
      If Not gbCreateField("Entradas", "ChaveReferenciada", dbText, 100) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
 'Michel
 If versao <= 70341 Then
  nPhase = nPhase + 1
    If Not gbGetField("Saídas", "aliquota_origem") Then
      If Not gbCreateField("Saídas", "aliquota_origem", dbText, 100) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    If Not gbGetField("Saídas", "aliquota_destino") Then
      If Not gbCreateField("Saídas", "aliquota_destino", dbText, 100) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If
  
  '357
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo("NFeRetorno", "DigestValue", dbText, 255) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""NFeRetorno"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '358
  If versao <= 70211 Then
    nPhase = nPhase + 1
    If Not gbAlteraTipoCampo("NFeRetorno", "Status", "Long") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""NFeRetorno"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '359
  If versao <= 70213 Then
    nPhase = nPhase + 1
    If Not gbGetField("Parâmetros Filial", "UltimaNFCe") Then
      If Not gbCreateField("Parâmetros Filial", "UltimaNFCe", dbLong) Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      Else
        db.Execute "UPDATE [Parâmetros Filial] SET UltimaNFCe = 0", dbFailOnError
      End If
    End If
  End If
  
  '360
  nPhase = nPhase + 1
  If Not gbGetField("NFe", "nrEvento") Then
    If Not gbCreateField("NFe", "nrEvento", dbInteger) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFe"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '361
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NFCe") Then
    If Not gbCreateField("Saídas", "NFCe", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '362
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "TotalCartaoDebito") Then
    If Not gbCreateField("Saídas", "TotalCartaoDebito", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '363
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "TotalCartaoCredito") Then
    If Not gbCreateField("Saídas", "TotalCartaoCredito", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '364
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "TotalCartaoCredito") Then
    If Not gbCreateField("Saídas", "TotalCartaoCredito", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '365
  nPhase = nPhase + 1
  If Not gbGetTable("Movimento - Cartoes") Then
    If Not gbCreateTableMovimentoCartoes Then
        ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Movimento - Cartoes"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
    End If
  End If
  
  '366
  nPhase = nPhase + 1
  If Not gbGetField("Movimento - Cartoes", "Credito") Then
    If Not gbCreateField("Movimento - Cartoes", "Credito", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Movimento - Cartoes"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '367
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "PermiteMostrarCliente") Then
    If Not gbCreateField("Operações Saída", "PermiteMostrarCliente", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '368
  nPhase = nPhase + 1
  If gbGetTable("NFCE_ENVI") = False Then
    If gbCreateTableNFCE_ENVI() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""NFCE_ENVI"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '369
  nPhase = nPhase + 1
  If gbGetTable("NFCE_job") = False Then
    If gbCreateTableNFCE_job() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""NFCE_job"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '371
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "ChaveNFCe") Then
    If Not gbCreateField("Saídas", "ChaveNFCe", dbText) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '372
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "CPF_CPNJ_Cliente") Then
    If Not gbCreateField("Saídas", "CPF_CPNJ_Cliente", dbText) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '373
  nPhase = nPhase + 1
  If gbGetTable("Cupom_temp") = False Then
    If gbCreateTableCupom_temp() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Cupom_temp"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '374
  nPhase = nPhase + 1
  If Not gbGetField("NFCE_job", "Processado") Then
    If Not gbCreateField("NFCE_job", "Processado", dbText, , , , , "N") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  Call AlteraDB4(nPhase)
  
End Function
Private Function AlteraDB4(ByRef nPhase As Integer)
  '375
  nPhase = nPhase + 1
    If Not gbGetField("NFCE_job", "CPF") Then
      If Not gbCreateField("NFCE_job", "CPF", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '376
    nPhase = nPhase + 1
    If Not gbGetField("NFCE_job", "Nome_Consumidor") Then
      If Not gbCreateField("NFCE_job", "Nome_Consumidor", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '377
    nPhase = nPhase + 1
    If Not gbGetField("NFCE_job", "Data_Emissao") Then
      If Not gbCreateField("NFCE_job", "Data_Emissao", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '378
    nPhase = nPhase + 1
    If Not gbGetField("NFCE_job", "Total_Tributos") Then
      If Not gbCreateField("NFCE_job", "Total_Tributos", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '379
    nPhase = nPhase + 1
    If Not gbGetField("NFCE_job", "Nome_Emitente") Then
      If Not gbCreateField("NFCE_job", "Nome_Emitente", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '380
    nPhase = nPhase + 1
    If Not gbGetField("NFCE_job", "Endereco_Emitente") Then
      If Not gbCreateField("NFCE_job", "Endereco_Emitente", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '381
    nPhase = nPhase + 1
    If Not gbGetField("NFCE_job", "IE_Emitente") Then
      If Not gbCreateField("NFCE_job", "IE_Emitente", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
  '382
  Dim ZZZ As Recordset
  Set ZZZ = db.OpenRecordset("Select * from ZZZ")
  Dim versao As Long
  versao = Replace(ZZZ("DBVersion"), ".", "")
  If versao <= 70270 Then
    nPhase = nPhase + 1
    If Not gbAlteraTipoCampo("NFCE_ENVI", "URL_QRCode", "Memo") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""NFCE_ENVI"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '383
  nPhase = nPhase + 1
    If Not gbGetField("Saídas", "Emitiu_Dados_Cliente_NFCe") Then
      If Not gbCreateField("Saídas", "Emitiu_Dados_Cliente_NFCe", dbBoolean, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '384
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "Emitiu_Somente_CPF_Cliente_NFCe") Then
      If Not gbCreateField("Saídas", "Emitiu_Somente_CPF_Cliente_NFCe", dbBoolean, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '385
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "Nome_Cliente_NFCe") Then
      If Not gbCreateField("Saídas", "Nome_Cliente_NFCe", dbText, , True, , , "") Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""NFCE_job"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  '386
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo("Saídas - Produtos", "Situação Tributária", dbText, 255) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""NFeRetorno"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If
  
  '387
  nPhase = nPhase + 1
  If gbGetTable("Ref_CEST_NCM") = False Then
    If gbCreateTableRef_CEST_NCM() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""Ref_CEST_NCM"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '388
  nPhase = nPhase + 1
  If gbGetTable("AliquotasNCM") = False Then
    If gbCreateTableAliquotasNCM() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""AliquotasNCM"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  Else
    If gbAlterTableAliquotasNCM() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Alteração da tabela ""AliquotasNCM"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  ' VERSÃO
  ' PILATTI/MAURO Novembro/17
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "FreteSomaOuNaoEstimativa") Then
    If Not gbCreateField("Saídas", "FreteSomaOuNaoEstimativa", dbBoolean, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Saídas][columa FreteSomaOuNaoEstimativa] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  ' Fevereiro/19
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "TipoSituacaoTributariaPIS") Then
    If Not gbCreateField("Parâmetros Filial", "TipoSituacaoTributariaPIS", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Parâmetros Filial][columa TipoSituacaoTributariaPIS] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "TipoSituacaoTributariaPIS") Then
    If Not gbCreateField("Produtos", "TipoSituacaoTributariaPIS", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Produtos][columa TipoSituacaoTributariaPIS] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  'Março/2019
  'Campos para NFCe - tratamento de contingencia
  
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "retNFCe") Then
    If Not gbCreateField("Saídas", "retNFCe", dbMemo, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Saídas][columa retNFCe] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NFCe_contingencia_num") Then
    If Not gbCreateField("Saídas", "NFCe_contingencia_num", dbLong, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Saídas][columa NFCe_contingencia_num] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NFCe_contingencia_serie") Then
    If Not gbCreateField("Saídas", "NFCe_contingencia_serie", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Saídas][columa NFCe_contingencia_serie] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NFCe_contingencia_status") Then
    If Not gbCreateField("Saídas", "NFCe_contingencia_status", dbText, 30, True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Saídas][columa NFCe_contingencia_status] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "retNFCe_contingencia") Then
    If Not gbCreateField("Saídas", "retNFCe_contingencia", dbMemo, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Saídas][columa retNFCe_contingencia] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "NFCe_contingencia_chave") Then
    If Not gbCreateField("Saídas", "NFCe_contingencia_chave", dbText, 50, True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Saídas][columa NFCe_contingencia_chave] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If gbGetTable("AcessoTabelasDePrecosProdutos") = False Then
    If gbCreateTableAcessoTabelasDePrecosProdutos() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""AcessoTabelasDePrecosProdutos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If gbGetTable("ProdutoPareamentoFornecedor") = False Then
    If gbCreateTableProdutoPareamentoFornecedor() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ProdutoPareamentoFornecedor"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If gbGetTable("TransferenciaEntreFiliais") = False Then
    If gbCreateTableTransferenciaEntreFiliais() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""TransferenciaEntreFiliais"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If gbGetTable("TransferenciaProdutos") = False Then
    If gbCreateTableTransferenciaProdutos() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""TransferenciaProdutos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If gbGetTable("DRE_quick") = False Then
    If gbCreateTableDRE_quick() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""DRE_quick"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If gbGetTable("DRE_anexos") = False Then
    If gbCreateTableDRE_anexos() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""DRE_anexos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If gbGetTable("SaidasChaves") = False Then
    If gbCreateTableSaidasChaves() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""SaidasChaves"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If Not gbGetField("Parâmetros Filial", "CobrarMultaAposVencimentoParcela") Then
    If Not gbCreateField("Parâmetros Filial", "CobrarMultaAposVencimentoParcela", dbBoolean, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Parâmetros Filial][columa CobrarMultaAposVencimentoParcela] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If Not gbGetField("Parâmetros Filial", "TaxaMultaParcelaVencida") Then
    If Not gbCreateField("Parâmetros Filial", "TaxaMultaParcelaVencida", dbSingle, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Parâmetros Filial][columa TaxaMultaParcelaVencida] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  If Not gbGetField("Parâmetros Filial", "MultaDiasAposParcelaVencida") Then
    If Not gbCreateField("Parâmetros Filial", "MultaDiasAposParcelaVencida", dbSingle, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Parâmetros Filial][columa MultaDiasAposParcelaVencida] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  ' ***
  ' Insert de registros na tabela DRE_anexos
  Dim rsDRE_anexos As Recordset
 
  Set rsDRE_anexos = db.OpenRecordset("Select * from DRE_Anexos", dbOpenDynaset)
  
  If rsDRE_anexos.EOF And rsDRE_anexos.BOF Then
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(1, 'Empresas do Comércio. Estão incluídas bares, restaurante e lojas em geral', 0, 180000, 4, 0)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(1, 'Empresas do Comércio. Estão incluídas bares, restaurante e lojas em geral', 180000.01, 360000, 7.3, 5940.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(1, 'Empresas do Comércio. Estão incluídas bares, restaurante e lojas em geral', 360000.01, 720000, 9.5, 13860.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(1, 'Empresas do Comércio. Estão incluídas bares, restaurante e lojas em geral', 720000.01, 1800000, 10.7, 22500.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(1, 'Empresas do Comércio. Estão incluídas bares, restaurante e lojas em geral', 1800000.01, 3600000, 14.3, 87300.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(1, 'Empresas do Comércio. Estão incluídas bares, restaurante e lojas em geral', 3600000.01, 4200000, 19, 378000.00)"
  
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(2, 'Indústria. Estão incluídas empresas industriais e fábricas.', 0, 180000, 4.5, 0)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(2, 'Indústria. Estão incluídas empresas industriais e fábricas.', 180000.01, 360000, 7.8, 5940.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(2, 'Indústria. Estão incluídas empresas industriais e fábricas.', 360000.01, 720000, 10, 13860.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(2, 'Indústria. Estão incluídas empresas industriais e fábricas.', 720000.01, 1800000, 11.2, 22500.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(2, 'Indústria. Estão incluídas empresas industriais e fábricas.', 1800000.01, 3600000, 14.7, 85000.00)"
      db.Execute "Insert into DRE_anexos (CodigoAnexo, Obs, ValorDe, ValorAte, Aliquota, ValorRedutor) values(2, 'Indústria. Estão incluídas empresas industriais e fábricas.', 3600000.01, 4200000, 30, 720000.00)"
  
  End If
  rsDRE_anexos.Close
  Set rsDRE_anexos = Nothing
  
  
  
  nPhase = nPhase + 1
  If Not gbGetField("SaidasComandas", "Filial") Then
    If Not gbCreateField("SaidasComandas", "Filial", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [SaidasComandas][columa Filial] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  ' ***
  ' Insert de registros na tabela de ZZZProgramas
  Dim rsZZZProgramas As Recordset
  Dim sZZZProgramas As String
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 184 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('PROGRAMA FIDELIDADE PARÂMETROS','Programa Fidelidade Parâmetros',184,40022) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 185 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('PROGRAMA FIDELIDADE OPER SAIDA','Programa Fidelidade x Operações Saída',185,40023) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 186 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('PROGRAMA FIDELIDADE CONSULTA GERENCIAL','Programa Fidelidade x Consultas Gerenciais',186,40024) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 187 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('PROGRAMA FIDELIDADE RESGATE PONTOS','Programa Fidelidade x Resgate Pontos',187,40025) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 188 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('PROGRAMA FIDELIDADE CLI ENTREGA RESGATE','Programa Fidelidade x Cliente entrega Resgate',188,40026) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 189 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('PROGRAMA FIDELIDADE CLIENTES NÃO PART','Programa Fidelidade x Clientes que não particupam',189,40027) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 190 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('PROGRAMA FIDELIDADE CNPJ GRUPOS','Programa Fidelidade x CNPJs do Grupo Fidelidade',190,40028) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 191 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('CADASTRO PRODUTO VINCULA CFOP','Cadastro Produto x CFOPs vinculados',191,40021) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 192 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('RELATÓRIO ESTRATÉGICO MAIORES PRODUTOS','Relatório Estratégico x Maiores produtos',192,304440) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 193 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('RELATÓRIO ESTRATÉGICO MAIORES CLIENTES','Relatório Estratégico x Maiores clientes',193,304470) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 194 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('RELATÓRIO ESTRATÉGICO MAIORES FORNECEDOR','Relatório Estratégico x Maiores fornecedores',194,304450) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
'  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 195 "
'  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
'
'  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
'      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
'      sZZZProgramas = sZZZProgramas & " Values ('VENDA RÁPIDA (SOMENTE ESTA TELA)','Venda Rápida (Somente esta tela)',195,0) "
'      db.Execute sZZZProgramas
'  End If
'  rsZZZProgramas.Close
'  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 196 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('TELA CONFIGURAÇÃO DE IMPRESSORAS','Tela Configuração de Impressoras',196,1207) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 197 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('CADASTRO DE NCM','Cadastro de NCM',197,1328) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 198 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('DEVOLUÇÕES','Devoluções/Troca de Produtos',198,50041) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  sZZZProgramas = "Select Número From ZZZProgramas Where Número = 199 "
  Set rsZZZProgramas = db.OpenRecordset(sZZZProgramas, dbOpenDynaset)
  
  If rsZZZProgramas.EOF And rsZZZProgramas.BOF Then
      sZZZProgramas = "Insert into ZZZProgramas ([Nome Programa], Descrição, Número, ToolID) "
      sZZZProgramas = sZZZProgramas & " Values ('RELATÓRIO SAÍDAS E ENTRADAS','Relatório de Saídas e Entradas',199,301103) "
      db.Execute sZZZProgramas
  End If
  rsZZZProgramas.Close
  Set rsZZZProgramas = Nothing
  
  ' ***
  
  If Not gbGetField("Funcionários", "bMostrarTelaPesquisaProdutoTipoFoto") Then
      If Not gbCreateField("Funcionários", "bMostrarTelaPesquisaProdutoTipoFoto", dbBoolean) Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
      End If
  End If
  
  If Not gbGetField("Funcionários", "bUsuarioAcessoApenasTelaVendaRapida") Then
      If Not gbCreateField("Funcionários", "bUsuarioAcessoApenasTelaVendaRapida", dbBoolean) Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Funcionários"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
      End If
  End If
  
  ' Se não existe...cria tabela de Cesta de Produtos
  If gbGetTable("ProdutoCesta") = False Then
    If gbCreateTableProdutoCesta() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ProdutoCesta"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  ' ***
  ' Insert de registro na tabela de [Operações Entrada]
  Dim rsEntradaQuick_DevolucaoProd As Recordset
  Dim sEntradaQuick As String
  
  sEntradaQuick = "Select Código From [Operações Entrada] Where Código = -1 "
  Set rsEntradaQuick_DevolucaoProd = db.OpenRecordset(sEntradaQuick, dbOpenDynaset)
  
  If rsEntradaQuick_DevolucaoProd.EOF And rsEntradaQuick_DevolucaoProd.BOF Then
      sEntradaQuick = "Insert into [Operações Entrada] (Código, Nome, Tipo, Estoque) "
      sEntradaQuick = sEntradaQuick & " Values (-1,'Devolução produto pelo cliente por troca','D',1) "
      db.Execute sEntradaQuick
  End If
  rsEntradaQuick_DevolucaoProd.Close
  Set rsEntradaQuick_DevolucaoProd = Nothing
  
  
  sEntradaQuick = "Select Código From [Operações Entrada] Where Código = -2 "
  Set rsEntradaQuick_DevolucaoProd = db.OpenRecordset(sEntradaQuick, dbOpenDynaset)
  
  If rsEntradaQuick_DevolucaoProd.EOF And rsEntradaQuick_DevolucaoProd.BOF Then
      sEntradaQuick = "Insert into [Operações Entrada] (Código, Nome, Tipo, Estoque, Comissão) "
      sEntradaQuick = sEntradaQuick & " Values (-2,'Devolução prod. c/comissão p/cliente por troca','D',1, 1) "
      db.Execute sEntradaQuick
  End If
  rsEntradaQuick_DevolucaoProd.Close
  Set rsEntradaQuick_DevolucaoProd = Nothing
  
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Quick_viaRDP") Then
    If Not gbCreateField("Parâmetros Filial", "Quick_viaRDP", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Parâmetros Filial][columa Quick_viaRDP] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetTable("CodigoBeneficio") = False Then
    If gbCreateTableCodigoBeneficio() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""CodigoBeneficio"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "CodigoBeneficio") Then
    If Not gbCreateField("Produtos", "CodigoBeneficio", dbText, 10, True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Produtos][columa CodigoBeneficio] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "SituacaoTributariaEntrada") Then
    If Not gbCreateField("Produtos", "SituacaoTributariaEntrada", dbText, 4, True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Produtos][columa SituacaoTributariaEntrada] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  nPhase = nPhase + 1
  If gbGetTable("ProdutoFavoritos") = False Then
    If gbCreateTableProdutoFavoritos() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""ProdutoFavoritos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  sZZZProgramas = "Update [Parâmetros Filial] set [VR Permite Rec Rápido]=0, DescSubTotalRateado=0 "
  db.Execute sZZZProgramas
  
  nPhase = nPhase + 1
  If Not gbGetField("NFeInutilizadas", "Modelo") Then
    If Not gbCreateField("NFeInutilizadas", "Modelo", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [NFeInutilizadas][columa Modelo] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
'  nPhase = nPhase + 1
'  If Not gbGetField("Produtos", "IPI_ValidoEntradaSaida") Then
'    If Not gbCreateField("Produtos", "IPI_ValidoEntradaSaida", dbInteger) Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'  End If
  
  nPhase = nPhase + 1
  If gbGetField("Produtos", "IPI_ValidoEntradaSaida") Then
    If Not gbDeleteField("Produtos", "IPI_ValidoEntradaSaida") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs1") Then
    If Not gbDeleteField("Saídas", "Obs_Obs1") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs2") Then
    If Not gbDeleteField("Saídas", "Obs_Obs2") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs3") Then
    If Not gbDeleteField("Saídas", "Obs_Obs3") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs4") Then
    If Not gbDeleteField("Saídas", "Obs_Obs4") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs5") Then
    If Not gbDeleteField("Saídas", "Obs_Obs5") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs6") Then
    If Not gbDeleteField("Saídas", "Obs_Obs6") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs7") Then
    If Not gbDeleteField("Saídas", "Obs_Obs7") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  nPhase = nPhase + 1
  If gbGetField("Saídas", "Obs_Obs8") Then
    If Not gbDeleteField("Saídas", "Obs_Obs8") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Saídas"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  
  nPhase = nPhase + 1
  If Not gbGetField("AliquotasNCM", "CEST") Then
      If Not gbCreateField("AliquotasNCM", "CEST", dbText, 10, True, , , "") Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""AliquotasNCM"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
      End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("AliquotasNCM", "TemFCP") Then
      If Not gbCreateField("AliquotasNCM", "TemFCP", dbBoolean, , True, , , "") Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""AliquotasNCM"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
      End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "Percentual_IPI_Entrada") Then
      If Not gbCreateField("Produtos", "Percentual_IPI_Entrada", dbSingle) Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
      End If
  End If
  
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "BaseCalculoICMSST_Saida") Then
    If Not gbCreateField("Produtos", "BaseCalculoICMSST_Saida", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
 
  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "BaseCalculoICMSST_Entrada") Then
    If Not gbCreateField("Produtos", "BaseCalculoICMSST_Entrada", dbDouble) Then
      ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Produtos"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "Percentual_ICMSST_Entrada") Then
      If Not gbCreateField("Produtos", "Percentual_ICMSST_Entrada", dbSingle) Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
      End If
  End If

  nPhase = nPhase + 1
  If Not gbGetField("Produtos", "Percentual_ICMSST_Saida") Then
      If Not gbCreateField("Produtos", "Percentual_ICMSST_Saida", dbSingle) Then
          Call ws.Rollback
          Screen.MousePointer = vbDefault
          gnStyle = vbOKOnly + vbCritical
          gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
          gsTitle = "Atenção"
          Call MsgBox(gsMsg, gnStyle, gsTitle)
          db.Close
          ws.Close
          End
      End If
  End If
  
'''  nPhase = nPhase + 1
'''  If Not gbGetField("Cli_For", "Pendencia") Then
'''    If Not gbCreateField("Cli_For", "Pendencia", dbBoolean, , True, , , "") Then
'''      Call ws.Rollback
'''      Screen.MousePointer = vbDefault
'''      gnStyle = vbOKOnly + vbCritical
'''      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Cli_For][columa Pendencia] - não foi possível."
'''      gsTitle = "Atenção"
'''      Call MsgBox(gsMsg, gnStyle, gsTitle)
'''      db.Close
'''      ws.Close
'''      End
'''    End If
'''  End If

  nPhase = nPhase + 1
  If gbGetField("Cli_For", "Pendencia") Then
    If Not gbDeleteField("Cli_For", "Pendencia") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Exclusão de campo na tabela ""Cli_For"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  nPhase = nPhase + 1
  If Not gbGetField("Contas a Receber", "Pendencia") Then
    If Not gbCreateField("Contas a Receber", "Pendencia", dbBoolean, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Contas a Receber][columa Pendencia] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "ObterTributosProduto_EntradaOuSaida") Then
    If Not gbCreateField("Operações Saída", "ObterTributosProduto_EntradaOuSaida", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Operações Saída][columa ObterTributosProduto_EntradaOuSaida] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
 
 
 
  '389
  '09/02/2017 Jean
  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Saída", "Saldo Anterior", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '390
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Saída", "Vendas Cliente", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '391
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Saída", "Devolução", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '392
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Saída", "Novo Empréstimo", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '393
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Saída", "Saldo Atual", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '394
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Saída", "Preço Unitário", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '395
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Entrada", "Saldo Anterior", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '396
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Entrada", "Vendas", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '397
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Entrada", "Devolução", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '398
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Entrada", "Empréstimo Recebido", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '399
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Entrada", "Saldo Atual", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
'
'  '400
'  '09/02/2017 Jean
'  'Alterações feitas para poder gravar quantidade fracionada nas tabelas de consignação
'  nPhase = nPhase + 1
'  If Not gbAlteraTipoCampo("Consignação Entrada", "Preço Unitário", "Double") Then
'      Call ws.Rollback
'      Screen.MousePointer = vbDefault
'      gnStyle = vbOKOnly + vbCritical
'      gsMsg = "Manutenção na Base de Dados: Adição de registro na tabela ""Consignação Saída"" - não foi possível."
'      gsTitle = "Atenção"
'      Call MsgBox(gsMsg, gnStyle, gsTitle)
'      db.Close
'      ws.Close
'      End
'    End If
  
End Function

'Utilizar esta função a partir de 2023
Private Function AlteraDB5(ByRef nPhase As Integer)


  '12/02/2023 - Pablo
  'Criado campo que diz a operação de entrada para transferência
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Transf_OpEntrada") Then
    If Not gbCreateField("Parâmetros Filial", "Transf_OpEntrada", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial (Transf_OpEntrada)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '12/02/2023 - Pablo
  'Criado campo que diz a operação de saída para transferência
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Transf_OpSaida") Then
    If Not gbCreateField("Parâmetros Filial", "Transf_OpSaida", dbLong) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial (Transf_OpSaida)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '12/02/2023 - Pablo
  'Criado campo que diz a tabela de preços para transferência
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Transf_TabelaPrecos") Then
    If Not gbCreateField("Parâmetros Filial", "Transf_TabelaPrecos", dbText, 15) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial (Transf_TabelaPrecos)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If
  
  '02/03/2023 - Pablo
  'Criado campo que permite ou não imprimir ticket via RDP
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "Quick_viaRDP_ticket") Then
    If Not gbCreateField("Parâmetros Filial", "Quick_viaRDP_ticket", dbInteger, , True, , , "") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela [Parâmetros Filial][columa Quick_viaRDP_ticket] - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      If load_Quick_viaRDP_ticket = False Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Adição de valor do campo na tabela [Parâmetros Filial][columa Quick_viaRDP_ticket] - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  End If

  '09/03/2023 - Pablo
  '
  '     Parâmetro de operações de saídas
  '     do número de documento (CPF ou CNPJ)
  '     Tabela     : Operações Saída
  '     Finalidade : Somar IPI ao total da nota
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "SomaIpiTotalNota") Then
    If Not gbCreateField("Operações Saída", "SomaIpiTotalNota", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" SomaIpiTotalNota - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE [Operações Saída] SET SomaIpiTotalNota = 0 ", dbFailOnError
    End If
  End If

  '14/03/2023 - Pablo
  '
  '     Parâmetro de operações de saídas
  '     do número de documento (CPF ou CNPJ)
  '     Tabela     : Operações Saída
  '     Finalidade : Somar IPI ao total da base de cálculo do ICMS
  nPhase = nPhase + 1
  If Not gbGetField("Operações Saída", "SomaIpiTotalBC") Then
    If Not gbCreateField("Operações Saída", "SomaIpiTotalBC", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações Saída"" SomaIpiTotalBC - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    Else
      db.Execute "UPDATE [Operações Saída] SET SomaIpiTotalBC = 0 ", dbFailOnError
    End If
  End If
  
  '31/10/2023 - Pablo
  '
  '     Tabela     : SaidasComandas
  '     Coluna     : Filial
  '     Finalidade : Alterar tipo da coluna para poder vincular FK
  nPhase = nPhase + 1
  If Not gbAlteraTipoCampo("SaidasComandas", "Filial", "Byte") Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração da tabela ""SaidasComandas"" - não foi possível alterar o tipo da coluna ""Filial""."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If

  '31/10/2023 - Pablo
  '
  '     Tabela     : SaidasComandas
  '     Coluna     : CodSaida
  '     Finalidade : Alterar tipo da coluna para poder vincular FK
  nPhase = nPhase + 1
  If Not gbAlteraTipoCampo("SaidasComandas", "CodSaida", "Long") Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração da tabela ""SaidasComandas"" - não foi possível alterar o tipo da coluna ""CodSaida""."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If

  '22/11/2023 - Pablo
  '
  '     Tabela     : Saidas
  '     Coluna     : Referência
  '     Finalidade : Alterar tamanho da coluna para 20 caracteres
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo("Saídas", "Referência", dbText, 20) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração da tabela ""Saidas"" - não foi possível alterar o tamanho da coluna ""Referência""."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If

  '07/05/2024 - Pablo
  'Criado campo que pergunta se oculta ou não os orçamentos da tela de pedido rápido
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "VR_OcultaOrc") Then
    If Not gbCreateField("Parâmetros Filial", "VR_OcultaOrc", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial (VR_OcultaOrc)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '07/05/2024 - Pablo
  'Criado campo que permite setar o prestador de serviço na venda rápida
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "comPrestServ") Then
    If Not gbCreateField("Parâmetros Filial", "comPrestServ", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial (comPrestServ)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '07/05/2024 - Pablo
  'Criado campo que marca usuário como prestador de serviços
  nPhase = nPhase + 1
  If Not gbGetField("Funcionários", "isPrestServ") Then
    If Not gbCreateField("Funcionários", "isPrestServ", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Funcionários (isPrestServ)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '07/05/2024 - Pablo
  'Criado campo que adiciona prestador de serviços para comissão
  nPhase = nPhase + 1
  If Not gbGetField("Saídas", "PrestadorServico") Then
    If Not gbCreateField("Saídas", "PrestadorServico", dbInteger, 0, True, False, True, "NÃO PONHA ZERO") Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas (PrestadorServico)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If


  '13/04/2025 - Pablo
  'Cria a tabela ProdutoNomeNFe caso o nome do produto precise ser alterado para a nota fiscal
  nPhase = nPhase + 1
  If gbGetTable("ProdutoNomeNFe") = False Then
    If p_blnCreateTableProdutoNomeNFe() = False Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Criação da tabela ""ProdutoNomeNFe"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '13/04/2025 - Pablo
  'Criado campo que permite editar o nome do produto
  nPhase = nPhase + 1
  If Not gbGetField("Parâmetros Filial", "EditarNomeProduto") Then
    If Not gbCreateField("Parâmetros Filial", "EditarNomeProduto", dbBoolean) Then
      Call ws.Rollback
      Screen.MousePointer = vbDefault
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Parâmetros Filial (EditarNomeProduto)"" - não foi possível."
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      db.Close
      ws.Close
      End
    End If
  End If

  '28/04/2025 - Pablo
  'Aumentando o tamanho do campo do número do cartão
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Saídas", "Recebe - Num Cartão", dbText, 30) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Saídas (Recebe - Num Cartão)"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If

  '29/04/2025 - Pablo
  'Aumentando o tamanho do campo do número do cartão
  nPhase = nPhase + 1
  If Not gbAlteraTamanhoCampo2("Movimento - Cartoes", "NumeroCartao", dbText, 30) Then
    Call ws.Rollback
    Screen.MousePointer = vbDefault
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Movimento - Cartoes (NumeroCartao)"" - não foi possível."
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    db.Close
    ws.Close
    End
  End If


End Function

'02/03/2023 - Pablo
'Preenche valores padrão para o Parâmetro Quick_viaRDP_ticket
Private Function load_Quick_viaRDP_ticket() As Boolean
  Dim rs As Recordset
  On Error GoTo ErrHandle
  Set rs = db.OpenRecordset("Parâmetros Filial", dbOpenDynaset)
  With rs
    Do While Not .EOF
      .Edit
      .Fields("Quick_viaRDP_ticket") = .Fields("Quick_viaRDP")
      .Update
      .MoveNext
    Loop
  End With
  rs.Close
  Set rs = Nothing
  load_Quick_viaRDP_ticket = True
  Exit Function

ErrHandle:
  load_Quick_viaRDP_ticket = False
  Exit Function
End Function

Private Function gbLoadValorP() As Boolean
  Dim rs As Recordset
  On Error GoTo ErrHandle
  Set rs = db.OpenRecordset("Funcionários", dbOpenDynaset)
  With rs
    Do While Not .EOF
      .Edit
      .Fields("ValorP") = CStr(CriptografaSenha(.Fields("Senha").Value))
      .Fields("Senha") = Format(Date, "yyyymmdd")
      .Update
      .MoveNext
    Loop
  End With
  rs.Close
  Set rs = Nothing
  gbLoadValorP = True
  Exit Function

ErrHandle:
  gbLoadValorP = False
  Exit Function
End Function

Private Sub CreateFieldsOBS(ByRef nPhase As Integer)
  '---[ 11/08/2003 - maikel - Criação dos campos de observações na tabela saídas]---'
  
    If Not gbGetField("Saídas", "obs_infCpl1") Then
      If Not gbCreateField("Saídas", "obs_infCpl1", dbText, 255) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
    If Not gbGetField("Saídas", "obs_infCpl2") Then
      If Not gbCreateField("Saídas", "obs_infCpl2", dbText, 255) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
    If Not gbGetField("Entradas", "obs_infCpl1") Then
      If Not gbCreateField("Entradas", "obs_infCpl1", dbText, 255) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
    If Not gbGetField("Entradas", "obs_infCpl2") Then
      If Not gbCreateField("Entradas", "obs_infCpl2", dbText, 255) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  
  
'    '85. Observação 1
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs1") Then
'      If Not gbCreateField("Saídas", "obs_Obs1", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
'
'    '86. Observação 2
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs2") Then
'      If Not gbCreateField("Saídas", "obs_Obs2", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
'
'    '87. Observação 3
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs3") Then
'      If Not gbCreateField("Saídas", "obs_Obs3", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
'
'    '88. Observação 4
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs4") Then
'      If Not gbCreateField("Saídas", "obs_Obs4", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
'
'    '89. Observação 5
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs5") Then
'      If Not gbCreateField("Saídas", "obs_Obs5", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
'
'    '90. Observação 6
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs6") Then
'      If Not gbCreateField("Saídas", "obs_Obs6", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
'
'    '91. Observação 7
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs7") Then
'      If Not gbCreateField("Saídas", "obs_Obs7", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
'
'    '92. Observação 8
'    nPhase = nPhase + 1
'    If Not gbGetField("Saídas", "obs_Obs8") Then
'      If Not gbCreateField("Saídas", "obs_Obs8", dbText, 30) Then
'        Call ws.Rollback
'        Screen.MousePointer = vbDefault
'        gnStyle = vbOKOnly + vbCritical
'        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
'        gsTitle = "Atenção"
'        Call MsgBox(gsMsg, gnStyle, gsTitle)
'        db.Close
'        ws.Close
'        End
'      End If
'    End If
    
    '93. Transportadora
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_Transportadora") Then
      If Not gbCreateField("Saídas", "obs_Transportadora", dbText, 50) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '94. Placa
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_Placa") Then
      If Not gbCreateField("Saídas", "obs_Placa", dbText, 8) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '95. UF
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_Uf") Then
      If Not gbCreateField("Saídas", "obs_Uf", dbText, 2) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '96. Qtde
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_Qtde") Then
      If Not gbCreateField("Saídas", "obs_Qtde", dbText, 10) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '97. Especie
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_Especie") Then
      If Not gbCreateField("Saídas", "obs_Especie", dbText, 10) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '98. Marca
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_Marca") Then
      If Not gbCreateField("Saídas", "obs_Marca", dbText, 10) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '99. Peso Liquido
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_PesoLiquido") Then
      If Not gbCreateField("Saídas", "obs_PesoLiquido", dbDouble) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '100. Qtde
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_PesoBruto") Then
      If Not gbCreateField("Saídas", "obs_PesoBruto", dbDouble) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '101. Frete pago por quem ( 1 - Emitente, 2 - Destinatário )
    nPhase = nPhase + 1
    If Not gbGetField("Saídas", "obs_FretePago") Then
      If Not gbCreateField("Saídas", "obs_FretePago", dbByte) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Saídas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  '---[ Criação dos campos de observações na tabela saídas]---'
  
  
  
  
  
  
  '---[ 11/08/2003 - maikel - Criação dos campos de observações na tabela Entradas]---'
    '85. Observação 1
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs1") Then
      If Not gbCreateField("Entradas", "obs_Obs1", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '86. Observação 2
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs2") Then
      If Not gbCreateField("Entradas", "obs_Obs2", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '87. Observação 3
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs3") Then
      If Not gbCreateField("Entradas", "obs_Obs3", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '88. Observação 4
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs4") Then
      If Not gbCreateField("Entradas", "obs_Obs4", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '89. Observação 5
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs5") Then
      If Not gbCreateField("Entradas", "obs_Obs5", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '90. Observação 6
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs6") Then
      If Not gbCreateField("Entradas", "obs_Obs6", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '91. Observação 7
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs7") Then
      If Not gbCreateField("Entradas", "obs_Obs7", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '92. Observação 8
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Obs8") Then
      If Not gbCreateField("Entradas", "obs_Obs8", dbText, 30) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '93. Transportadora
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Transportadora") Then
      If Not gbCreateField("Entradas", "obs_Transportadora", dbText, 50) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '94. Placa
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Placa") Then
      If Not gbCreateField("Entradas", "obs_Placa", dbText, 8) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '95. UF
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Uf") Then
      If Not gbCreateField("Entradas", "obs_Uf", dbText, 2) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '96. Qtde
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Qtde") Then
      If Not gbCreateField("Entradas", "obs_Qtde", dbText, 10) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '97. Especie
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Especie") Then
      If Not gbCreateField("Entradas", "obs_Especie", dbText, 10) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '98. Marca
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_Marca") Then
      If Not gbCreateField("Entradas", "obs_Marca", dbText, 10) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '99. Peso Liquido
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_PesoLiquido") Then
      If Not gbCreateField("Entradas", "obs_PesoLiquido", dbDouble) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '100. Qtde
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_PesoBruto") Then
      If Not gbCreateField("Entradas", "obs_PesoBruto", dbDouble) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
    
    '101. Frete pago por quem ( 1 - Emitente, 2 - Destinatário )
    nPhase = nPhase + 1
    If Not gbGetField("Entradas", "obs_FretePago") Then
      If Not gbCreateField("Entradas", "obs_FretePago", dbByte) Then
        Call ws.Rollback
        Screen.MousePointer = vbDefault
        gnStyle = vbOKOnly + vbCritical
        gsMsg = "Manutenção na Base de Dados: Inclusão de campo na tabela ""Entradas"" - não foi possível."
        gsTitle = "Atenção"
        Call MsgBox(gsMsg, gnStyle, gsTitle)
        db.Close
        ws.Close
        End
      End If
    End If
  '---[ Criação dos campos de observações na tabela entradas]---'
End Sub

Private Function gbCreateIndexFieldCodigosAcesso() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
 On Error Resume Next
  
  Set td = db.TableDefs("Acessos")
  
  td.Indexes.Delete "Programa"
  td.Indexes.Refresh
      
  Err.Clear
  
  Set iX = td.CreateIndex
  With iX
    .Name = "Programa"
    .Fields.Append .CreateField("Numero")
    .Fields.Append .CreateField("Usuário")
    .Primary = True
    .Unique = False
  End With
  td.Indexes.Append iX

  ' Refresh collection so that you can access new Index objects.
  td.Indexes.Refresh

  Set iX = Nothing
  Set td = Nothing
  
End Function

'10/12/2009 - Andrea
Private Function gbCreateIndexFieldMovimentoCartoes() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
 On Error Resume Next
  
  Set td = db.TableDefs("[Movimento - Cartoes]")
  
  td.Indexes.Delete "Ordem"
  td.Indexes.Refresh
      
  Err.Clear
  
  Set iX = td.CreateIndex("Ordem")
  With iX
    .Name = "Ordem"
    .Fields.Append .CreateField("Filial")
    .Fields.Append .CreateField("Sequência")
    .Fields.Append .CreateField("Ordem")
    .Primary = True
    .Unique = True
  End With
  td.Indexes.Append iX

  ' Refresh collection so that you can access new Index objects.
  td.Indexes.Refresh

  Set iX = Nothing
  Set td = Nothing
  
End Function


Private Function gbCreateTableReports() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  Dim nI As Integer
  Dim sSql As String
  Dim rs As Recordset
  
  On Error GoTo ErrCreate
  
  'Função alterada para o tratamento de cores ao relatório zebrado por mpdea
  
  'Exclui a tabela ParametrosTMP do banco de dados
  If gbGetTable("ParametrosTMP") Then
    Call db.Execute("DROP TABLE ParametrosTMP", dbFailOnError)
  End If
'  Call dbTemp.Execute("DROP TABLE ParametrosTMP", dbFailOnError)
  
  Set td = db.CreateTableDef("Reports")
  
  Set fd = td.CreateField("InRelZebrados", dbBoolean)
  td.Fields.Append fd
  Set fd = td.CreateField("nColorRed", dbByte)
  td.Fields.Append fd
  Set fd = td.CreateField("nColorGreen", dbByte)
  td.Fields.Append fd
  Set fd = td.CreateField("nColorBlue", dbByte)
  td.Fields.Append fd
  
  db.TableDefs.Append td
  
  Set td = Nothing
  
  'Grava o valor padrão
  Set rs = db.OpenRecordset("Reports", dbOpenDynaset)
  With rs
    .AddNew
    !InRelZebrados = True
    'Amarelo palha
    !nColorRed = 255
    !nColorGreen = 255
    !nColorBlue = 174
    .Update
    .Close
  End With
  Set rs = Nothing
  
  gbCreateTableReports = True
  Exit Function
  
ErrCreate:
  gbCreateTableReports = False
  
End Function

Private Function gbCreateTableCliForCaract() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("CliForCaract")
  
  Set fd = td.CreateField("CodCliCaract", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("TipoCliCaract", dbText, 1)
  td.Fields.Append fd
  Set fd = td.CreateField("CodCaract", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("ValCaract", dbText, 30)
  fd.AllowZeroLength = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("CodCliCaract")
  iX.Fields.Append iX.CreateField("CodCaract")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableCliForCaract = True
  Exit Function
  
ErrCreate:
  gbCreateTableCliForCaract = False
  
End Function

Private Function gbCreateTableGruposDeClientes() As Boolean
  '07/07/2004 - Daniel
  'Tabela inteligente com informações de grupos de clientes
  'Case: TV Shopping e liberado apenas para a TV Shopping
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("GruposDeClientes")
  
  'Filial
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  'Nome dos Grupos de 1 a 4
  Set fd = td.CreateField("NomeG1", dbText, 40)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("NomeG2", dbText, 40)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("NomeG3", dbText, 40)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("NomeG4", dbText, 40)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  'Limite Inicial do Grupo 1 a 4
  Set fd = td.CreateField("LimiteIniG1", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("LimiteIniG2", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("LimiteIniG3", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("LimiteIniG4", dbDouble)
  td.Fields.Append fd
  
  'Limite Final do Grupo 1 a 3 pois o limite final do grupo 4 será infinito
  Set fd = td.CreateField("LimiteFinG1", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("LimiteFinG2", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("LimiteFinG3", dbDouble)
  td.Fields.Append fd
  
  'Código para cada grupo do 1 ao 4
  Set fd = td.CreateField("CodigoG1", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("CodigoG2", dbByte)
  td.Fields.Append fd
    
  Set fd = td.CreateField("CodigoG3", dbByte)
  td.Fields.Append fd
    
  Set fd = td.CreateField("CodigoG4", dbByte)
  td.Fields.Append fd
    
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Filial")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableGruposDeClientes = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableGruposDeClientes = False

End Function

Private Function gbCreateTableDiferimento() As Boolean
'14/05/2004 - Daniel
'Tabela criada para armazenar as informações sobre o Diferimento
'Case: Embalavi mas aberto para todos.
'Diferimento: Se o Cliente é do Estado do PR por exemplo e trata-se de Pessoa Jurídica
'então é startado um cálculo especial de diferimento
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Diferimento")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Total", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Base", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("EstadoCorrente", dbText, 2)
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("ObsDiferimento", dbText, 70)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Filial")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableDiferimento = True
  Exit Function

ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableDiferimento = False

End Function

Private Function gbCreateTablePrestacaoContas() As Boolean
  'Function criada em 17/09/2004
  'Finalidade...: Armazenar as [Entradas - Produtos] que poderão ser acertadas ou editadas
  'Case.........: Resultado
  'Criada por...: Daniel
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("PrestacaoContas")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Fornecedor", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Sequencia", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Linha", dbByte)
  td.Fields.Append fd
    
  Set fd = td.CreateField("Produto", dbText, 20)
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Custo", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("QtdeOriginal", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("QtdeDevolvida", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("QtdeVendida", dbDouble)
  td.Fields.Append fd
    
  Set fd = td.CreateField("QtdeComprada", dbDouble)
  td.Fields.Append fd
    
  Set fd = td.CreateField("DatadaGeracao", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Finalizado", dbBoolean)
  td.Fields.Append fd
  
  Set fd = td.CreateField("DatadaFinalizacao", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("ImpressoNF", dbBoolean)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Resultado", dbByte)
  td.Fields.Append fd
  
    
  'Set iX = td.CreateIndex("PrimaryKey")
  'iX.Fields.Append iX.CreateField("Filial")
  'iX.Fields.Append iX.CreateField("Sequencia")
  'iX.Fields.Append iX.CreateField("Linha")
  'iX.Primary = True
  'iX.Unique = True
  'td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTablePrestacaoContas = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTablePrestacaoContas = False

End Function

Private Function gbCreateTableParamDevoMat() As Boolean
  'Function criada em 15/09/2004
  'Finalidade...: Valores padrões para cada saída gerada na Devolução de Materiais
  'Case.........: Resultado
  'Criada por...: Daniel
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ParamDevoMat")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Operacao", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Caixa", dbByte)
  td.Fields.Append fd
    
  Set fd = td.CreateField("Tabela", dbText, 15)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Filial")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableParamDevoMat = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableParamDevoMat = False

End Function

Private Function gbCreateTableRetencao() As Boolean
  '21/03/2005 - Daniel
  'Case: Enxovais Bem Me Quer
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Retencao")
  
  Set fd = td.CreateField("Código", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome", dbText, 50)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  Set fd = td.CreateField("NomeDaFinanceira", dbText, 10)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  Set fd = td.CreateField("ValorRetencao", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Tipo", dbText, 16)
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Código")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableRetencao = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableRetencao = False
  
End Function

Private Function gbCreateTableCodigoNBM() As Boolean
  '20/06/2005 - Daniel
  'Solicitante: Pneus & Cia (PE)
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("CodigoNBM")
  
  Set fd = td.CreateField("Código", dbText, 8)
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome", dbText, 100)
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Código")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableCodigoNBM = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableCodigoNBM = False
  
End Function

Private Function gbCreateTableBooksVendidos() As Boolean
'13/12/2004 - Daniel
'Case: Livraria Resultado
  Dim td As TableDef
  Dim fd As Field
  'Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("BooksVendidos")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Sequencia", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Linha", dbByte)
  td.Fields.Append fd
    
'  Set iX = td.CreateIndex("PrimaryKey")
'  iX.Fields.Append iX.CreateField("Código")
'  iX.Primary = True
'  iX.Unique = True
'  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableBooksVendidos = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableBooksVendidos = False

End Function

Private Function gbCreateTableTipoComercial() As Boolean
'31/03/2004 - Daniel
'Case: STC (Sistema Trídio de Comunicação - Caxias do Sul)
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("TipoComercial")
  
  Set fd = td.CreateField("Código", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Descricao", dbText, 50)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Código")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableTipoComercial = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableTipoComercial = False

End Function

Private Function gbCreateTableRadio() As Boolean
'31/03/2004 - Daniel
'Case: STC (Sistema Trídio de Comunicação - Caxias do Sul)
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Radio")
  
  Set fd = td.CreateField("Código", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome", dbText, 50)
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Endereco", dbText, 50)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cidade", dbText, 30)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Estado", dbText, 2)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("CNPJ", dbText, 20)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Inscricao", dbText, 20)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Telefone", dbText, 20)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Contatos", dbText, 40)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Código")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableRadio = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableRadio = False

End Function

Private Function gbCreateTableProgramacao() As Boolean
'06/04/2004 - Daniel
'Case: STC (Sistema Trídio de Comunicação - Caxias do Sul)
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Programacao")
  
  Set fd = td.CreateField("Num Autorizacao", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("MesX", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Programacao", dbText, 25)
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 01", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 02", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 03", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 04", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 05", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 06", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 07", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 08", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 09", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 10", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 11", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 12", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 13", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 14", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 15", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 16", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 17", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 18", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 19", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 20", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 21", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 22", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 23", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 24", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 25", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 26", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 27", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 28", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 29", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 30", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dia 31", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Total de Insercoes", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor Unitario", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor Total", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Periodo Ini", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Periodo Fin", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Faixa Ini", dbText, 7)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Faixa Fin", dbText, 7)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Frequencia", dbText, 3)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Duracao", dbText, 5)
  fd.AllowZeroLength = True
  td.Fields.Append fd
    
  Set fd = td.CreateField("Mes", dbText, 3)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Condicoes Pagamento", dbText, 30)
  fd.AllowZeroLength = True
  td.Fields.Append fd
    
  Set fd = td.CreateField("Gerar Etiqueta", dbBoolean)
  fd.Required = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cancela Contrato", dbBoolean)
  fd.Required = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Faturado", dbBoolean)
  fd.Required = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor1", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor2", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor3", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor4", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Vencimento1", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Vencimento2", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Vencimento3", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Vencimento4", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Status1", dbBoolean)
  fd.Required = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Status2", dbBoolean)
  fd.Required = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Status3", dbBoolean)
  fd.Required = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Status4", dbBoolean)
  fd.Required = False
  td.Fields.Append fd
  
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Num Autorizacao")
  iX.Fields.Append iX.CreateField("MesX")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableProgramacao = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableProgramacao = False

End Function

Private Function gbCreateTableContrato() As Boolean
'20/01/2004 - Daniel
'Case: STC (Sistema Trídio de Comunicação - Caxias do Sul)
'Tabela reescrita em 02/04/2004
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Contrato")
  
  Set fd = td.CreateField("Num Autorizacao", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cod Cliente", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cod Radio", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cod Fornecedor", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Patrocinio", dbMemo, 1200)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Observacoes", dbMemo, 1200)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cod TipoComercial", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Data Assinatura", dbDate)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cod Vendedor", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Comissao", dbDouble)
  td.Fields.Append fd
    
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Num Autorizacao")
  iX.Fields.Append iX.CreateField("Cod Cliente")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableContrato = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableContrato = False
End Function

Private Function gbCreateTableParamFaturameAuto() As Boolean
  'Function criada em 02/08/2004
  'Finalidade: Atender as necessidades de faturamento automático da STC de Caxias do Sul
  'Criada por: Daniel
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ParamFaturameAuto")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Operacao", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Servico", dbInteger)
  td.Fields.Append fd
    
  Set fd = td.CreateField("Caixa", dbByte)
  td.Fields.Append fd
    
  Set fd = td.CreateField("Tabela", dbText, 15)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  Set fd = td.CreateField("ISS", dbDouble)
  td.Fields.Append fd
    
    
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Filial")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableParamFaturameAuto = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableParamFaturameAuto = False

End Function

Private Function gbCreateTableSupervisores() As Boolean
  'Function criada em 29/07/2004
  'Por: Daniel
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Supervisores")
  
  Set fd = td.CreateField("Código", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome", dbText, 50)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  Set fd = td.CreateField("Obs", dbMemo, 1200)
  fd.AllowZeroLength = True
  td.Fields.Append fd
    
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Código")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableSupervisores = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableSupervisores = False

End Function

Private Function gbCreateTableAcertoConsignacaoEntrada() As Boolean
  'Function criada em 29/07/2004
  'Por: Daniel
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("AcertoConsignacaoEntrada")
  
  'Codigo da filial
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  'Numero da sequencia de entrada
  Set fd = td.CreateField("Sequencia", dbLong)
  td.Fields.Append fd
  
  'Data da movimentacao de saida que gera o acerto de emprestimo
  Set fd = td.CreateField("DataAcerto", dbDate)
  td.Fields.Append fd
  
  'Linha do produto na movimentação
  Set fd = td.CreateField("LinhaProduto", dbLong)
  td.Fields.Append fd
  
  'Código do produto
  Set fd = td.CreateField("CodigoProduto", dbText, 100)
  td.Fields.Append fd
    
  'Qtde Vendida
  Set fd = td.CreateField("QtdeVendida", dbDouble)
  td.Fields.Append fd
  
  'Filial Venda
  Set fd = td.CreateField("FilialVenda", dbByte)
  td.Fields.Append fd
  
  'Sequência Venda
  Set fd = td.CreateField("SequenciaVenda", dbLong)
  td.Fields.Append fd
    
    
'  Set iX = td.CreateIndex("PrimaryKey")
'  iX.Fields.Append iX.CreateField("Código")
'  iX.Primary = True
'  iX.Unique = True
'  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableAcertoConsignacaoEntrada = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbCreateTableAcertoConsignacaoEntrada = False

End Function
Private Function gbAlterTableCliForCaract() As Boolean
  Dim rs As Recordset
  Dim td As TableDef
  Dim fd As Field
  Dim bGotValor As Boolean
  Dim sField As String
  
  Set td = db.TableDefs("CliForCaract")
  
  On Error Resume Next
  sField = td.Fields("ValorCaract").Name
  If Err.Number <> 0 Then
    bGotValor = False
  Else
    bGotValor = True
  End If
  
  Err.Clear
  
  On Error GoTo ErrAlter
  
  '17/08/2007 - Anderson
  'Alteração realizada por causa de uma alteração de tamanho de campo para 255 na versão 6.55.85
  'Set fd = td.CreateField("Val2Caract", dbText, 30)
  Set fd = td.CreateField("Val2Caract", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set td = Nothing
  
  Set rs = db.OpenRecordset("CliForCaract")
  If Not rs.EOF Then
    Do While Not rs.EOF
      rs.Edit
      If bGotValor = True Then
        rs("Val2Caract").Value = rs("ValorCaract").Value & ""
      Else
        rs("Val2Caract").Value = rs("ValCaract").Value & ""
      End If
      rs.Update
      rs.MoveNext
    Loop
  End If
  
  rs.Close
  Set rs = Nothing
  
  Set td = db.TableDefs("CliForCaract")
  If bGotValor = True Then
    td.Fields.Delete "ValorCaract"
  Else
    td.Fields.Delete "ValCaract"
  End If
  td.Fields("Val2Caract").Name = "ValCaract"
  Set td = Nothing
  
  gbAlterTableCliForCaract = True
  Exit Function
  
ErrAlter:
  gbAlterTableCliForCaract = False
  
End Function

Private Function gbAlteraClassificacaoFiscal(ByVal sTable As String) As Boolean
'13/05/2010 - Andrea
  
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  Dim rs As Recordset
  Dim rsTemp As Recordset
    
  On Error GoTo ErrCreate
    
  Set td = db.TableDefs(sTable)

  If td("Código").Type = dbInteger Then
    gbAlteraClassificacaoFiscal = True
    Set td = Nothing
    Exit Function
  End If

  'Cria uma tabela temporária
  Set td = db.CreateTableDef("tmpClassificacaoFiscal")
  
  Set fd = td.CreateField("Código", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome", dbText, 15)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("Código")
  iX.Fields.Append iX.CreateField("Código")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  
  'Copia os dados da tabela origem para a tabela temporária
  db.Execute "INSERT INTO tmpClassificacaoFiscal SELECT * FROM [Classificação Fiscal]", dbFailOnError
  
  'Exclui a tabela origem
  db.Execute "DROP TABLE [Classificação Fiscal]", dbFailOnError
  
 
  'Cria novamente a Tabela de Classificação Fiscal
  Set td = db.CreateTableDef("Classificação Fiscal")
  
  Set fd = td.CreateField("Código", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome", dbText, 15)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  'Índice para Codigo
  Set iX = td.CreateIndex("Código")
  iX.Fields.Append iX.CreateField("Código")
  iX.Unique = True
  iX.Primary = True
  td.Indexes.Append iX
  
  'Índice para Nome
  Set iX = td.CreateIndex("Nome")
  iX.Fields.Append iX.CreateField("Nome")
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
    
  'Copia os dados da tabela temporaria para a tabela temporária
  db.Execute "INSERT INTO [Classificação Fiscal] SELECT * FROM [tmpClassificacaoFiscal]", dbFailOnError
  
  'Exclui a tabela origem
  db.Execute "DROP TABLE [tmpClassificacaoFiscal]", dbFailOnError

  
  gbAlteraClassificacaoFiscal = True
  Exit Function
  
ErrCreate:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  gbAlteraClassificacaoFiscal = False

End Function



Private Function gbAlteraClassificacaoFiscalProduto(ByVal sTable As String) As Boolean
  Dim rs As Recordset
  Dim td As TableDef
  Dim fd As Field
  Dim bGotValor As Boolean
  Dim sField As String
  
  On Error GoTo ErrAlter
  
  Set td = db.TableDefs(sTable)
  
  If td("Classificação Fiscal").Type = dbInteger Then
    gbAlteraClassificacaoFiscalProduto = True
    Set td = Nothing
    Exit Function
  End If

  
  Set fd = td.CreateField("ClassificacaoFiscal", dbInteger)
  td.Fields.Append fd
  Set td = Nothing

  Set rs = db.OpenRecordset(sTable)
  If Not rs.EOF Then
    Do While Not rs.EOF
      rs.Edit
      rs("ClassificacaoFiscal").Value = rs("Classificação Fiscal").Value & ""
      rs.Update
      rs.MoveNext
    Loop
  End If

  rs.Close
  Set rs = Nothing

  Set td = db.TableDefs(sTable)
  td.Fields.Delete "Classificação Fiscal"
  td.Fields("ClassificacaoFiscal").Name = "Classificação Fiscal"
  Set td = Nothing
  
  gbAlteraClassificacaoFiscalProduto = True
  
 
  Exit Function
  
ErrAlter:
  If Err.Number = 3280 Then
    DoEvents
    td.Indexes.Delete ("Código Fiscal")
    Resume
  Else
    Screen.MousePointer = vbDefault
    Select Case frmErro.gnShowErr(Err.Number, "Alterar Código Fiscal")
      Case 0 'Repetir
        Resume
      Case 1 'Prosseguir
        Resume Next
      Case 2 'Sair
        Exit Function
      Case 3 'Encerrar
        End
    End Select
  End If
  gbAlteraClassificacaoFiscalProduto = False

End Function


Private Function gbAlteraCodigoFiscal(ByVal sTable As String) As Boolean
  Dim rs As Recordset
  Dim td As TableDef
  Dim fd As Field
  Dim bGotValor As Boolean
  Dim sField As String
  
  On Error GoTo ErrAlter
  
  Set td = db.TableDefs(sTable)
  
  If td("Código Fiscal").Size = 14 Then
    gbAlteraCodigoFiscal = True
    Set td = Nothing
    Exit Function
  End If
  
  Set fd = td.CreateField("CodFiscal2", dbText, 14)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set td = Nothing

  Set rs = db.OpenRecordset(sTable)
  If Not rs.EOF Then
    Do While Not rs.EOF
      rs.Edit
      rs("CodFiscal2").Value = rs("Código Fiscal").Value & ""
      rs.Update
      rs.MoveNext
    Loop
  End If

  rs.Close
  Set rs = Nothing

  Set td = db.TableDefs(sTable)
  td.Fields.Delete "Código Fiscal"
  td.Fields("CodFiscal2").Name = "Código Fiscal"
  Set td = Nothing
  
  gbAlteraCodigoFiscal = True
  
  If gbFirstCFOP = False Then
    DisplayMsg "O tamanho do campo ""Código Fiscal"" (CFOP) nas telas " & _
      vbCrLf & "Cadastro/Operações/Entrada e Saída foi alterado de 4 para " & _
      vbCrLf & "14 caracteres. No entanto, os lay-outs de Notas Fiscais " & _
      vbCrLf & "para exibirem este novo tamanho necessitarão de uma atualização " & _
      vbCrLf & "manual deste campo via Gerador/Lay-out de Notas..."
    gbFirstCFOP = True
  End If
  
  Exit Function
  
ErrAlter:
  If Err.Number = 3280 Then
    DoEvents
    td.Indexes.Delete ("Código Fiscal")
    Resume
  Else
    Screen.MousePointer = vbDefault
    Select Case frmErro.gnShowErr(Err.Number, "Alterar Código Fiscal")
      Case 0 'Repetir
        Resume
      Case 1 'Prosseguir
        Resume Next
      Case 2 'Sair
        Exit Function
      Case 3 'Encerrar
        End
    End Select
  End If
  gbAlteraCodigoFiscal = False

End Function

Private Function gbCreateTableTabCaractCliFor() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("TabCaractCliFor")
  
  Set fd = td.CreateField("CodCaract", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("TipoCliCaract", dbText, 1)
  td.Fields.Append fd
  Set fd = td.CreateField("DescCaract", dbText, 30)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("CodCaract")
  iX.Fields.Append iX.CreateField("TipoCliCaract")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableTabCaractCliFor = True
  Exit Function
  
ErrCreate:
  gbCreateTableTabCaractCliFor = False
  
End Function

Private Function gbCreateTableCliForNumeravel() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("CliForNumeravel")
  
  Set fd = td.CreateField("CodCliNumer", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("TipoCliNumer", dbText, 1)
  td.Fields.Append fd
  Set fd = td.CreateField("CodNumer", dbText, 15)
  td.Fields.Append fd
  Set fd = td.CreateField("CodProdNumer", dbText, 20)
  fd.AllowZeroLength = True
  fd.Required = False
  td.Fields.Append fd
  Set fd = td.CreateField("Data1Numer", dbDate)
  fd.Required = False
  td.Fields.Append fd
  Set fd = td.CreateField("Data2Numer", dbDate)
  fd.Required = False
  td.Fields.Append fd
  Set fd = td.CreateField("CodRefDocNumer", dbText, 20)
  fd.AllowZeroLength = True
  fd.Required = False
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("CodCliNumer")
  iX.Fields.Append iX.CreateField("TipoCliNumer")
  iX.Fields.Append iX.CreateField("CodNumer")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableCliForNumeravel = True
  Exit Function
  
ErrCreate:
  gbCreateTableCliForNumeravel = False
  
End Function

'Função criada por José por problemas de gravação parâmetros
Private Function gbCreateFieldComprim(ByVal sTableName As String, ByVal sFieldName As String, ByVal nType As Integer, ByVal nSize As Integer, ByVal sNum As String) As Boolean
  Dim td As TableDef
  Dim fd As Field
  'nType = dbBoolean, dbByte, dbInteger, dbSingle, dbDouble, dbCurrency, ...
  'nSize ignored for fixed-size fields and numeric fields...
  On Error GoTo ErrCreate
  If gbGetField("Parâmetros Filial", "Cód Comprim " & sNum) = False Then
    Set td = db.TableDefs(sTableName)
    Set fd = td.CreateField(sFieldName, nType, nSize)
    fd.AllowZeroLength = True
    td.Fields.Append fd
  Else
    Set td = db.TableDefs(sTableName)
    If (td.Fields(gnNum).AllowZeroLength = False) Then
      td.Fields.Delete ("Cód Comprim " & sNum)
      td.Fields.Refresh
      Set fd = td.CreateField(sFieldName, nType, nSize)
      fd.AllowZeroLength = True
      td.Fields.Append fd
    End If
  End If
  Set fd = Nothing
  Set td = Nothing

  gbCreateFieldComprim = True
  Exit Function
  
ErrCreate:
  gbCreateFieldComprim = False
End Function

Private Function gbAlteraZZZProgramas() As Boolean
  
  Dim sSql As String
  Dim nI As Integer
  Dim rsZZZ As Recordset
  
  
  sSql = "SELECT * FROM [ZZZProgramas]"
  Set rsZZZ = db.OpenRecordset(sSql, dbOpenDynaset)
  
  If rsZZZ.RecordCount = 0 Then Exit Function
  Call ws.BeginTrans
  rsZZZ.MoveLast
  rsZZZ.MoveFirst
  
  For nI = 0 To rsZZZ.RecordCount - 1
     If rsZZZ("Nome Programa") = "RELATÓRIO ENTRADAS" Then
        rsZZZ.Edit
           rsZZZ("ToolID") = 301101
        rsZZZ.Update
     ElseIf rsZZZ("Nome Programa") = "RELATÓRIO SAÍDAS" Then
        rsZZZ.Edit
           rsZZZ("ToolID") = 301102
        rsZZZ.Update
     ElseIf rsZZZ("Nome Programa") = "RELATÓRIO EMPRÉSTIMO ENTRADA" Then
        rsZZZ.Edit
           rsZZZ("ToolID") = 320039
        rsZZZ.Update
     ElseIf rsZZZ("Nome Programa") = "RELATÓRIO EMPRÉSTIMO SAÍDA" Then
        rsZZZ.Edit
           rsZZZ("ToolID") = 320040
        rsZZZ.Update
     ElseIf rsZZZ("Nome Programa") = "RELATÓRIO CLIENTES/FORNECEDORES" Then
        rsZZZ.Edit
           rsZZZ("Descrição") = "Relatório de Clientes e fornecedores"
        rsZZZ.Update
     End If
     rsZZZ.MoveNext
  Next nI
  
  Call ws.CommitTrans
  
  rsZZZ.Close
  Set rsZZZ = Nothing
  gbAlteraZZZProgramas = True
  
End Function

Private Function AddFileZZZProgramas() As Boolean
  Dim rsZZZ As Recordset
  
  On Error GoTo TratarErro
  
  Set rsZZZ = db.OpenRecordset("ZZZProgramas")
  
  rsZZZ.Index = "Nome"
  rsZZZ.Seek "=", "RELATÓRIO MOVIMENTOS"
  If rsZZZ.NoMatch Then
      rsZZZ.AddNew
         rsZZZ("Nome Programa") = "RELATÓRIO MOVIMENTOS"
         rsZZZ("Descrição") = "Relatório de Movimentos"
         rsZZZ("ToolID") = 320037
         rsZZZ("Número") = 0
      rsZZZ.Update
  End If
  
  With rsZZZ
    .Index = "Nome"
    .Seek "=", "MANUTENCAO DE ORCAMENTOS"
    If .NoMatch Then
      'Movimento -> Manutenção de Orçamento
      .AddNew
      .Fields("Nome Programa").Value = "MANUTENCAO DE ORCAMENTOS"
      .Fields("Descrição").Value = "Manutenção de Orçamentos"
      .Fields("Número").Value = 157
      .Fields("ToolID").Value = 320046
      .Update
    End If
  
    .Index = "Nome"
    .Seek "=", "MANUTENCAO DE CONSIGNACAO"
    If .NoMatch Then
      'Movimento -> Manutenção de Consignação
      .AddNew
      .Fields("Nome Programa").Value = "MANUTENCAO DE CONSIGNACAO"
      .Fields("Descrição").Value = "Manutenção de Consignação"
      .Fields("Número").Value = 158
      .Fields("ToolID").Value = 320045
      .Update
    End If
  End With
  
  rsZZZ.Close
  Set rsZZZ = Nothing
  
  AddFileZZZProgramas = True

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function
Private Function AddFileZZZProgramas2() As Boolean
  '11/05/2004 - Daniel
  'Adicionando dois registros: Rel. de Estoque por filiais (ToolID: 320043)
  'e Rel. Localização de Produtos (ToolID: 320051)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO DE ESTOQUE POR FILIAIS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO DE ESTOQUE POR FILIAIS"
      .Fields("Descrição").Value = "Relatório de Estoque por Filiais"
      .Fields("Número").Value = 159
      .Fields("ToolID").Value = 320043
      .Update
    End If
  End With
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO LOCALIZAÇÃO DE PRODUTOS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO LOCALIZAÇÃO DE PRODUTOS"
      .Fields("Descrição").Value = "Relatório Localização de Produtos"
      .Fields("Número").Value = 160
      .Fields("ToolID").Value = 320051
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas2 = True

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function

Private Function AddFileZZZProgramas3() As Boolean
  '27/05/2004 - Daniel
  'Adicionando um registro: Manutenção de Reservas (ToolID: 320050)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "MANUTENÇÃO DE RESERVAS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "MANUTENÇÃO DE RESERVAS"
      .Fields("Descrição").Value = "Manutenção de Reservas"
      .Fields("Número").Value = 161
      .Fields("ToolID").Value = 320050
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas3 = True

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function

Private Function AddFileZZZProgramas4() As Boolean
  '14/07/2004 - Daniel
  'Adicionando um registro: Classificação de Clientes (ToolID: 320058)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "CLASSIFICAÇÃO DOS CLIENTES"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "CLASSIFICAÇÃO DOS CLIENTES"
      .Fields("Descrição").Value = "Classificação dos Clientes"
      .Fields("Número").Value = 162
      .Fields("ToolID").Value = 320058
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas4 = True

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function

Private Function AddFileZZZProgramas5() As Boolean
  '01/10/2004 - Daniel
  'Adicionando um registro: Gerenciador de Loja Virtual (ToolID: 320042)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "GERENCIADOR DE PEDIDOS DA LOJA VIRTUAL"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "GERENCIADOR DE PEDIDOS DA LOJA VIRTUAL"
      .Fields("Descrição").Value = "Gerenciador de Pedidos da Loja Virtual"
      .Fields("Número").Value = 163
      .Fields("ToolID").Value = 320042
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas5 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas6() As Boolean
  '11/11/2004 - Daniel
  'Adicionando um registro: Relatório de Compras (ToolID: 301030)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO DE COMPRAS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO DE COMPRAS"
      .Fields("Descrição").Value = "Relatório de Compras"
      .Fields("Número").Value = 164
      .Fields("ToolID").Value = 301030
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas6 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas7() As Boolean
  '11/11/2004 - Daniel
  'Adicionando um registro: Relatório de Vendas por Clientes (ToolID: 320047)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO DE VENDAS POR CLIENTES"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO DE VENDAS POR CLIENTES"
      .Fields("Descrição").Value = "Relatório de Vendas por Clientes"
      .Fields("Número").Value = 165
      .Fields("ToolID").Value = 320047
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas7 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas8() As Boolean
  '11/11/2004 - Daniel
  'Adicionando um registro: Relatório C.R. - A Receber por Vendedor (ToolID: 302293)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO C.R. - A RECEBER POR VENDEDOR"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO C.R. - A RECEBER POR VENDEDOR"
      .Fields("Descrição").Value = "Relatório C.R. - A Receber por Vendedor"
      .Fields("Número").Value = 166
      .Fields("ToolID").Value = 302293
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas8 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas9() As Boolean
  '11/11/2004 - Daniel
  'Adicionando um registro: Relatório C.R. - Posição Financeira (ToolID: 320057)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO C.R. - POSIÇÃO FINANCEIRA"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO C.R. - POSIÇÃO FINANCEIRA"
      .Fields("Descrição").Value = "Relatório C.R. - Posição Financeira"
      .Fields("Número").Value = 167
      .Fields("ToolID").Value = 320057
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas9 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas10() As Boolean
  '11/11/2004 - Daniel
  'Adicionando um registro: Relatório C.R. - Emissão de Boletos (ToolID: 302300)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO C.R. - EMISSÃO DE BOLETOS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO C.R. - EMISSÃO DE BOLETOS"
      .Fields("Descrição").Value = "Relatório C.R. - Emissão de Boletos"
      .Fields("Número").Value = 168
      .Fields("ToolID").Value = 302300
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas10 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas11() As Boolean
  '11/11/2004 - Daniel
  'Adicionando um registro: Relatório C.R. - Emissão de Carnês (ToolID: 302301)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO C.R. - EMISSÃO DE CARNÊS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO C.R. - EMISSÃO DE CARNÊS"
      .Fields("Descrição").Value = "Relatório C.R. - Emissão de Carnês"
      .Fields("Número").Value = 169
      .Fields("ToolID").Value = 302301
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas11 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas12() As Boolean
  '06/06/2005 - Daniel
  'Adicionando o registro: Relatório de Usuários/Funcionários (ToolID: 40050)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "RELATÓRIO DE USUÁRIOS/FUNCIONÁRIOS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO DE USUÁRIOS/FUNCIONÁRIOS"
      .Fields("Descrição").Value = "Relatório de Usuários/Funcionários"
      .Fields("Número").Value = 170
      .Fields("ToolID").Value = 40050
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas12 = True

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

Private Function AddFileZZZProgramas13() As Boolean
  '08/08/2005 - Daniel
  'Adicionando o registro: Configuração de Impressoras (ToolID: 30030)
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    .Seek "=", "CONFIGURAÇÃO DE IMPRESSORAS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "CONFIGURAÇÃO DE IMPRESSORAS"
      .Fields("Descrição").Value = "Configuração de Impressoras"
      .Fields("Número").Value = 171
      .Fields("ToolID").Value = 30030
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas13 = True

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Function

'25/01/2006 - mpdea
'Inclusão de permissões
Private Function AddFileZZZProgramas14() As Boolean
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    'Descrição: Estoque das Filiais e Preços (Personalizado)
    'ToolID   : 320083
    .Seek "=", "RELATÓRIO ESTOQUE DAS FILIAIS E PREÇOS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO ESTOQUE DAS FILIAIS E PREÇOS"
      .Fields("Descrição").Value = "Relatório personalizado com estoque e preços"
      .Fields("Número").Value = 172
      .Fields("ToolID").Value = 320083
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas14 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro ao incluir permissão (14): " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'26/01/2006 - mpdea
'Inclusão de permissões
Private Function AddFileZZZProgramas15() As Boolean
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    'Descrição: Cadastro de Grupo Fiscal
    'ToolID   : 320084
    .Seek "=", "CADASTRO DE GRUPO FISCAL"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "CADASTRO DE GRUPO FISCAL"
      .Fields("Descrição").Value = "Cadastro de Grupo Fiscal"
      .Fields("Número").Value = 173
      .Fields("ToolID").Value = 320084
      .Update
    End If
    'Descrição:
    'ToolID   : 320085
    .Seek "=", "CADASTRO DE MENSAGENS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "CADASTRO DE MENSAGENS"
      .Fields("Descrição").Value = "Cadastro de Mensagens"
      .Fields("Número").Value = 174
      .Fields("ToolID").Value = 320085
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas15 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro ao incluir permissão (15): " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'19/07/2006 - Andrea
'Inclusão de permissões
Private Function AddFileZZZProgramas16() As Boolean
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    'Descrição: Relatório de Recebimentos por Forma de Pagamento
    'ToolID   : 320086
    .Seek "=", "RELATÓRIO RECEBIMENTOS FORMA PAGAMENTO"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO RECEBIMENTOS FORMA PAGAMENTO"
      .Fields("Descrição").Value = "Relatório de Recebimentos por Forma de Pagamento"
      .Fields("Número").Value = 175
      .Fields("ToolID").Value = 320086
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas16 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro ao incluir permissão (16): " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'16/07/2007 - Anderson
'Inclusão de permissões para contatos efetuados
Private Function AddFileZZZProgramas17() As Boolean
  Dim rstZZZProgramas As Recordset
  Dim rstFuncionarios As Recordset
  Dim rstAcessos As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    'Descrição: CONTATOS EFETUADOS
    'ToolID   :
    .Seek "=", "CONTATOS EFETUADOS"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "CONTATOS EFETUADOS"
      .Fields("Descrição").Value = "Contatos Efetuados"
      .Fields("Número").Value = 176
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  Set rstFuncionarios = db.OpenRecordset("SELECT Código, Liberado FROM Funcionários WHERE Liberado=-1")
  
  Do Until rstFuncionarios.EOF
  
    Set rstAcessos = db.OpenRecordset("SELECT * FROM Acessos WHERE Usuário=" & rstFuncionarios("Código") & " AND Numero=176")
    
    If rstAcessos.BOF And rstAcessos.EOF Then
      db.Execute "INSERT INTO Acessos (Programa,Usuário,Gravar,Apagar,Numero) Values ('CONTATOS EFETUADOS'," & rstFuncionarios("Código") & ",-1,-1,176)", dbFailOnError
    End If
    
    rstAcessos.Close
    Set rstAcessos = Nothing
    
    rstFuncionarios.MoveNext
    
  Loop
  
  rstFuncionarios.Close
  Set rstFuncionarios = Nothing
  
  AddFileZZZProgramas17 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro ao incluir permissão (17): " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'07/08/2007 - Anderson
'Implementação de relatório de Comissão por vendedor para CandyClean
Private Function AddFileZZZProgramas18() As Boolean
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    'Descrição: Relatório de Comissões por Vendedor
    'ToolID   :
    .Seek "=", "RELATÓRIO DE COMISSÕES POR VENDEDOR"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO DE COMISSÕES POR VENDEDOR"
      .Fields("Descrição").Value = "Relatório de Comissões por Vendedor"
      .Fields("Número").Value = 177
      .Fields("ToolID").Value = 320089
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
    
  AddFileZZZProgramas18 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro ao incluir permissão (18): " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'30/10/2007 - Anderson
'Programa para relatório de produtos a comprar
Private Function AddFileZZZProgramas19() As Boolean
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome" 'Campo Chave
    'Descrição: Relatório de produtos a comprar com fator
    'ToolID   :
    .Seek "=", "RELATÓRIO DE PRODUTOS A COMPRAR FATOR"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO DE PRODUTOS A COMPRAR FATOR"
      .Fields("Descrição").Value = "Relatório de Produtos a Comprar com Fator"
      .Fields("Número").Value = 178
      .Fields("ToolID").Value = 320090
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
    
  AddFileZZZProgramas19 = True
  
  Exit Function

TratarErro:
  MsgBox "Erro ao incluir permissão (19): " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'25/09/2008 - mpdea
'Inclusão de permissões ausentes para relatórios de vendas
'Solicitado pelo Patrício (Technomax)
Private Function AddFileZZZProgramas20() As Boolean
  Dim rstZZZProgramas As Recordset
  
  On Error GoTo TratarErro
  
  Set rstZZZProgramas = db.OpenRecordset("ZZZProgramas")
  
  With rstZZZProgramas
    .Index = "Nome"
    .Seek "=", "RELATÓRIO VENDAS POR VENDEDOR/COMISSÕES"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO VENDAS POR VENDEDOR/COMISSÕES"
      .Fields("Descrição").Value = "Relatório de Vendas por Vendedor e Comissões"
      .Fields("Número").Value = 179
      .Fields("ToolID").Value = 320076
      .Update
    End If
  End With
  
  With rstZZZProgramas
    .Index = "Nome"
    .Seek "=", "RELATÓRIO VENDAS POR TAMANHO"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO VENDAS POR TAMANHO"
      .Fields("Descrição").Value = "Relatório de Vendas por Tamanho"
      .Fields("Número").Value = 180
      .Fields("ToolID").Value = 320048
      .Update
    End If
  End With
  
  With rstZZZProgramas
    .Index = "Nome"
    .Seek "=", "RELATÓRIO VENDAS POR EDITORA"
    If .NoMatch Then
      .AddNew
      .Fields("Nome Programa").Value = "RELATÓRIO VENDAS POR EDITORA"
      .Fields("Descrição").Value = "Relatório de Vendas por Editora"
      .Fields("Número").Value = 181
      .Fields("ToolID").Value = 320066
      .Update
    End If
  End With
  
  rstZZZProgramas.Close
  Set rstZZZProgramas = Nothing
  
  AddFileZZZProgramas20 = True

  Exit Function

TratarErro:
  MsgBox "Erro ao incluir permissão (20): " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Function

Private Function UpdateRecordParametros() As Boolean
  '12/05/2004 - Daniel
  'Esta função têm a finalidade de colocar zero
  'nos campos CSLL, COFINS, PIS, IRRF da tabela
  'de Parâmetros caso o campo esteja vazio...
'  Dim rstParametros As Recordset

  On Error GoTo ErrHandler

  'Manutenido em 26/05/2004 - Daniel & MPDEA
  db.Execute "UPDATE [Parâmetros Filial] SET CSLL = 0 WHERE CSLL IS NULL;", dbFailOnError
  db.Execute "UPDATE [Parâmetros Filial] SET COFINS = 0 WHERE COFINS IS NULL;", dbFailOnError
  db.Execute "UPDATE [Parâmetros Filial] SET PIS = 0 WHERE PIS IS NULL;", dbFailOnError
  db.Execute "UPDATE [Parâmetros Filial] SET IRRF = 0 WHERE IRRF IS NULL;", dbFailOnError
'  DBEngine.Idle dbRefreshCache
  
'  Set rstParametros = db.OpenRecordset("SELECT CSLL, COFINS, PIS, IRRF FROM [Parâmetros Filial]", dbOpenDynaset)
'
'  'Caso não tenha nada na tabela
'  'sai fora e seta a função como True
'  If rstParametros.RecordCount = 0 Then
'    UpdateRecordParametros = True
'
'    rstParametros.Close
'    Set rstParametros = Nothing
'
'    Exit Function
'  Else
'    rstParametros.MoveLast
'    rstParametros.MoveFirst
'  End If
'
'  With rstParametros
'    If Not (.BOF And .EOF) Then
'      '.MoveFirst
'
'      Do Until .EOF
'        .Edit
'
'        'CSLL
'        If Not IsNumeric(.Fields("CSLL").Value) Then
'          .Fields("CSLL").Value = 0
'        End If
'        'COFINS
'        If Not IsNumeric(.Fields("COFINS").Value) Then
'          .Fields("COFINS").Value = 0
'        End If
'        'PIS
'        If Not IsNumeric(.Fields("PIS").Value) Then
'          .Fields("PIS").Value = 0
'        End If
'        'IRRF
'        If Not IsNumeric(.Fields("IRRF").Value) Then
'          .Fields("IRRF").Value = 0
'        End If
'
'        .Update
'        .MoveNext
'      Loop
'
'    End If
'    .Close
'  End With
'
'  Set rstParametros = Nothing
  UpdateRecordParametros = True
  
  Exit Function
  
ErrHandler:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  UpdateRecordParametros = False

End Function

'31/07/2002 - mpdea
'Alterado tratamento da atualização do campo e erro com tabela em aberto
Private Function gbGravaTrueParamSaiPrecos() As Boolean

  On Error GoTo TrataErro
  
  db.Execute "UPDATE [Parâmetros Filial] SET [Saida Altera Preco] = True", dbFailOnError
  gbGravaTrueParamSaiPrecos = True
  
  Exit Function
  
TrataErro:
  gbGravaTrueParamSaiPrecos = False
  
End Function

'19/09/2002 - mpdea
'Comentado função não utilizada

'Private Function gbAlteraSitTributaria(ByVal sTable As String) As Boolean
'  Dim rs As Recordset
'  Dim td As TableDef
'  Dim fd As Field
'  Dim bGotValor As Boolean
'  Dim sField As String
'
'  On Error GoTo ErrAlter
'
'  Set td = db.TableDefs(sTable)
'
'  If td("Situação Tributária").Size = 14 Then
'    gbAlteraSitTributaria = True
'    Set td = Nothing
'    Exit Function
'  End If
'
'  Set fd = td.CreateField("Situa2", dbText, 3)
'  fd.AllowZeroLength = True
'  td.Fields.Append fd
'  Set td = Nothing
'
'  Set rs = db.OpenRecordset(sTable)
'  If Not rs.EOF Then
'    Do While Not rs.EOF
'      rs.Edit
'      rs("Situa2").Value = rs("Situação Tributária").Value & ""
'      rs.Update
'      rs.MoveNext
'    Loop
'  End If
'
'  rs.Close
'  Set rs = Nothing
'
'  Set td = db.TableDefs(sTable)
'  td.Fields.Delete "Situação Tributária"
'  td.Fields("Situa2").Name = "Situação Tributária"
'  Set td = Nothing
'
'  gbAlteraSitTributaria = True
'
''  If gbFirstCFOP = False Then
''    DisplayMsg "O tamanho do campo ""Código Fiscal"" (CFOP) nas telas " & _
''      vbCrLf & "Cadastro/Operações/Entrada e Saída foi alterado de 4 para " & _
''      vbCrLf & "14 caracteres. No entanto, os lay-outs de Notas Fiscais " & _
''      vbCrLf & "para exibirem este novo tamanho necessitarão de uma atualização " & _
''      vbCrLf & "manual deste campo via Gerador/Lay-out de Notas..."
''    gbFirstCFOP = True
''  End If
''
'  Exit Function
'
'ErrAlter:
'  If Err.Number = 3280 Then
'    DoEvents
''    td.Indexes.Delete ("Código Fiscal")
'    Resume
'  Else
'    Screen.MousePointer = vbDefault
'    Select Case frmErro.gnShowErr(Err.Number, "Alterar Situação Tributária")
'      Case 0 'Repetir
'        Resume
'      Case 1 'Prosseguir
'        Resume Next
'      Case 2 'Sair
'        Exit Function
'      Case 3 'Encerrar
'        End
'    End Select
'  End If
'  gbAlteraSitTributaria = False
'
'End Function

Private Function gbAlteraIcmEntra(ByVal sTable As String) As Boolean

  Dim rs As Recordset
  Dim td As TableDef
  Dim fd As Field
  Dim bGotValor As Boolean
  Dim sField As String
  
  On Error GoTo ErrAlter
  
  Set td = db.TableDefs(sTable)
  
  If gbGetField("Produtos", "Percentual Icm Entrada") = True Then
    gbAlteraIcmEntra = True
    Set td = Nothing
    Exit Function
  End If
  
  Set fd = td.CreateField("Percentual Icm Entrada", dbSingle)
'  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set td = Nothing

  Set rs = db.OpenRecordset(sTable)
  If Not rs.EOF Then
    Do While Not rs.EOF
      rs.Edit
      rs("Percentual Icm Entrada").Value = rs("Percentual Icm").Value & ""
      rs.Update
      rs.MoveNext
    Loop
  End If

  rs.Close
  Set rs = Nothing

  Set td = Nothing
  
  gbAlteraIcmEntra = True
  
'  If gbFirstCFOP = False Then
'    DisplayMsg "O tamanho do campo ""Código Fiscal"" (CFOP) nas telas " & _
'      vbCrLf & "Cadastro/Operações/Entrada e Saída foi alterado de 4 para " & _
'      vbCrLf & "14 caracteres. No entanto, os lay-outs de Notas Fiscais " & _
'      vbCrLf & "para exibirem este novo tamanho necessitarão de uma atualização " & _
'      vbCrLf & "manual deste campo via Gerador/Lay-out de Notas..."
'    gbFirstCFOP = True
'  End If
'
  Exit Function
  
ErrAlter:
  If Err.Number = 3280 Then
    DoEvents
'    td.Indexes.Delete ("Código Fiscal")
    Resume
  Else
    Screen.MousePointer = vbDefault
    Select Case frmErro.gnShowErr(Err.Number, "Alterar Icm Entrada")
      Case 0 'Repetir
        Resume
      Case 1 'Prosseguir
        Resume Next
      Case 2 'Sair
        Exit Function
      Case 3 'Encerrar
        End
    End Select
  End If
  gbAlteraIcmEntra = False

End Function

Private Function gbAlteraIcmSai(ByVal sTable As String) As Boolean

  Dim rs As Recordset
  Dim td As TableDef
  Dim fd As Field
  Dim bGotValor As Boolean
  Dim sField As String
  
  On Error GoTo ErrAlter
  
  Set td = db.TableDefs(sTable)
  
  If gbGetField("Produtos", "Percentual Icm Saida") = True Then
    gbAlteraIcmSai = True
    Set td = Nothing
    Exit Function
  End If
  
  Set fd = td.CreateField("Percentual Icm Saida", dbSingle)
'  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set td = Nothing

  Set rs = db.OpenRecordset(sTable)
  If Not rs.EOF Then
    Do While Not rs.EOF
      rs.Edit
      rs("Percentual Icm Saida").Value = rs("Percentual Icm").Value & ""
      rs.Update
      rs.MoveNext
    Loop
  End If

  rs.Close
  Set rs = Nothing

  Set td = Nothing
  
  gbAlteraIcmSai = True
  
'  If gbFirstCFOP = False Then
'    DisplayMsg "O tamanho do campo ""Código Fiscal"" (CFOP) nas telas " & _
'      vbCrLf & "Cadastro/Operações/Entrada e Saída foi alterado de 4 para " & _
'      vbCrLf & "14 caracteres. No entanto, os lay-outs de Notas Fiscais " & _
'      vbCrLf & "para exibirem este novo tamanho necessitarão de uma atualização " & _
'      vbCrLf & "manual deste campo via Gerador/Lay-out de Notas..."
'    gbFirstCFOP = True
'  End If
'
  Exit Function
  
ErrAlter:
  If Err.Number = 3280 Then
    DoEvents
'    td.Indexes.Delete ("Código Fiscal")
    Resume
  Else
    Screen.MousePointer = vbDefault
    Select Case frmErro.gnShowErr(Err.Number, "Alterar Icm Saida")
      Case 0 'Repetir
        Resume
      Case 1 'Prosseguir
        Resume Next
      Case 2 'Sair
        Exit Function
      Case 3 'Encerrar
        End
    End Select
  End If
  gbAlteraIcmSai = False

End Function

Private Function m_blnCreateIndexFabricante() As Boolean
  '29/03/2005 - Daniel
  '
  'Private criada para atender a necessidade de busca
  'do fabricante de cada produto
  'Solicitação: Elétrica Leal
  Dim iX As Index
  Dim td As TableDef
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Produtos")
  Set iX = td.CreateIndex
  
  With iX
    .Name = "Fabricante"
    .Fields.Append .CreateField("Código")
    .Fields.Append .CreateField("Fabricante")
    .Primary = False
    .Unique = False
    .IgnoreNulls = False
  End With
  td.Indexes.Append iX
  
  td.Indexes.Refresh
  
  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexFabricante = True
  
  Exit Function

ErrHandler:
  m_blnCreateIndexFabricante = False
  MsgBox "Erro ao criar índice [Fabricante] na tabela [Produtos]: " & _
  Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'15/01/2005 - Daniel
'Cria índice na tabela de [Contas a Receber] para otimizar buscas
'do Quick CNAB
'Solicitação: TV Shopping Brasil
Private Function m_blnCreateIndexCNAB() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Contas a Receber")
  Set iX = td.CreateIndex
  
  With iX
    .Name = "CNAB"
    .Fields.Append .CreateField("Tipo")
    .Fields.Append .CreateField("Data Emissão")
    .Fields.Append .CreateField("Vencimento")
    .Fields.Append .CreateField("Valor Recebido")
    .Fields.Append .CreateField("CNAB_NossoNumero")
    .Primary = False
    .Unique = False
    .IgnoreNulls = False
  End With
  td.Indexes.Append iX
  
  td.Indexes.Refresh
  
  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexCNAB = True
  
  Exit Function

ErrHandler:
  m_blnCreateIndexCNAB = False
  MsgBox "Erro ao criar índice [CNAB] na tabela [Contas a Receber]: " & _
  Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'01/03/2005 - Daniel
'Cria índice na tabela de funcionários para otimizar o acesso no momento
'do login
Private Function m_blnCreateIndexAcessando() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Funcionários")
  Set iX = td.CreateIndex
  
  With iX
    .Name = "Acessando"
    .Fields.Append .CreateField("Código")
    .Fields.Append .CreateField("Senha")
    .Primary = False
    .Unique = False
    .IgnoreNulls = False
  End With
  td.Indexes.Append iX
  
  td.Indexes.Refresh
  
  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexAcessando = True
  
  Exit Function

ErrHandler:
  m_blnCreateIndexAcessando = False
  MsgBox "Erro ao criar índice [Acessando] na tabela [Funcionários]: " & _
  Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'01/03/2005 - Daniel
'Cria índice na tabela de saídas para otimizar a busca por nota
Private Function m_blnCreateIndexNota() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Saídas")
  Set iX = td.CreateIndex
  
  With iX
    .Name = "Nota"
    .Fields.Append .CreateField("Filial")
    .Fields.Append .CreateField("Nota Impressa")
    .Primary = False
    .Unique = False
    .IgnoreNulls = False
  End With
  td.Indexes.Append iX
  
  td.Indexes.Refresh
  
  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexNota = True
  
  Exit Function

ErrHandler:
  m_blnCreateIndexNota = False
  MsgBox "Erro ao criar índice [Nota] na tabela [Saídas]: " & _
  Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'24/06/2005 - Daniel
'Cria índice na tabela de produtos para otimizar a busca pelo CodigoNBM
Private Function m_blnCreateIndexCodigoNBM() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Produtos")
  Set iX = td.CreateIndex
  
  With iX
    .Name = "CodigoNBM"
    .Fields.Append .CreateField("Código")
    .Fields.Append .CreateField("CodigoNBM")
    .Primary = False
    .Unique = False
    .IgnoreNulls = False
  End With
  td.Indexes.Append iX
  
  td.Indexes.Refresh
  
  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexCodigoNBM = True
  
  Exit Function

ErrHandler:
  m_blnCreateIndexCodigoNBM = False
  MsgBox "Erro ao criar índice [CodigoNBM] na tabela [Produtos]: " & _
  Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'24/09/2003 - mpdea
'Cria índice na tabela de saídas para agilizar a pesquisa
'de movimentações
Private Function m_blnCreateIndexVrAchaVenda() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Saídas")
  Set iX = td.CreateIndex
  With iX
    .Name = "VrAchaVenda"
    .Fields.Append .CreateField("Filial")
    .Fields.Append .CreateField("Efetivada")
    .Fields.Append .CreateField("Data")
    .Primary = False
    .Unique = False
    .IgnoreNulls = False
  End With
  td.Indexes.Append iX

  ' Refresh collection so that you can access new Index objects.
  td.Indexes.Refresh

  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexVrAchaVenda = True
  
  Exit Function
  
ErrHandler:
  m_blnCreateIndexVrAchaVenda = False
  MsgBox "Erro ao criar índice [VrAchaVenda] na tabela [Saídas]: " & _
    Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Function

'24/09/2003 - mpdea
'Cria índice na tabela de saídas para agilizar a pesquisa
'de movimentações
Private Function m_blnCreateIndexSaidasDataMov() As Boolean
  Dim iX As Index
  Dim td As TableDef
  
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Saídas")
  Set iX = td.CreateIndex
  With iX
    .Name = "DataMov"
    .Fields.Append .CreateField("Data")
    .Primary = False
    .Unique = False
    .IgnoreNulls = False
  End With
  td.Indexes.Append iX

  ' Refresh collection so that you can access new Index objects.
  td.Indexes.Refresh

  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexSaidasDataMov = True
  
  Exit Function
  
ErrHandler:
  m_blnCreateIndexSaidasDataMov = False
  MsgBox "Erro ao criar índice [DataMov] na tabela [Saídas]: " & _
    Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Function

'26/01/2006 - mpdea
'Inclusão da tabela Grupo Fiscal
Private Function m_blnCreateTableGrupoFiscal() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  
  On Error GoTo ErrCreate
  
  
  Set td = db.CreateTableDef("GrupoFiscal")
  
  Set fd = td.CreateField("Código", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome", dbText, 50)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  Set fd = td.CreateField("Data Alteração", dbText, 10)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Código")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  m_blnCreateTableGrupoFiscal = True
  Exit Function
  
ErrCreate:
  m_blnCreateTableGrupoFiscal = False
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'26/01/2006 - mpdea
'Inclusão da tabela Mensagens para Nota Fiscal
Private Function m_blnCreateTableMensagensNotaFiscal() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  
  On Error GoTo ErrCreate
  
  
  Set td = db.CreateTableDef("MensagensNotaFiscal")
  
  Set fd = td.CreateField("Codigo", dbLong)
  fd.Attributes = dbAutoIncrField
  td.Fields.Append fd
  
  Set fd = td.CreateField("Ordem", dbInteger)
  fd.Required = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("TipoFiltroProduto", dbByte)
  fd.Required = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("TipoFiltroOpSaida", dbByte)
  fd.Required = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("TipoFiltroUF", dbByte)
  fd.Required = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("FiltroProduto", dbText, 20)
  fd.Required = True
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("FiltroOpSaida", dbText, 20)
  fd.Required = True
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("FiltroUF", dbText, 20)
  fd.Required = True
  fd.AllowZeroLength = False
  td.Fields.Append fd
  
  Set fd = td.CreateField("Mensagem", dbText, 80)
  fd.Required = True
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  
  'Índice para Codigo
  Set iX = td.CreateIndex("Codigo")
  iX.Fields.Append iX.CreateField("Codigo")
  iX.Unique = True
  td.Indexes.Append iX
  'Índice para Ordem
  Set iX = td.CreateIndex("Ordem")
  iX.Fields.Append iX.CreateField("Ordem")
  td.Indexes.Append iX
  
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  m_blnCreateTableMensagensNotaFiscal = True
  Exit Function
  
ErrCreate:
  m_blnCreateTableMensagensNotaFiscal = False
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'26/01/2007 - Anderson
'Tabela para armazenar o recebimento da manutenção da conta do cliente
Private Function m_blnCreateTableContaClienteRecebimento() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
    
  On Error GoTo ErrCreate
    
  Set td = db.CreateTableDef("ContaClienteRecebimento")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Contador", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Sequencia", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Recebe - Dinheiro", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Recebe - Emp Cartão", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Recebe - Num Cartão", dbText, 20)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Recebe - Cartão", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Recebe - Vale", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Total Prazo", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Tipo Parcela", dbText, 1)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Conta", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Parcela Cartão", dbText, 1)
  td.Fields.Append fd
    
  Set fd = td.CreateField("Qtde Parcelas", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor Parcela", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Valor Recebido", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Troco", dbDouble)
  td.Fields.Append fd
        
  db.TableDefs.Append td
  
  Set td = Nothing
  
  m_blnCreateTableContaClienteRecebimento = True
  Exit Function
  
ErrCreate:
  m_blnCreateTableContaClienteRecebimento = False
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

Private Function gbCreateTableProdutoCesta() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ProdutoCesta")
  
  Set fd = td.CreateField("CodigoCesta", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoItem", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("QuantidadeItem", dbLong)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("CodigoCesta")
  iX.Fields.Append iX.CreateField("CodigoItem")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableProdutoCesta = True
  Exit Function
  
ErrCreate:
  gbCreateTableProdutoCesta = False
  
End Function


Private Function gbCreateTableProdutoPareamentoFornecedor() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ProdutoPareamentoFornecedor")
  
  Set fd = td.CreateField("Produto", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("Tipo", dbText, 1)
  td.Fields.Append fd
  Set fd = td.CreateField("ProdutoForn", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("Fornecedor", dbLong)
  td.Fields.Append fd

    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Produto")
  iX.Fields.Append iX.CreateField("ProdutoForn")
  iX.Fields.Append iX.CreateField("Fornecedor")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableProdutoPareamentoFornecedor = True
  Exit Function
  
ErrCreate:
  gbCreateTableProdutoPareamentoFornecedor = False
  
End Function

Private Function gbCreateTableTransferenciaEntreFiliais() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("TransferenciaEntreFiliais")
  
  Set fd = td.CreateField("CodigoTransf", dbLong)
  fd.Attributes = dbAutoIncrField
  td.Fields.Append fd
  Set fd = td.CreateField("FilialLogada", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("FilialExportada", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoFornecedor", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoCliente", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoOperSaida", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoOperEntrada", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("TabelaPrecos", dbText, 15)
  td.Fields.Append fd
  Set fd = td.CreateField("PermitirTransfEstoqueInsuf", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("Data", dbDate)
  td.Fields.Append fd
  Set fd = td.CreateField("Status", dbInteger)    '(1-Gravado;2-Efetivada)
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoUsuario", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("QuantidadeItens", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("NumItens", dbInteger)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("CodigoTransf")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableTransferenciaEntreFiliais = True
  Exit Function
  
ErrCreate:
  gbCreateTableTransferenciaEntreFiliais = False
  
End Function


Private Function gbCreateTableTransferenciaProdutos() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("TransferenciaProdutos")
  
  Set fd = td.CreateField("CodigoTransf", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("codigoProduto", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("nomeProduto", dbText, 100)
  td.Fields.Append fd
  Set fd = td.CreateField("Quantidade", dbInteger)
  td.Fields.Append fd
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableTransferenciaProdutos = True
  Exit Function
  
ErrCreate:
  gbCreateTableTransferenciaProdutos = False
  
End Function


'08/12/2006 - Anderson
'Criação da tabela para o registro de CFOP por produto
Private Function gbCreateTableProdutoCFOP() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ProdutoCFOP")
  
  Set fd = td.CreateField("CodProduto", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("CodOperacao", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("CFOP", dbText, 14)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("CodProduto")
  iX.Fields.Append iX.CreateField("CodOperacao")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableProdutoCFOP = True
  Exit Function
  
ErrCreate:
  gbCreateTableProdutoCFOP = False
  
End Function

'15/12/2006 - Anderson
'Criação da tabela para o registro de CFOP por serviço
Private Function gbCreateTableServicoCFOP() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ServicoCFOP")
  
  Set fd = td.CreateField("CodServico", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("CodOperacao", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("CFOP", dbText, 14)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("CodServico")
  iX.Fields.Append iX.CreateField("CodOperacao")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableServicoCFOP = True
  Exit Function
  
ErrCreate:
  gbCreateTableServicoCFOP = False
  
End Function

'15/12/2006 - Anderson
'Criação da tabela para o registro de CFOP por serviço
Private Function gbCreateTableCNABCarteira() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("CNABCarteira")
  
  Set fd = td.CreateField("NumeroCarteira", dbText, 3)
  td.Fields.Append fd
  Set fd = td.CreateField("Banco", dbText, 25)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("NumeroCarteira")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableCNABCarteira = True
  Exit Function
  
ErrCreate:
  gbCreateTableCNABCarteira = False
  
End Function

'15/12/2006 - Anderson
'Alterar campo de ICMS para valores quebrados
Private Function gbUpdateFieldEstados() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index

  On Error GoTo ErrCreate

  Set td = db.TableDefs("Estados")

  Set fd = td.Fields("ICM")
  fd.Type = 6
  td.Fields.Append fd

  db.TableDefs.Append td

  Set td = Nothing

  gbUpdateFieldEstados = True
  Exit Function

ErrCreate:
  gbUpdateFieldEstados = False
  
End Function

Private Function m_blnCreateIndexCarneCodigoBarras() As Boolean
  '25/09/2007 - Anderson
  'Função Criada para otimizar o processo de manutenção de contas a receber através de carnês
  'Solicitação: Naativa (QS73159-473)
  Dim iX As Index
  Dim td As TableDef
  
  On Error GoTo ErrHandler
  
  Set td = db.TableDefs("Contas a Receber")
  Set iX = td.CreateIndex
  
  With iX
    .Name = "CarneCodigoBarras"
    .Fields.Append .CreateField("CarneCodigoBarras")
    .Primary = False
    .Unique = False
    .IgnoreNulls = True
  End With
  td.Indexes.Append iX
  
  td.Indexes.Refresh
  
  Set iX = Nothing
  Set td = Nothing
  
  m_blnCreateIndexCarneCodigoBarras = True
  
  Exit Function

ErrHandler:
  m_blnCreateIndexCarneCodigoBarras = False
  MsgBox "Erro ao criar índice [CarneCodigoBarras] na tabela [Contas a Receber]: " & _
  Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'10/12/2009 - Andrea
'Criação da tabela para o armazenas os dados de recebimento em cartoes
Private Function gbCreateTableMovimentoCartoes() As Boolean
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Movimento - Cartoes")
  
  Set fd = td.CreateField("Filial", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("Sequência", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("Ordem", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("Administradora", dbText, 25)
  td.Fields.Append fd
  Set fd = td.CreateField("Valor", dbDouble)
  td.Fields.Append fd
  Set fd = td.CreateField("Parcelas", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("ValorParcelas", dbDouble)
  td.Fields.Append fd
  Set fd = td.CreateField("NumeroCartao", dbText, 25)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableMovimentoCartoes = True
  Exit Function
  
ErrCreate:
  gbCreateTableMovimentoCartoes = False
  
End Function


'19/12/2007 - Anderson
'Criação da tabela para o registro de CFOP por serviço
Private Function gbCreateTableNSU() As Boolean
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("NSU")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  Set fd = td.CreateField("NSU", dbText, 10)
  td.Fields.Append fd
  Set fd = td.CreateField("Movimento", dbText, 10)
  td.Fields.Append fd
  Set fd = td.CreateField("Motivo", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("Sequencia", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("NotaFiscal", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("Data_Hora", dbDate)
  td.Fields.Append fd
  Set fd = td.CreateField("Total", dbDouble)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableNSU = True
  Exit Function
  
ErrCreate:
  gbCreateTableNSU = False
  
End Function

'30/01/2009 - mpdea
'Criação da tabela para configuração de envio de e-mail
Private Function gbCreateTableEmail() As Boolean
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Email")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  Set fd = td.CreateField("ServidorSmtp", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("ServidorPop3", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("Autenticacao", dbBoolean)
  td.Fields.Append fd
  Set fd = td.CreateField("AutenticacaoPop3", dbBoolean)
  td.Fields.Append fd
  Set fd = td.CreateField("Usuario", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("Senha", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("NomeExibicaoRemetente", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("EmailRemetente", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableEmail = True
  Exit Function
  
ErrCreate:
  gbCreateTableEmail = False
  
End Function

Private Function gbCreateTableSaidasComandas()
'15/05/2013-Alexandre Afornali
'Case DiskEmbalagens
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("SaidasComandas")
  
  Set fd = td.CreateField("CodComanda", dbText, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("CodSaida", dbText, 20)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableSaidasComandas = True
  Exit Function
  
ErrCreate:
  gbCreateTableSaidasComandas = False
End Function

Private Function gbCreateTableAliquotasNCM()
  '26/06/2013-Alexandre Afornali
  'Case lei De Olho nos Impostos
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("AliquotasNCM")
  
  Set fd = td.CreateField("Codigo", dbText, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("EX", dbText, 2)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Tabela", dbLong, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("AliqNacional", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("AliqImportacao", dbDouble)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableAliquotasNCM = True
  Exit Function
  
ErrCreate:
  gbCreateTableAliquotasNCM = False
End Function

Private Function gbCreateTableAcessoTabelasDePrecosProdutos()
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("AcessoTabelasDePrecosProdutos")
  
  Set fd = td.CreateField("Usuario", dbLong, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Tabela", dbText, 15)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableAcessoTabelasDePrecosProdutos = True
  Exit Function
  
ErrCreate:
  gbCreateTableAcessoTabelasDePrecosProdutos = False
End Function

''
Private Function gbCreateTableSaidasChaves()
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("SaidasChaves")
  
  Set fd = td.CreateField("Filial", dbLong, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Sequencia", dbLong, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Chave", dbText, 55)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableSaidasChaves = True
  Exit Function
  
ErrCreate:
  gbCreateTableSaidasChaves = False
End Function

''

Private Function gbCreateTableDRE_quick()
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("DRE_quick")
  
  Set fd = td.CreateField("CodigoDRE", dbLong)
  fd.Attributes = dbAutoIncrField
  td.Fields.Append fd
  
  Set fd = td.CreateField("Filial", dbLong, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Usuario", dbLong, 13)
  td.Fields.Append fd
  
'  Set fd = td.CreateField("DataIni", dbDate)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("DataFim", dbDate)
'  td.Fields.Append fd

  Set fd = td.CreateField("DataANO", dbLong)
  td.Fields.Append fd

  Set fd = td.CreateField("DataMES", dbLong)
  td.Fields.Append fd

  Set fd = td.CreateField("DataCriacao", dbDate)
  td.Fields.Append fd

  Set fd = td.CreateField("Obs", dbText, 255)
  td.Fields.Append fd
  
  Set fd = td.CreateField("ReceitaBruta", dbCurrency)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Devolucoes", dbCurrency)
  td.Fields.Append fd
  
  Set fd = td.CreateField("ImpostoSobreVendas", dbCurrency)
  td.Fields.Append fd
  
  Set fd = td.CreateField("ReceitaOperacionalLiquida", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("CMV", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("LucroBruto", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("DespesasAdministrativas", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("DespesasComerciais", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("DespesasDepreciacao", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("DespesasFinanceiras", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("ReceitasFinanceiras", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("LucroPrejuizoOperacional", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("DespesasNaoOperacionais", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("ReceitasNaoOperacionais", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("LAIR", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("ProvisaoIR", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("ProvisaoCSLL", dbCurrency)
  td.Fields.Append fd

  Set fd = td.CreateField("LucroLiquido", dbCurrency)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableDRE_quick = True
  Exit Function
  
ErrCreate:
  gbCreateTableDRE_quick = False
End Function

Private Function gbCreateTableDRE_anexos()
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("DRE_anexos")
  
  Set fd = td.CreateField("CodigoAnexo", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Obs", dbText, 150)
  td.Fields.Append fd
  
  Set fd = td.CreateField("ValorDe", dbCurrency)
  td.Fields.Append fd
  
  Set fd = td.CreateField("ValorAte", dbCurrency)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Aliquota", dbCurrency)
  td.Fields.Append fd
  
  Set fd = td.CreateField("ValorRedutor", dbCurrency)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableDRE_anexos = True
  Exit Function
  
ErrCreate:
  gbCreateTableDRE_anexos = False
End Function

Private Function gbCreateTableProdutoFavoritos()
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ProdutoFavoritos")
  
  Set fd = td.CreateField("Filial", dbInteger)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Produto", dbText, 20)
  td.Fields.Append fd
  
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableProdutoFavoritos = True
  Exit Function
  
ErrCreate:
  gbCreateTableProdutoFavoritos = False
End Function

Private Function gbCreateTableCodigoBeneficio()
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("CodigoBeneficio")
  
  Set fd = td.CreateField("Estado", dbText, 2)
  td.Fields.Append fd
  
  Set fd = td.CreateField("CodigoBenef", dbText, 10)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableCodigoBeneficio = True
  Exit Function
  
ErrCreate:
  gbCreateTableCodigoBeneficio = False
End Function


Private Function gbCreateTableTBLRelVendasPorVendedor()
  '14/11/2014 - Eduardo
  'Tabela tblRelVendasPorVendedor - Case InfoSocial
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
 
  If gnOpenTempDB(gsTempDBFileName, False) = 0 Then
  
    If gbGetTableTemp("tblRelVendasPorVendedor") = False Then
    'Set td = dbTemp.CreateTableDef("tblRelVendasPorVendedor")
      dbTemp.Execute "CREATE TABLE tblRelVendasPorVendedor([Filial] INTEGER,[Vendedor] INTEGER,[DataIni1] DATE,[DataFim1] DATE,[DataIni2] DATE,[DataFim2] DATE,[DataIni3] DATE,[DataFim3] DATE,[Operacao] INTEGER,[Cliente] LONG,[SumMes1] DOUBLE,[SumMes2] DOUBLE,[SumMes3] DOUBLE,[SumMeses] DOUBLE)"
    Else
      dbTemp.Execute "DROP TABLE tblRelVendasPorVendedor"
      dbTemp.Execute "CREATE TABLE tblRelVendasPorVendedor([Filial] INTEGER,[Vendedor] INTEGER,[DataIni1] DATE,[DataFim1] DATE,[DataIni2] DATE,[DataFim2] DATE,[DataIni3] DATE,[DataFim3] DATE,[Operacao] INTEGER,[Cliente] LONG,[SumMes1] DOUBLE,[SumMes2] DOUBLE,[SumMes3] DOUBLE,[SumMeses] DOUBLE)"
    End If
'  Set fd = td.CreateField("Filial", dbInteger)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("Vendedor", dbInteger)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("DataIni1", dbDate)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("DataFim1", dbDate)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("DataIni2", dbDate)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("DataFim2", dbDate)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("DataIni3", dbDate)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("DataFim3", dbDate)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("Operacao", dbInteger)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("Cliente", dbLong)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("SumMes1", dbDouble)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("SumMes2", dbDouble)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("SumMes3", dbDouble)
'  td.Fields.Append fd
'
'  Set fd = td.CreateField("SumMeses", dbDouble)
'  td.Fields.Append fd
'
'  dbTemp.TableDefs.Append td
'
'  Set td = Nothing
  
  gbCreateTableTBLRelVendasPorVendedor = True
  Exit Function
  
  End If
  
ErrCreate:
  gbCreateTableTBLRelVendasPorVendedor = False
End Function

Private Function gbIncluiCamposNFE3_10()
  '26/06/2013-Alexandre Afornali
  'Case lei De Olho nos Impostos
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("AliquotasNCM")
  
  Set fd = td.CreateField("Codigo", dbText, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("EX", dbText, 2)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Tabela", dbLong, 13)
  td.Fields.Append fd
  
  Set fd = td.CreateField("AliqNacional", dbDouble)
  td.Fields.Append fd
  
  Set fd = td.CreateField("AliqImportacao", dbDouble)
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableAliquotasNCM = True
  Exit Function
  
ErrCreate:
  gbCreateTableAliquotasNCM = False
End Function

Private Function gbCreateTableNFCE_ENVI()
  '26/06/2013-Alexandre Afornali
  'Case lei De Olho nos Impostos
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("NFCE_ENVI")
  
  Set fd = td.CreateField("CNPJ", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("ID", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Serie", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("N_NF", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("C_NF", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Chave", dbText, 100)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Detalhe_Autorizacao", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Detalhe_Cancelamento", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Dh_Autorizacao", dbText, 20)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Em_Contingencia", dbText, 1)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Ex_Message", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Numero", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Numero_Protocolo_Autorizacao", dbText, 100)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("O_Id", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Status_Autorizacao", dbText, 100)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Status_Cancelamento", dbText, 100)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("URL_QRCode", dbMemo)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Protocolo_Xml", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableNFCE_ENVI = True
  Exit Function
  
ErrCreate:
  gbCreateTableNFCE_ENVI = False
End Function

Private Function gbCreateTableNFCE_job()
  '26/06/2013-Alexandre Afornali
  'Case lei De Olho nos Impostos
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("NFCE_job")
  
  Set fd = td.CreateField("CNPJ", dbText)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Xml", dbMemo)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Tipo", dbText)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Serie", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("N_NF", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Chave", dbText, 100)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Cupom", dbMemo)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Justificativa", dbText, 100)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Processado", dbText)
  fd.DefaultValue = "N"
  td.Fields.Append fd
  
  Set fd = td.CreateField("CPF", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome_Consumidor", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Data_Emissao", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Total_Tributos", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Nome_Emitente", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Endereco_Emitente", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("IE_Emitente", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableNFCE_job = True
  Exit Function
  
ErrCreate:
  gbCreateTableNFCE_job = False
End Function

Private Function gbCreateTableCupom_temp()
  '26/06/2013-Alexandre Afornali
  'Case lei De Olho nos Impostos
  Dim td As TableDef
  Dim fd As Field
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Cupom_temp")
  
  Set fd = td.CreateField("N_NF", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Serie", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("CNPJ", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Codigo", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Descricao", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Qtd", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("Un", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("vl_unit", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set fd = td.CreateField("vl_total", dbText)
  fd.AllowZeroLength = True
  td.Fields.Append fd

  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableCupom_temp = True
  Exit Function
  
ErrCreate:
  gbCreateTableCupom_temp = False
End Function

Private Function gbCreateTableRef_CEST_NCM() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim sSql As String
  Dim rs As Recordset
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("Ref_CEST_NCM")
  Set fd = td.CreateField("cest", dbText, 8)
  td.Fields.Append fd
  Set fd = td.CreateField("ncm", dbText, 7)
  td.Fields.Append fd
  
  db.TableDefs.Append td
  
  Set td = Nothing
 
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000100', '2522');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000200', '3816001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000200', '3824500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000300', '3214900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000400', '391000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000500', '3916');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000600', '3917');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000700', '3918');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000800', '3919');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000900', '3919');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000900', '3920');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1000900', '3921');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100100', '3815121');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100100', '3815129');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001000', '3921');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001100', '3921');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001200', '3921');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001300', '3922');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001400', '3924');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001500', '3925100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001600', '392590');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001700', '3925100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001700', '392590');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001800', '3925200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1001900', '3925300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100200', '3917');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002000', '392690');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002100', '4814');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002200', '6810190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002300', '6811');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002400', '6811');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002500', '6901000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002600', '6902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002700', '6904');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002800', '6905');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1002900', '6906000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100300', '3918100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003000', '6907');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003000', '6908');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003001', '6907');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003001', '6908');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003100', '6910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003200', '6912000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003300', '7003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003400', '7004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003500', '7005');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003600', '7007190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003700', '7007290');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003800', '7008');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1003900', '7016');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100400', '3923300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004000', '7214200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004100', '7308901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004200', '7214200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004300', '7213');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004300', '7308901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004400', '7217109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004400', '7312');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004500', '721720');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004600', '7307');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004700', '7308300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004800', '7308400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004800', '730890');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1004900', '7308400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100500', '3926300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005000', '7308909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005100', '7310');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005200', '7313000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005300', '7314');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005400', '7315110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005500', '7315129');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005600', '7315820');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005700', '731700');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005800', '7318');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1005900', '7323');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100600', '40103');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100600', '5910000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006000', '7324');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006100', '7325');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006200', '7326');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006300', '7407');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006400', '7411101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006500', '7412');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006600', '7415');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006700', '7418200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006800', '7607199');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1006900', '7608');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100700', '4016930');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100700', '4823909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007000', '7609000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007100', '7610');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007200', '7615200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007300', '7616');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007400', '8302410');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007500', '8301');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007600', '8302100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007700', '8307');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007800', '8311');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1007900', '8481');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100800', '4016101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100900', '4016999');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('100900', '5705000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101000', '5903900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101100', '5909000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101200', '63061');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101300', '6506100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101400', '6813');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101500', '7007110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101500', '7007210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101600', '7009100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101700', '7014000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101800', '7311000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('101900', '7311000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102000', '7320');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102100', '7325');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102200', '780600');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102300', '8007009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102400', '830120');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102400', '830160');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102500', '830170');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102600', '8302100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102600', '8302300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102700', '831000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102800', '84073');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('102900', '840820');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103000', '84099');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103100', '84122');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103200', '841330');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103300', '8414100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103400', '8414801');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103400', '8414802');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103500', '8413919');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103500', '8414901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103500', '8414903');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103500', '8414903');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103600', '841520');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103700', '8421230');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103800', '8421299');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('103900', '84219');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104000', '8424100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104100', '8421310');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104200', '8421392');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104300', '8425420');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104400', '8431101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104500', '8431492');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104500', '8433909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104600', '8481100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104700', '84812');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104800', '8481809');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('104900', '8482');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105000', '8483');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105100', '8484');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105200', '850520');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105300', '8507100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105400', '8511');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105500', '851220');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105500', '851240');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105500', '8512900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105600', '8517121');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105700', '8518');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105800', '8518500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('105900', '851981');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106000', '8525501');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106000', '8525601');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106100', '85272');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106200', '8527219');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106200', '8521909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106300', '8529109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106400', '8534000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106500', '853530');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106500', '853650');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106600', '8536100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106700', '8536200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106800', '85364');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('106900', '8538');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107000', '8536509');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107100', '853910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107200', '85392');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107300', '8544200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107400', '8544300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107500', '8707');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107600', '8708');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107700', '87141');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107800', '8716909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('107900', '902610');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108000', '902620');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108100', '9029');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108200', '9030332');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108300', '9031804');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108400', '9032892');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108500', '9104000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108600', '9401200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108600', '9401909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108700', '9613800');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108800', '4009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108900', '4504900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('108900', '6812991');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109000', '4823400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109100', '3919100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109100', '3919900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109100', '8708299');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109200', '8412311');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109300', '8413190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109300', '8413509');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109300', '8413810');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109400', '8413601');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109400', '8413701');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109500', '8414591');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109500', '8414599');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109600', '8421399');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109700', '8501101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109800', '8501311');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('109900', '8504500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110000', '850720');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110000', '850730');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100100', '2828901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100100', '2828901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100100', '3206410');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100100', '3808941');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100200', '3401209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100300', '3401209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100400', '3402200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100500', '3402200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100600', '3402200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100700', '3402');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100800', '3809919');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100900', '3924100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100900', '3924900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100900', '6805301');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1100900', '6805309');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110100', '8512300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1101000', '2207');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1101100', '7323100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110200', '9032898');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110200', '9032899');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110300', '9027100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110400', '4008110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110500', '5601221');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110600', '5703200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110700', '5703300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110800', '5911900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('110900', '6903909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111000', '7007290');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111100', '7314500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111200', '7315110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111300', '7315121');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111400', '8418990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111500', '841950');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111600', '8424909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111700', '8425491');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111800', '8431410');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('111900', '8501610');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112000', '8531109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112100', '9014100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112200', '9025199');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112300', '9025901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112400', '902690');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112500', '9032101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112600', '9032109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112700', '9032200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('112800', '871690');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200100', '8504');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200200', '8516');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200300', '8535');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200400', '8536');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200500', '8538');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200600', '7413000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200700', '8544');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200700', '7605');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200700', '7614');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200800', '8546');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1200900', '8547');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300100', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300100', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300101', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300101', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300102', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300102', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300200', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300200', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300201', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300201', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300202', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300202', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300300', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300300', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300301', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300301', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300302', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300302', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300400', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300400', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300401', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300401', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300402', '3003');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300402', '3004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300500', '300660');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300600', '2936');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300700', '300630');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300800', '3002');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1300900', '3005');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1301000', '4015110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1301000', '4015190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1301100', '4014100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1301200', '901831');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1301300', '9018321');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1301400', '3926909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1301400', '9018909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1400100', '4823209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1400200', '48236');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1400300', '4813100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1500100', '3919');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1500100', '3920');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1500100', '3921');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1500200', '3924');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1500300', '3924100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600100', '4011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600200', '4011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600300', '4011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600400', '4011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600500', '4011500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600600', '40121');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600700', '401290');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600701', '401290');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600800', '4013');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1600900', '4013200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700100', '1704901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700200', '1806311');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700200', '1806312');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700300', '1806321');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700300', '1806322');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700400', '1806900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700500', '1806900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700600', '1806900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700700', '1704909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700800', '1806900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700900', '210120');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1700900', '2202900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701000', '2202100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701100', '2202900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701200', '2009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701300', '20098');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701400', '2202900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701500', '2202900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701600', '2202100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701700', '4021');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701700', '4022');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701700', '4029');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701800', '1901102');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1701900', '1901101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702000', '1901109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702000', '1901103');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702100', '4011010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702100', '4012010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702101', '4011010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702101', '4012010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702200', '4014010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702200', '4015010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702201', '4014010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702201', '4015010');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702300', '4011090');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702300', '4012090');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702301', '4011090');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702301', '4012090');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702400', '401402');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702400', '4022130');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702400', '4022930');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702400', '4029');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702401', '401402');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702401', '4022130');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702401', '4022930');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702401', '4029');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702402', '40110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702402', '40120');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702402', '40150');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702402', '40210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702402', '4022920');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702500', '4029');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702501', '4029');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702600', '403');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702601', '403');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702700', '4039000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702800', '406');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702801', '406');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1702900', '406');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703000', '4051000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703001', '4051000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703100', '1517100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703200', '1517100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703201', '1517100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703202', '151790');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703300', '1516200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703301', '1516200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703400', '1901902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703500', '1904100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703500', '1904900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703600', '1905909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703700', '2005200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703700', '20059');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703800', '20081');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703801', '20081');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1703900', '2103201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704000', '2103902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704000', '2103909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704100', '2103101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704200', '2103301');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704300', '2103302');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704400', '2103901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704500', '2002');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704600', '2103201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704700', '1704909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704700', '1904200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704700', '1904900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704800', '1806312');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704800', '1806322');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704800', '1806900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704900', '1101001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1704901', '1101001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705000', '1101002');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705100', '1901200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705100', '1901909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705200', '1902300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705300', '1902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705301', '1902400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705400', '19021');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705500', '190520');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705600', '1905209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705700', '1905201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705800', '190531');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1705900', '190531');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706000', '190531');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706100', '1905902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706200', '190532');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706300', '190532');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706400', '190540');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706500', '1905901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706600', '1905902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706700', '1905909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706800', '1905100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1706900', '1905909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707000', '190590');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707100', '1507901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707200', '1508');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707300', '1509');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707301', '1509');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707302', '1509');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707400', '1510000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707500', '1512191');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707500', '1512291');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707600', '15141');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707700', '1515190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707800', '1515291');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1707900', '1512299');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708000', '1517901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708100', '1511');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708100', '1513');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708100', '1514');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708100', '1515');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708100', '1516');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708100', '1518');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708200', '1601000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708300', '1601000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708400', '1601000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708500', '1602');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708600', '1604');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708700', '1604');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708800', '1605');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708900', '201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708900', '202');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708900', '204');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708900', '206');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708900', '2102000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708900', '2109900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1708900', '1502');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709000', '204');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709100', '203');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709100', '206');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709100', '207');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709100', '209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709100', '2101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709100', '2109900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709100', '1501');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709200', '710');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709201', '710');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709300', '811');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709301', '811');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709400', '2001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709401', '2001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709500', '2004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709501', '2004');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709600', '2005');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709601', '2005');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709700', '2006000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709701', '2006000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709800', '2007');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709801', '2007');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709900', '2008');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1709901', '2008');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710000', '901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710001', '901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710100', '902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710100', '1211909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710100', '2106909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710200', '90300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710300', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710300', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710301', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710301', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710302', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710302', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710400', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710401', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710402', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710500', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710500', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710501', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710501', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710502', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710502', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710600', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710601', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710602', '170191');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710700', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710700', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710701', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710701', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710702', '17011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710702', '1701990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710800', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710801', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710802', '1701910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710900', '1702');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710901', '1702');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1710902', '1702');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1711000', '2008190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1711100', '21011');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1711200', '210120');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1711300', '1901909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1711300', '2101119');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1711300', '2101120');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1800100', '6911101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1800200', '6911109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1800300', '6912000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1800400', '6912000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900100', '3213100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900200', '3916200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900300', '3926100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900400', '42021');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900400', '42029');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900500', '3926909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900600', '4802209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900600', '4811909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900700', '4802549');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900800', '4802549');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900800', '4802579');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900800', '4816200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900900', '4802569');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900900', '4802579');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1900900', '4802589');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901000', '3703101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901000', '3703102');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901000', '3703200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901000', '3703901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901000', '3704000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901000', '4802200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901100', '4810139');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901200', '4816100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901300', '3920201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901400', '4806200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901500', '4808100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901600', '4810229');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901700', '4809');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901700', '4816');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901800', '4817');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1901900', '4820100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902000', '4820200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902100', '4820300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902200', '4820400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902300', '4820500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902400', '4820900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902500', '4909000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902600', '9608100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902700', '9608200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902800', '9608300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1902900', '9608');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('1903000', '480256');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000100', '1211909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000101', '1211909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000200', '2712100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000300', '2814200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000400', '2847000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000500', '3006700');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000600', '3301');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000700', '3303001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000800', '3303002');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2000900', '3304100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200100', '2205');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200100', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001000', '3304201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001100', '3304209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001200', '3304300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001300', '3304910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001400', '3304991');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001500', '3304999');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001600', '3304999');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001700', '3305100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001800', '3305200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2001900', '3305300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200200', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002000', '3305900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002100', '3305900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002200', '3305900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002300', '3306100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002400', '3306200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002500', '3306900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002600', '3307100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002700', '3307201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002800', '3307201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2002900', '3307209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200300', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003000', '3307209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003100', '3307300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003200', '3307900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003300', '3307900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003400', '3401119');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003500', '3401190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003600', '3401201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003700', '3401300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003800', '4014901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003900', '4014909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2003901', '3926904');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200400', '220720');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200400', '2208400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004000', '42021');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004100', '4818100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004200', '4818100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004300', '4818200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004400', '4818200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004500', '4818300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004600', '4818909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004700', '9619000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004800', '9619000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2004900', '9619000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200500', '2205');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200500', '2206009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200500', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005000', '5601219');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005100', '5603929');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005200', '8203209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005300', '8214100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005400', '8214200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005500', '9025111');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005500', '9025199');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005600', '96032');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005700', '9603210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005800', '9603300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2005900', '9605000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200600', '2208200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006000', '9615');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006100', '9616200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006200', '3923300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006200', '3924900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006200', '3924100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006200', '4014909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006200', '7010200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006300', '8212102');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2006300', '8212201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200700', '2206009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200700', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200800', '2208500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200900', '2205');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200900', '2206009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('200900', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201000', '2208700');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201100', '2208200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201200', '2208400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201300', '2206009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201400', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201500', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201600', '220830');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201700', '2205');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201800', '2208600');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('201900', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202000', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202100', '2208200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202200', '2206001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202300', '2205');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202300', '2206009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202300', '2208900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202400', '2204');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202500', '2204');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202500', '2205');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202500', '2206');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202500', '2207');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('202500', '2208');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100100', '7321110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100100', '7321810');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100100', '7321900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100200', '8418100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100300', '8418210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100400', '8418290');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100500', '8418300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100600', '8418400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100700', '841850');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100800', '8418699');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2100900', '8418699');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101000', '8418990');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101100', '842112');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101200', '8421199');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101300', '8418693');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101400', '84219');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101500', '8422110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101500', '8422901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101600', '844331');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101700', '844332');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101800', '844399');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2101900', '8450110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102000', '8450120');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102100', '8450190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102200', '845020');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102300', '845090');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102400', '8451210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102500', '8451299');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102600', '845190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102700', '8452100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102800', '847130');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2102900', '84714');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103000', '8471501');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103100', '8471605');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103200', '8471609');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103300', '847170');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103400', '847190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103500', '847330');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103600', '85043');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103700', '8504401');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103800', '8504404');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2103900', '8507800');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104000', '8508');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104100', '8509');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104200', '8509801');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104300', '8516100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104400', '8516400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104500', '8516500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104600', '8516600');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104700', '8516600');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104800', '8516710');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2104900', '8516720');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105000', '851679');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105100', '8516900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105200', '8517110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105300', '851712');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105400', '8517189');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105500', '8517625');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105600', '8518');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105700', '8519');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105700', '8522');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105700', '85271');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105800', '8519819');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2105900', '8521901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106000', '8521909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106100', '8523511');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106200', '8523520');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106300', '8525802');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106400', '85279');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106500', '8528492');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106500', '8528592');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106500', '852869');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106600', '8528512');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106700', '85287');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106800', '85287');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2106900', '85287');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107000', '85287');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107100', '85287');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107200', '900610');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107300', '9006400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107400', '9018905');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107500', '9019100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107600', '9032891');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107700', '9504500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107800', '8517621');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2107900', '8517622');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108000', '8517623');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108100', '8517624');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108200', '8517626');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108300', '8517629');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108400', '8517702');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108500', '821490');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108500', '8510');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108600', '84145');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108700', '8414599');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108800', '8414600');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2108900', '8414902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109000', '841510');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109000', '84158');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109100', '8415101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109200', '8415101');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109300', '8415109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109400', '8415901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109500', '8415902');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109600', '8421210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109700', '8424301');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109700', '8424309');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109700', '8424909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109800', '8467210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2109900', '85162');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110000', '8516310');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110100', '8516320');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110200', '8518');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110300', '8527');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110400', '8521909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110500', '8479600');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110600', '8415909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110700', '8525801');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110800', '8423100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2110900', '8540');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111000', '8517');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111100', '8517');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111200', '8529');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111300', '8531');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111400', '853110');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111500', '8531800');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111600', '853400');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111700', '8541401');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111700', '8541402');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111700', '8541402');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111800', '8543709');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2111900', '90303');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2112000', '903089');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2112100', '910700');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2112200', '9405');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2200100', '2309');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2300100', '210500');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2300200', '1806');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2300200', '1901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2300200', '2106');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2400100', '3208');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2400100', '3209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2400100', '3210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2400200', '2821');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2400200', '3204170');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2400200', '3206');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500100', '8702100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500200', '8702909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500300', '8703210');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500400', '8703221');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500500', '8703229');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500600', '8703231');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500700', '8703239');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500800', '8703241');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2500900', '8703249');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501000', '8703321');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501100', '8703329');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501200', '8703331');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501300', '8703339');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501400', '8704211');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501500', '8704212');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501600', '8704213');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501700', '8704219');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501800', '8704311');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2501900', '8704312');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2502000', '8704313');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2502100', '8704319');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2600100', '8711');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2700100', '7009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2700200', '7013');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2700300', '7013370');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2700400', '7013429');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800100', '3303001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800200', '3303002');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800300', '3304100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800400', '3304201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800500', '3304209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800600', '3304300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800700', '3304910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800800', '3304991');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2800900', '3304999');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801000', '3304999');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801100', '3305100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801200', '3305200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801300', '3305900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801400', '3305900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801500', '3307100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801600', '3307201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801700', '3307209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801800', '3307900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2801900', '3307900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802000', '3401119');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802100', '3401190');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802200', '3401201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802300', '3401300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802400', '4818200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802500', '8214100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802600', '8214200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802700', '9603290');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802800', '9603300');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2802900', '9616100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803000', '9616200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803100', '42021');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803200', '9615');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803300', '3924100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803300', '3924900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803300', '4014909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803400', '4014909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803500', '33');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803500', '34');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803600', '44');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803600', '64');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803600', '65');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803600', '82');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803600', '90');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803600', '96');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803700', '39');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803700', '48');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803700', '91');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803700', '42');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803700', '71');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803700', '83');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803700', '90');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803800', '61');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803800', '62');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803800', '64');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803900', '42');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803900', '52');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803900', '55');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803900', '58');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803900', '63');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2803900', '65');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '39');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '40');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '56');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '63');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '66');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '69');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '70');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '73');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '82');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '83');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '84');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '91');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '94');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804000', '96');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804100', '13');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804100', '15');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804100', '23');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804200', '33');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '22');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '27');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '28');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '29');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '33');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '34');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '35');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '38');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '39');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '63');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '68');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '73');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '84');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '85');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('2804300', '86');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300100', '2201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300200', '2201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300300', '2201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300400', '2201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300500', '2201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300600', '2201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300600', '2202');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300700', '2202');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300800', '2202');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('300900', '2106901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('301000', '2202900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('301100', '210690');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('301200', '2203000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('301300', '2202900');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('301400', '2203000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('400100', '2402');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('400200', '24031');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('500100', '2523');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600100', '220710');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600200', '2710125');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600300', '2710191');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600400', '2710192');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600500', '2710193');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600600', '2710199');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600700', '27109');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600800', '2711');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('600900', '2713');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('601000', '3826000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('601100', '3403');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('601200', '2710200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('700100', '2716000');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800100', '4016999');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800200', '4417001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800200', '4417009');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800300', '6804');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800400', '8201');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800500', '8202200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800600', '8202910');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800700', '8202');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800800', '8203');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('800900', '8204');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801000', '8205');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801100', '8206');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801200', '820740');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801200', '820760');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801200', '820770');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801300', '8207');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801400', '8208');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801500', '8209001');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801600', '8209');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801700', '8211');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801800', '8213');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('801900', '8467');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802000', '9015');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802100', '9017200');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802100', '901730');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802100', '901780');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802100', '9017909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802200', '9025119');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802200', '9025901');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802300', '902519');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('802300', '9025909');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('900100', '8539');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('900200', '8540');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('900300', '8504100');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('900400', '853650');", dbFailOnError
  db.Execute "INSERT INTO [Ref_CEST_NCM] (cest, ncm) VALUES ('900500', '8543709');", dbFailOnError
  
  gbCreateTableRef_CEST_NCM = True
  Exit Function
  
ErrCreate:
  gbCreateTableRef_CEST_NCM = False
  
End Function


Private Function gbAlterTableAliquotasNCM()
  On Error GoTo ErrCreate
  
  If db.TableDefs("AliquotasNCM").Fields("Codigo").Type <> dbText Then
     MsgBox "A tabela AliquotasNCM está com o tipo de campo para Codigo incorreto. Entre urgente em contato com o suporte."
     gbAlterTableAliquotasNCM = False
     Exit Function
  End If
  
  gbAlterTableAliquotasNCM = True
  Exit Function
  
ErrCreate:
  gbAlterTableAliquotasNCM = False
End Function

Private Function p_blnCreateTableProdutoNomeNFe() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("ProdutoNomeNFe")
  
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Sequencia", dbLong)
  td.Fields.Append fd
  
  Set fd = td.CreateField("Codigo", dbText, 20)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  Set fd = td.CreateField("Nome", dbText, 80)
  fd.AllowZeroLength = False
  td.Fields.Append fd
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  p_blnCreateTableProdutoNomeNFe = True
  Exit Function
  
ErrCreate:
  p_blnCreateTableProdutoNomeNFe = False
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

