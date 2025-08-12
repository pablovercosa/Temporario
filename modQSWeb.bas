Attribute VB_Name = "modQSWeb"
Option Explicit

'Tabela de preços para os produtos de pedidos da Loja Virtual
'onde TQW significa Tabela Quick Web
'e ____________ (12 underscores) é o local
'onde será substituído pelo número do pedido formatado com zeros
'totalizando 15 casas que é o limite para tabela no banco de dados
'Ex.: TQW000000000001
Public Const LIST_PRICE_WEB As String = "TQW____________"
'Texto as ser substituído na constante LIST_PRICE_WEB
Public Const REPLACE_TQW As String = "____________"

'Tipos de status (passo) para os pedidos da Loja Virtual
Public Enum enWEB_OrderFormStep
  ofsAll = -1
  ofsReceived = 0
  ofsConfirmedPayment = 10
  ofsPacked = 15
  ofsHasSent = 30
  ofsCanceled = 90
End Enum

'-------------------------------------------------------------------------------
'26/04/2002 - mpdea
'
'Ajustes gerais para implantação da função e suas subfunções
'-------------------------------------------------------------------------------
Public Function AlteraDBWeb()
  Dim nStep As Integer
  Dim lngX As Long
  
  On Error GoTo ErrHandler
  
  '1. Inclusão de campo em Cli_For
  nStep = nStep + 1
  If gbGetField("Cli_For", "WebShopperID") = False Then
    If gbCreateFieldZeroLenght("Cli_For", "WebShopperID", dbText, 32) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      GoTo ErrInStep
    End If
  End If

  '2. Inclusão de campo em Cli_For
  nStep = nStep + 1
  If gbGetField("Cli_For", "WebDataCadastro") = False Then
    If gbCreateField("Cli_For", "WebDataCadastro", dbDate) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '3. Inclusão de campo em Cli_For
  nStep = nStep + 1
  If gbGetField("Cli_For", "WebCountry") = False Then
    If gbCreateFieldZeroLenght("Cli_For", "WebCountry", dbText, 50) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      GoTo ErrInStep
    End If
  End If

  '4. Inclusão de campo em Cli_For
  nStep = nStep + 1
  If gbGetField("Cli_For", "DataNascimento") = False Then
    If gbCreateField("Cli_For", "DataNascimento", dbDate) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '5. Inclusão de campo em Cli_For
  nStep = nStep + 1
  If gbGetField("Cli_For", "WebBonus") = False Then
    'Faltou valor default = 0
    If gbCreateField("Cli_For", "WebBonus", dbLong) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '6. Alteração do tamanho do campo Nome em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampoIndex("Cli_For", "Nome", dbText, 100, "Nome", "Nome", "Código", True, False) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
   
  '7. Alteração do tamanho do campo Endereço em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampo("Cli_For", "Endereço", dbText, 200) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
  
  '8. Alteração do tamanho do campo Estado em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampo("Cli_For", "Estado", dbText, 40) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
  
  '9. Inclusão de campo em Cli_For
  nStep = nStep + 1
  If gbGetField("Cli_For", "Sexo") = False Then
    If gbCreateField("Cli_For", "Sexo", dbText, 1) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Cli_For"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '10. Alteração do tamanho do campo e-mail em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampo("Cli_For", "email", dbText, 100) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
  
  '11. Alteração do tamanho do campo Cep em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampo("Cli_For", "CEP", dbText, 15) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
  
  '12. Alteração do tamanho do cidade Nome em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampoIndex("Cli_For", "Cidade", dbText, 50, "Cidade", "Cidade", "Código", False, False) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
  
  '13. Alteração do tamanho do fone1  em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampoIndex("Cli_For", "Fone 1", dbText, 35, "Telefone", "Fone 1", "Código", False, False) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
  
  '14. Alteração do tamanho do CGC em Cli_For
  nStep = nStep + 1
  If gbAlteraTamanhoCampoIndex("Cli_For", "CGC", dbText, 20, "CGC", "CGC", "Código", False, False) = False Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Cli_For"" - não foi possível."
    GoTo ErrInStep
  End If
  
  '15. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebIncluded") = False Then
    If gbCreateField("Produtos", "WebIncluded", dbBoolean) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '16. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebSynchronize") = False Then
    'Faltou valor default = True
    If gbCreateField("Produtos", "WebSynchronize", dbBoolean) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '17. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebLastOp") = False Then
    If gbCreateFieldZeroLenght("Produtos", "WebLastOp", dbText, 1) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If

  '18. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebBonus") = False Then
    If gbCreateField("Produtos", "WebBonus", dbLong) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If

  '19. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebOfferDateStart") = False Then
    If gbCreateField("Produtos", "WebOfferDateStart", dbDate) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If

  '20. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebOfferDateEnd") = False Then
    If gbCreateField("Produtos", "WebOfferDateEnd", dbDate) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If

  '21. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebOfferTablePrice") = False Then
    If gbCreateFieldZeroLenght("Produtos", "WebOfferTablePrice", dbText, 15) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
 
  '22. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebSaleTablePrice") = False Then
    If gbCreateField("Produtos", "WebSaleTablePrice", dbText, 15) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '--------------------------------------------------------------------------------
  '10/10/2002 - mpdea
  'Removido a criação do campo de configuração da tabela de custo
  'Fixada como CUSTO no Quick Web
'  '23. Inclusão de campo em Produtos
  nStep = nStep + 1
'  If gbGetField("Produtos", "WebCostTablePrice") = False Then
'    If gbCreateField("Produtos", "WebCostTablePrice", dbText, 15) = False Then
'      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
'      GoTo ErrInStep
'    End If
'  End If
  
  'Remove campo de configuração da tabela de custo
  If gbGetField("Produtos", "WebCostTablePrice") Then
    db.TableDefs("Produtos").Fields.Delete "WebCostTablePrice"
  End If
  '--------------------------------------------------------------------------------
  
  
  '24. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebAttribFabricante") = False Then
    If gbCreateField("Produtos", "WebAttribFabricante", dbBoolean) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '--------------------------------------------------------------------------------
  '22/08/2002 - mpdea
  'Removido a criação do campo de atributo único 'Unidade de Venda'
'  '25. Inclusão de campo em Produtos
  nStep = nStep + 1
'  If gbGetField("Produtos", "WebAttribUnidVenda") = False Then
'    If gbCreateField("Produtos", "WebAttribUnidVenda", dbBoolean) = False Then
'      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
'      GoTo ErrInStep
'    End If
'  End If
  
  '20/09/2002 - mpdea
  'Para manter compatibilidade com versões já jançadas remove campo caso exista
  If gbGetField("Produtos", "WebAttribUnidVenda") Then
    db.TableDefs("Produtos").Fields.Delete "WebAttribUnidVenda"
  End If
  '--------------------------------------------------------------------------------
  
  
  '26. Inclusão de campo em Produtos
  nStep = nStep + 1
  If gbGetField("Produtos", "WebAttribPesquisa123") = False Then
    If gbCreateField("Produtos", "WebAttribPesquisa123", dbBoolean) = False Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Produtos"" - não foi possível."
      GoTo ErrInStep
    End If
  End If

  '27. Criação da Tabela WEB_Config
  nStep = nStep + 1
  If gbGetTable("WEB_Config") = False Then
    If gbCreateTableWEB_Config() = False Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_Config"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '28. Criação da Tabela WEB_ProdutosExcluir
  nStep = nStep + 1
  If gbGetTable("WEB_ProdutosExcluir") = False Then
    If gbCreateTableWEB_ProdutosExcluir() = False Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_ProdutosExcluir"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '29 Criação da Tabela WEB_OrderForms
  nStep = nStep + 1
  If gbGetTable("WEB_OrderForms") = False Then
    If gbCreateTableWEB_OrderForms() = False Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_OrderForms"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '30 Criação da Tabela WEB_OrderItens
  nStep = nStep + 1
  If gbGetTable("WEB_OrderItens") = False Then
    If gbCreateTableWEB_OrderItens() = False Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_OrderItens"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '31 Criação da Tabela WEB_OrderStatus
  nStep = nStep + 1
  If gbGetTable("WEB_OrderStatus") = False Then
    If gbCreateTableWEB_OrderStatus() = False Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_OrderStatus"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '32 Criação da Tabela WEB_PayamentMethods
  nStep = nStep + 1
  If gbGetTable("WEB_PaymentMethods") = False Then
    If gbCreateTableWEB_PaymentMethods() = False Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_PaymentMethods"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  '33 Criação da Tabela WEB_ShippingMethods
  nStep = nStep + 1
  If gbGetTable("WEB_ShippingMethods") = False Then
    If gbCreateTableWEB_ShippingMethods() = False Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_ShippingMethods"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '34. Inclusão de campo na tabela de Saídas
  nStep = nStep + 1
  If Not gbGetField("Saídas", "WebOrderFormID") Then
    'Faltou valor default = 0
    If Not gbCreateField("Saídas", "WebOrderFormID", dbLong) Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Saídas"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '35. Criação da Tabela WEB_OrderStatusHistoric
  nStep = nStep + 1
  If Not gbGetTable("WEB_OrderStatusHistoric") Then
    If Not gbCreateTableWEB_OrderStatusHistoric() Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_OrderStatusHistoric"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '36. Inclusão de campo na tabela de Entradas
  nStep = nStep + 1
  If Not gbGetField("Entradas", "WebOrderFormID") Then
    'Faltou valor default = 0
    If Not gbCreateField("Entradas", "WebOrderFormID", dbLong) Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Entradas"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '37. Inclusão de campo na tabela de Op. de Entrada
  nStep = nStep + 1
  If Not gbGetField("Operações Entrada", "Locked") Then
    If Not gbCreateField("Operações Entrada", "Locked", dbBoolean) Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações de Entrada"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '38. Inclusão de campo na tabela de Op. de Saída
  nStep = nStep + 1
  If Not gbGetField("Operações Saída", "Locked") Then
    If Not gbCreateField("Operações Saída", "Locked", dbBoolean) Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""Operações de Saída"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '39. Inclusão do registro de operação de Entrada para as operações WEB
  nStep = nStep + 1
  If CInt("0" & gvGetValueInTable("WEB_Config", "CodOpCancelamento", ftNumero, "ID", ftNumero, "1")) = 0 Then
    'Próximo código livre
    lngX = glngNextFreeCode("Operações Entrada", "Código", 1, 999)
    If lngX = -1 Then 'Nenhum código livre
      gsMsg = "Nenhum código livre para a criação da Operação de Entrada."
      GoTo ErrInStep
    Else
      'Operação de entrada para cancelamento de pedido WEB
      db.Execute "INSERT INTO [Operações Entrada] " & _
        "(Código, Nome, Tipo, Estoque, Locked) " & _
        "VALUES (" & lngX & _
        ", 'Pedido Cancelado da Loja Virtual', 'D', True, True)", _
        dbFailOnError
      'Atualiza tabela de configurações
      db.Execute "UPDATE WEB_Config SET CodOpCancelamento = " & lngX & " WHERE ID = 1", dbFailOnError
    End If
  End If
  
  
  '40. Inclusão do registro de operação de Saída para as operações WEB
  nStep = nStep + 1
  If CInt("0" & gvGetValueInTable("WEB_Config", "CodOpReserva", ftNumero, "ID", ftNumero, "1")) = 0 Then
    'Próximo código livre
    lngX = glngNextFreeCode("Operações Saída", "Código", 1, 999)
    If lngX = -1 Then 'Nenhum código livre
      gsMsg = "Nenhum código livre para a criação da Operação de Saída."
      GoTo ErrInStep
    Else
      'Operação de saída para ajuste/reserva de estoque do pedido WEB
      db.Execute "INSERT INTO [Operações Saída] " & _
        "(Código, Nome, Tipo, Estoque, Locked) " & _
        "VALUES (" & lngX & _
        ", 'Reserva de Estoque para Pedido da Loja Virtual', 'V', True, True)", _
        dbFailOnError
      'Atualiza tabela de configurações
      db.Execute "UPDATE WEB_Config SET CodOpReserva = " & lngX & " WHERE ID = 1", dbFailOnError
    End If
  End If
  
  
  '41. Inclusão do registro de operação de Saída para as operações WEB
  nStep = nStep + 1
  If CInt("0" & gvGetValueInTable("WEB_Config", "CodOpVenda", ftNumero, "ID", ftNumero, "1")) = 0 Then
    'Próximo código livre
    lngX = glngNextFreeCode("Operações Saída", "Código", 1, 999)
    If lngX = -1 Then 'Nenhum código livre
      gsMsg = "Nenhum código livre para a criação da Operação de Saída."
      GoTo ErrInStep
    Else
      'Operação de saída para confirmação do pedido WEB
      db.Execute "INSERT INTO [Operações Saída] " & _
        "(Código, Nome, Tipo, Dinheiro, Comissão, Nota, " & _
        "[Soma Frete], InTelaObsTransp, ICM, Locked) " & _
        "VALUES (" & lngX & ", 'Venda da Loja Virtual', 'V', " & _
        "True, True, True, True, True, True, True)", _
        dbFailOnError
      'Atualiza tabela de configurações
      db.Execute "UPDATE WEB_Config SET CodOpVenda = " & lngX & " WHERE ID = 1", dbFailOnError
    End If
  End If
  
  
  '02/12/2002 - mpdea
  '42. Inclusão de campo para permissão na exportação de produtos sem descrição
  nStep = nStep + 1
  If Not gbGetField("WEB_Config", "AllowExp_SemDescricao") Then
    If Not gbCreateField("WEB_Config", "AllowExp_SemDescricao", dbBoolean) Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""WEB_Config"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '02/12/2002 - mpdea
  '43. Inclusão de campo para unidade de venda padrão para os produtos
  '    que não estejam configurados como 'kg', 'k' ou 'g'
  nStep = nStep + 1
  If Not gbGetField("WEB_Config", "UnitSaleDefault") Then
    If Not gbCreateField("WEB_Config", "UnitSaleDefault", dbText, 1) Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""WEB_Config"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '27/05/2003 - mpdea
  '44. Inclusão de campo para escolha do tipo de classificação
  '    para exportação dos produtos (somente classe ou classe com sub classes)
  nStep = nStep + 1
  If Not gbGetField("WEB_Config", "ExportWithClasseAndSubClasse") Then
    If Not gbCreateField("WEB_Config", "ExportWithClasseAndSubClasse", dbBoolean) Then
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""WEB_Config"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '12/05/2004 - mpdea
  '45. Inclusão de campo para configuração da conexão em modo passivo
  nStep = nStep + 1
  If Not gbGetField("WEB_Config", "PassiveMode") Then
    If gbCreateField("WEB_Config", "PassiveMode", dbBoolean) Then
      db.Execute "UPDATE WEB_Config SET PassiveMode = TRUE;"
    Else
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""WEB_Config"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '21/05/2004 - mpdea
  '46. Alterações na tabela WEB_OrderForms
  '    Verifica a existência de campo inicial das alterações
  nStep = nStep + 1
  If Not gbGetField("WEB_OrderForms", "Comentario") Then
    If Not gbChangeTableWEB_OrderForms Then
      gsMsg = "Manutenção na Base de Dados: Alteração da tabela ""WEB_OrderForms"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
    
  '24/05/2004 - mpdea
  '47. Criação da Tabela WEB_ClienteOrigem
  nStep = nStep + 1
  If Not gbGetTable("WEB_ClienteOrigem") Then
    If Not gbCreateTableWEB_ClienteOrigem() Then
      gsMsg = "Manutenção na Base de Dados: Adição da tabela ""WEB_ClienteOrigem"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '26/05/2004 - mpdea
  '48. Alterações na tabela Cli_For
  '    Verifica a existência de campo inicial das alterações
  nStep = nStep + 1
  If Not gbGetField("Cli_For", "Web") Then
    If Not gbChangeTableCliFor Then
      gsMsg = "Manutenção na Base de Dados: Alteração da tabela ""Clientes / Fornecedores"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  '----------------------------------------------------------------------------------
  '22/06/2004 - mpdea
  '26/05/2004 - mpdea
  'Alterações na tabela Cli_For
  '
  '49. Alterado o tamanho do campo Complemento de 15 para 50
  nStep = nStep + 1
  If Not gbAlteraTamanhoCampo2("Cli_For", "Complemento", dbText, 50) Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo da tabela ""Clientes / Fornecedores"" - não foi possível."
    GoTo ErrInStep
  End If
  '
  '50. Alterado o tamanho do campo Bairro de 20 para 50
  nStep = nStep + 1
  If Not gbAlteraTamanhoCampo2("Cli_For", "Bairro", dbText, 50) Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo da tabela ""Clientes / Fornecedores"" - não foi possível."
    GoTo ErrInStep
  End If
  '
  '51. Alterado o tamanho do campo Fone 2 de 15 para 43
  nStep = nStep + 1
  If Not gbAlteraTamanhoCampo2("Cli_For", "Fone 2", dbText, 43) Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo da tabela ""Clientes / Fornecedores"" - não foi possível."
    GoTo ErrInStep
  End If
  '
  '52. Alterado o tamanho do campo Inscrição de 18 para 23
  nStep = nStep + 1
  If Not gbAlteraTamanhoCampo2("Cli_For", "Inscrição", dbText, 23) Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo da tabela ""Clientes / Fornecedores"" - não foi possível."
    GoTo ErrInStep
  End If
  '
  '53. Alteração do tamanho do campo Endereço em Cli_For
  nStep = nStep + 1
  If Not gbAlteraTamanhoCampo2("Cli_For", "Endereço", dbText, 211) Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Clientes / Fornecedores"" - não foi possível."
    GoTo ErrInStep
  End If
  '
  '54. Alterado o tamanho do campo Fone 1 de 35 para 43
  nStep = nStep + 1
  If Not gbAlteraTamanhoCampoIndex("Cli_For", "Fone 1", dbText, 43, "Telefone", "Fone 1", "Código", False, False) Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo da tabela ""Clientes / Fornecedores"" - não foi possível."
    GoTo ErrInStep
  End If
  '
  '55. Alteração do tamanho do campo Cargo em Cli_For - Crédito
  nStep = nStep + 1
  If Not gbAlteraTamanhoCampo2("Cli_For - Crédito", "Cargo", dbText, 100) Then
    gsMsg = "Manutenção na Base de Dados: Alteração de campo na tabela ""Clientes /Fornecedores - Crédito"" - não foi possível."
    GoTo ErrInStep
  End If
  '----------------------------------------------------------------------------------
  
  
  '11/04/2005 - Daniel
  '
  '56. Inclusão na Tabela WEB_OrderForms
  '    do Campo Seguro
  '
  'Solicitante: Aura Prata
  '
  nStep = nStep + 1
  If Not gbGetField("WEB_OrderForms", "Seguro") Then
    If gbCreateField("WEB_OrderForms", "Seguro", dbDouble) Then
      db.Execute "UPDATE WEB_OrderForms SET Seguro = 0;"
    Else
      gsMsg = "Manutenção na Base de Dados: Adição de campo na tabela ""WEB_OrderForms"" - não foi possível."
      GoTo ErrInStep
    End If
  End If
  
  
  Exit Function

ErrInStep:
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gnStyle = vbOKOnly + vbCritical
  MsgBox gsMsg & vbCrLf & "Fase da Alteração: " & CStr(nStep), gnStyle, gsTitle

ErrHandler:
  Screen.MousePointer = vbDefault
  Call ws.Rollback
  
  '21/05/2004 - mpdea
  'Incluído interceptação de erro
  If Err.Number <> 0 Then
    gsTitle = LoadResString(201)
    gnStyle = vbOKOnly + vbCritical
    gsMsg = "Manutenção na Base de Dados (WEB) - Alterações Vitais na Base de Dados não foram possíveis."
    gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
    gsMsg = gsMsg & vbCrLf & "Fase da Alteração: " & CStr(nStep)
    MsgBox gsMsg, gnStyle, gsTitle
  End If
  
  db.Close
  ws.Close
  Set db = Nothing
  Set dbFoo = Nothing
  Set ws = Nothing
  End

End Function

Private Function gbCreateTableWEB_OrderForms() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_OrderForms")
  
  Set fd = td.CreateField("ID", dbLong)
  fd.Attributes = dbAutoIncrField
  td.Fields.Append fd
  Set fd = td.CreateField("Filial", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Sequencia", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("OrderID", dbText, 26)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Origem", dbText, 1)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Total", dbCurrency)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Passo", dbByte, 1)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("StatusShopper", dbText, 255)
  td.Fields.Append fd
  Set fd = td.CreateField("StatusAdmin", dbText, 255)
  td.Fields.Append fd
  Set fd = td.CreateField("Data", dbDate)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("CodPagamento", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Boleto", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BonusTotal", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BonusUtilizado", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("SubTotal", dbCurrency)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShippingMethod", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShippingTotal", dbCurrency)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("TraceCode", dbText, 20)
  td.Fields.Append fd
  Set fd = td.CreateField("ShopperID", dbText, 32)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShipName", dbText, 100)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShipAddress", dbText, 200)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShipCity", dbText, 50)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShipState", dbText, 40)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShipZip", dbText, 15)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShipCountry", dbText, 50)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ShipPhone", dbText, 35)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BillName", dbText, 100)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BillAddress", dbText, 200)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BillCity", dbText, 50)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BillState", dbText, 40)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BillZip", dbText, 15)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BillCountry", dbText, 50)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("BillPhone", dbText, 35)
  fd.Required = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableWEB_OrderForms = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_OrderForms = False

End Function

'21/05/2004 - mpdea
'Inclusão de novos campos
Private Function gbChangeTableWEB_OrderForms() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.TableDefs("WEB_OrderForms")
    
  Set fd = td.CreateField("Comentario", dbText, 255)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("NumParcelas", dbInteger)
  fd.Required = False
  td.Fields.Append fd
  '
  Set fd = td.CreateField("CCName", dbText, 255)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("CCType", dbText, 255)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("BancoNum", dbText, 4)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("BancoNome", dbText, 255)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("Agencia", dbText, 10)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("ContaCorrente", dbText, 20)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("CPF_CNPJ", dbText, 20)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("Titular", dbText, 100)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("ShipStreetNumber", dbText, 10)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("ShipStreetCompl", dbText, 50)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("ShipDistrict", dbText, 50)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("ShipDDDPhone", dbText, 7)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("BillStreetNumber", dbText, 10)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("BillStreetCompl", dbText, 50)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("BillDistrict", dbText, 50)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("BillDDDPhone", dbText, 7)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set td = Nothing
  
  gbChangeTableWEB_OrderForms = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbChangeTableWEB_OrderForms = False

End Function

Private Function gbCreateTableWEB_OrderItens() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_OrderItens")
  
  Set fd = td.CreateField("ID", dbLong)
  fd.Attributes = dbAutoIncrField
  td.Fields.Append fd
  Set fd = td.CreateField("OrderFormID", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("sku", dbText, 100)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Quantity", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ListPrice", dbCurrency)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Moeda", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Discount", dbCurrency)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Total", dbCurrency)
  fd.Required = True
  td.Fields.Append fd
  
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableWEB_OrderItens = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_OrderItens = False

End Function

Private Function gbCreateTableWEB_OrderStatusHistoric() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_OrderStatusHistoric")
  
  Set fd = td.CreateField("ID", dbLong)
  fd.Attributes = dbAutoIncrField
  td.Fields.Append fd
  Set fd = td.CreateField("OrderFormID", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Passo", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("StatusShopper", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("StatusAdmin", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("Data", dbDate)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("WebSynchronize", dbBoolean)
  fd.DefaultValue = True
  td.Fields.Append fd
  
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableWEB_OrderStatusHistoric = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_OrderStatusHistoric = False

End Function

Private Function gbCreateTableWEB_OrderStatus() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_OrderStatus")
  
  Set fd = td.CreateField("ID", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Name", dbText, 255)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("StatusShopper", dbText, 255)
  fd.AllowZeroLength = True
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("StatusAdmin", dbText, 255)
  fd.AllowZeroLength = True
  fd.Required = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
   
  Set td = Nothing
  
  'Insere as informações
  db.Execute "INSERT INTO WEB_OrderStatus " & _
    "(ID, Name, StatusShopper, StatusAdmin) VALUES " & _
    "(0, 'Pedido Recebido', '', '')", dbFailOnError
  db.Execute "INSERT INTO WEB_OrderStatus " & _
    "(ID, Name, StatusShopper, StatusAdmin) VALUES " & _
    "(10, 'Pagamento Confirmado', 'Pagamento confirmado, preparando envio', " & _
    "'Pagamento confirmado, preparar envio')", dbFailOnError
  db.Execute "INSERT INTO WEB_OrderStatus " & _
    "(ID, Name, StatusShopper, StatusAdmin) VALUES " & _
    "(15, 'Embalado (Recibo, Etiqueta)', 'Pagamento confirmado, preparando envio', " & _
    "'Recibo e Etiquetas Impressos')", dbFailOnError
  db.Execute "INSERT INTO WEB_OrderStatus " & _
    "(ID, Name, StatusShopper, StatusAdmin) VALUES " & _
    "(30, 'Produto Enviado', 'Pedido Enviado', 'Pedido Enviado')", dbFailOnError
  db.Execute "INSERT INTO WEB_OrderStatus " & _
    "(ID, Name, StatusShopper, StatusAdmin) VALUES " & _
    "(90, 'Pedido Cancelado', 'Pedido Cancelado', 'Pedido Cancelado')", dbFailOnError
  
  gbCreateTableWEB_OrderStatus = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_OrderStatus = False

End Function

Private Function gbCreateTableWEB_PaymentMethods() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_PaymentMethods")
  
  Set fd = td.CreateField("ID", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Name", dbText, 255)
  fd.Required = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  'Insere as informações
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(1, 'Cartões de Crédito Offline')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(2, 'Contra-Entrega')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(3, 'Depósito Bancário')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(4, 'Boleto enviado com o produto')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(5, 'Boleto Bradesco')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(6, 'Pag. Fácil Bradesco')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(7, 'Pag. Carteira Bradesco')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(8, 'Boleto Online Itaú')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(9, 'Transf. Itaú')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(10, 'Visanet MOSET')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(11, 'Visanet Setfull')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(12, 'Moset ou Setfull (problema de cookie)')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(13, 'Boleto via Paguei')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(14, 'Visa Offline')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(15, 'Master Offline')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(16, 'Diners Offline')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(17, 'Amex Offline')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(18, 'Safenet')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(19, 'Safenet Master Online')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(20, 'Safenet Diners Online')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(21, 'Itaú Shopline (aguarda retorno)')", dbFailOnError
  db.Execute "INSERT INTO WEB_PaymentMethods (ID, Name) VALUES " & _
    "(99, 'Não identificado')", dbFailOnError
  
  gbCreateTableWEB_PaymentMethods = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_PaymentMethods = False

End Function

Private Function gbCreateTableWEB_ShippingMethods() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_ShippingMethods")
  
  Set fd = td.CreateField("ID", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Name", dbText, 255)
  fd.Required = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  'Insere as informações
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(1, 'Sedex')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(2, 'Encomenda Normal')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(3, 'Motoboy')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(4, 'Entrega Própria')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(5, 'Exporte Fácil - Econômico')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(6, 'Exporte Fácil - Prioritário')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(7, 'Exporte Fácil - Expresso')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(8, 'Kwikasair')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(9, 'Virtual')", dbFailOnError
  db.Execute "INSERT INTO WEB_ShippingMethods (ID, Name) VALUES " & _
    "(10, 'Velog')", dbFailOnError
  
  gbCreateTableWEB_ShippingMethods = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_ShippingMethods = False

End Function

'24/05/2004 - mpdea
'Cria a tabela de origem do cliente
Private Function gbCreateTableWEB_ClienteOrigem() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_ClienteOrigem")
  
  Set fd = td.CreateField("ID", dbText, 50)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Origem", dbText, 255)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  'Insere as informações
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('NOT', '')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('CADE', 'Pelo Cadê')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('EMAIL', 'Recebi um Email')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('MB', 'Por outro mecanismo de procura')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('HP', 'Link em outra HomePage')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('AMIGO', 'Por conhecido')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('PANFLETO', 'Por Panfleto')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('JORNAL', 'Propaganda no Jornal')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('TV', 'Propaganda na TV')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('OUTDOOR', 'OUTDOOR')", dbFailOnError
  db.Execute "INSERT INTO WEB_ClienteOrigem (ID, Origem) VALUES " & _
    "('OUTRO', 'Outro')", dbFailOnError
  
  gbCreateTableWEB_ClienteOrigem = True
  Exit Function
  
ErrCreate:
  gbCreateTableWEB_ClienteOrigem = False
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"

End Function

'27/05/2004 - mpdea
'Incluído campo de identificação de clientes 'Web'
'
'26/05/2004 - mpdea
'Inclusão de novos campos
Private Function gbChangeTableCliFor() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.TableDefs("Cli_For")
    
  Set fd = td.CreateField("Web", dbBoolean)
  fd.Required = False
  fd.DefaultValue = False
  td.Fields.Append fd
  
  
  '22/06/2004 - mpdea
  'Comentado inclusão de campo para nova revisão
'  '
'  Set fd = td.CreateField("WebSenha", dbText, 20)
'  fd.Required = False
'  fd.AllowZeroLength = True
'  td.Fields.Append fd


  '
  Set fd = td.CreateField("Endereço Número", dbText, 10)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("DDD_Fone1", dbText, 7)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("DDD_Fone2", dbText, 7)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("RG_UF", dbText, 2)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  '
  Set fd = td.CreateField("WebEMailMerco", dbBoolean)
  fd.Required = False
  fd.DefaultValue = False
  td.Fields.Append fd
  '
  Set fd = td.CreateField("WebEMailLoja", dbBoolean)
  fd.Required = False
  fd.DefaultValue = False
  td.Fields.Append fd
  '
  Set fd = td.CreateField("WebOrigem", dbText, 50)
  fd.Required = False
  fd.AllowZeroLength = True
  td.Fields.Append fd
  
  Set td = Nothing
  
  'Atualiza clientes web
  db.Execute "UPDATE Cli_For SET Tipo = 'C', Web = TRUE WHERE Tipo = 'W'", dbFailOnError
  
  gbChangeTableCliFor = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbChangeTableCliFor = False

End Function

'Procura o próximo código livre para o registro na base de dados
'O campo deve ser do tipo numérico (Byte, Integer ou Long)
'e positivo (retorno = -1 significa nenhum intervalo livre)
Private Function glngNextFreeCode(ByVal strTableName As String, _
  ByVal strFieldName As String, ByVal lngStartValue As Long, _
  ByVal lngFinishValue As Long) As Long
  
  Dim rsGet As Recordset
  Dim lngX As Long
  Dim lngFreeCode As Long
  
  lngFreeCode = -1
  lngX = lngStartValue
  Set rsGet = db.OpenRecordset("SELECT [" & strFieldName & "] FROM [" & strTableName & "] ORDER BY [" & strFieldName & "]", dbOpenSnapshot)
  With rsGet
    If .RecordCount > 0 Then
      Do Until .EOF
        If CLng(.Fields(strFieldName).Value) > lngX Then
          lngFreeCode = lngX
          Exit Do
        End If
        lngX = lngX + 1
        If lngX >= lngFinishValue Then Exit Do 'Limite
        .MoveNext
      Loop
      '25/10/2002 - mpdea
      'Corrigido a obtenção do próximo código livre em registros
      'sequênciais sem intervalo que iniciam com lngStartValue
      If lngFreeCode = -1 And lngX > 0 Then
        lngFreeCode = lngX
      End If
    Else
      lngFreeCode = lngStartValue
    End If
    .Close
  End With
  Set rsGet = Nothing
  
  glngNextFreeCode = lngFreeCode
  
End Function

Private Function gbCreateTableWEB_Config() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_Config")
  
  Set fd = td.CreateField("ID", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("xml", dbLong)
  td.Fields.Append fd
  Set fd = td.CreateField("image", dbLongBinary)
  td.Fields.Append fd
  Set fd = td.CreateField("Filial", dbByte)
  td.Fields.Append fd
  Set fd = td.CreateField("CNX_User", dbText, 255)
  td.Fields.Append fd
  Set fd = td.CreateField("CNX_Password", dbText, 255)
  td.Fields.Append fd
  Set fd = td.CreateField("CNX_Store", dbText, 255)
  td.Fields.Append fd
  Set fd = td.CreateField("Password", dbText, 255)
  fd.AllowZeroLength = True
  td.Fields.Append fd
  Set fd = td.CreateField("CodOpReserva", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("CodOpVenda", dbInteger)
  td.Fields.Append fd
  Set fd = td.CreateField("CodOpCancelamento", dbInteger)
  td.Fields.Append fd
  
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("ID")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  'Insere as informações de configuração inicial
  db.Execute "INSERT INTO WEB_Config (ID, xml) VALUES (1, 0)", dbFailOnError
  
  gbCreateTableWEB_Config = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_Config = False

End Function

Private Function gbCreateTableWEB_ProdutosExcluir() As Boolean
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  On Error GoTo ErrCreate
  
  Set td = db.CreateTableDef("WEB_ProdutosExcluir")
  
  Set fd = td.CreateField("Codigo", dbText, 20)
  fd.Required = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Codigo")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  gbCreateTableWEB_ProdutosExcluir = True
  Exit Function
  
ErrCreate:
  MsgBox Err.Number & "-" & Err.Description, vbCritical, "Erro"
  gbCreateTableWEB_ProdutosExcluir = False

End Function

'Marca o flag de sincronismo do produto com a Loja Virtual
Public Sub WEB_SynchronizeProduct(ByVal strCode As String)
  Call db.Execute("UPDATE Produtos SET WebSynchronize = True WHERE Código = '" _
    & strCode & "' AND WebIncluded", dbFailOnError)
End Sub

'Obtém o código de operação de venda configurado para a Loja Virtual
Public Sub GetWEBCod_Op(ByRef intCodOpReserva As Integer, _
  ByRef intCodOpVenda As Integer, ByRef intCodOpCancelamento As Integer)
  
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CodOpReserva, CodOpVenda, CodOpCancelamento FROM WEB_Config WHERE ID = 1"
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not .BOF And Not .EOF Then
      intCodOpReserva = .Fields("CodOpReserva").Value
      intCodOpVenda = .Fields("CodOpVenda").Value
      intCodOpCancelamento = .Fields("CodOpCancelamento").Value
    End If
    .Close
  End With
  Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------------
'Funções Loja Virtual (Quick Web)
'
'29/04/2002 - mpdea
'<<-----------------------------------------------------------------------------------

'Obtém o ID, código e nome do comprador (cliente) através de seu ID ou código
Public Sub WEB_GetShopperData(ByRef strID As String, ByRef lngCodigo As Long, ByRef strNome As String)
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Código, WebShopperID, Nome FROM Cli_For WHERE "
  If strID <> "" Then
    strSQL = strSQL & "WebShopperID = '" & strID & "'"
  Else
    strSQL = strSQL & "Código = " & lngCodigo
  End If
  
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not .BOF And Not .EOF Then
      strID = .Fields("WebShopperID").Value & ""
      lngCodigo = .Fields("Código").Value
      strNome = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  Set rs = Nothing
  
End Sub

Public Function gstrWEB_GetDescOrigem(ByVal strID As String) As String
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Origem FROM WEB_ClienteOrigem WHERE ID = '" & strID & "'"
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not (.BOF And .EOF) Then
      gstrWEB_GetDescOrigem = .Fields("Origem").Value & ""
    End If
    .Close
  End With
  Set rs = Nothing
  
End Function

Public Function gstrWEB_GetDescPasso(ByVal bytID As Byte) As String
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Name FROM WEB_OrderStatus WHERE ID = " & bytID
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not (.BOF And .EOF) Then
      gstrWEB_GetDescPasso = .Fields("Name").Value & ""
    End If
    .Close
  End With
  Set rs = Nothing
  
End Function

Public Sub GetDataDescPasso(ByVal enuStep As enWEB_OrderFormStep, _
  ByRef strStatusShopper As String, ByRef strStatusAdmin As String)
  
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT StatusShopper, StatusAdmin FROM WEB_OrderStatus WHERE ID = " & enuStep
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not .BOF And Not .EOF Then
      strStatusShopper = .Fields("StatusShopper").Value
      strStatusAdmin = .Fields("StatusAdmin").Value
    End If
    .Close
  End With
  Set rs = Nothing
  
End Sub

'----------------------------------------------------------------------------------->>
