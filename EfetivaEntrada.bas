Attribute VB_Name = "modEfetivaEntrada"
Option Explicit

Public Function Desefetiva_Entrada(Filial As Integer, Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  'Inserido os recordsets que estavam a nível modular sem necessidade,
  'ocupando mais memória
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
'  Dim rsParametros As Recordset
  Dim rsOp_Entrada As Recordset
  Dim rsContas_Pagar As Recordset
  Dim rsResumo_Diário As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
  Dim rsEstoque_Final As Recordset
  Dim rsPreços As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
'  Dim rsGrade As Recordset
  Dim rsEntradas As Recordset
  Dim rsEntra_Prod As Recordset
  Dim rsEntraProd As Recordset
  Dim rsMovi_Parcelas As Recordset
  Dim rsLançamentos As Recordset
  Dim rsComissões As Recordset
  Dim rsFuncionarios As Recordset
  Dim rsTabelas As Recordset
  
  '08/01/2010 - Andrea
  Dim rsMovi_Cheques As Recordset
  Dim rsContas_Receber As Recordset
  '---------------------------------------------------------------------------------
  
  Dim Aux_Str As String
  Dim Mes_Atual As Integer
  Dim Ano_Atual As Integer
  Dim Erro As Integer
  Dim Ordem As Integer
  Dim Aux_Int As Integer
  Dim Saldo_Ant As Double
  Dim Saldo As Double
  Dim Caixa_Novo As Integer
  Dim Tot_Dinheiro As Double
  Dim Tot_Cheques As Double
  Dim Tot_Cheques_Pre As Double
  Dim Tot_Cartões As Double
  Dim Tot_Vales As Double
  Dim Tot_Parcelamento As Double
  Dim Cód As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Aux_Prod As String
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Edição As Long
  Dim Estoque_Final As Double
  Dim Estoque2 As Double
  Dim Custo_Médio As Double
  Dim Criar_Registro As Integer
  Dim Mensagem As String
  Dim Saldo_Conta As Double
  Dim Aux_Sequência As Long
  
  'Variável de Tratamento de Erro
  Dim nRepeatUpdateLocked As Integer
  
  Dim strSQL As String
  
  On Error GoTo Processa_Erro
  
  Screen.MousePointer = vbHourglass
  
  Set rsEntradas = db.OpenRecordset("Entradas")
  Set rsContas_Pagar = db.OpenRecordset("Contas a Pagar")
  Set rsProdutos = db.OpenRecordset("Produtos")
'  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsOp_Entrada = db.OpenRecordset("Operações Entrada", , dbReadOnly)
  Set rsResumo_Diário = db.OpenRecordset("Resumo Diário")
  Set rsEtiquetas = db.OpenRecordset("Etiquetas")
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsRes_Financeiro = db.OpenRecordset("Resumo Diário Financeiro")
'  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsPreços = db.OpenRecordset("Preços")
  Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
  Set rsEmprestimos = db.OpenRecordset("Consignação Entrada")
  Set rsCliFor = db.OpenRecordset("Cli_For")
'  Set rsGrade = db.OpenRecordset("Códigos da Grade", , dbReadOnly)
  Set rsEntra_Prod = db.OpenRecordset("Entradas - Produtos", , dbReadOnly)
  Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  Set rsComissões = db.OpenRecordset("Comissão")
  
  '08/01/2010 - Andrea
  Set rsMovi_Cheques = db.OpenRecordset("Movimento - Cheques")
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
  
  Screen.MousePointer = vbDefault
  
  rsEntradas.Index = "Sequência"
  rsEntradas.Seek "=", Filial, Mov
  If rsEntradas.NoMatch Then
    Desefetiva_Entrada = 1
    Exit Function
  End If
  
  
  ' Encontrou Movimentação
  ' Separa mes e ano
  Aux_Str = Format$(rsEntradas("Data"), "dd/mm/yyyy")
  Ano_Atual = Val(Right(Aux_Str, 4))
  Mes_Atual = Val(Mid(Aux_Str, 4, 2))
  
  ' Encontra a tabela de operações
  rsOp_Entrada.Index = "Código"
  rsOp_Entrada.Seek "=", rsEntradas("Operação")
  If rsOp_Entrada.NoMatch Then
    Desefetiva_Entrada = 2
    Exit Function
  End If
  
  ' Encontra Fornecedor
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", rsEntradas("Fornecedor")
  If rsCliFor.NoMatch Then
    Desefetiva_Entrada = 3
    Exit Function
  End If
  
  
  ' Diminui comissão se for devolução
  If rsOp_Entrada("Tipo") = "D" Then
    Aux_Sequência = 0
    Erro = False
    rsComissões.Index = "Sequência"
    Do
      rsComissões.Seek ">", Filial, Mov, 0
      If rsComissões.NoMatch Then Erro = True
      If Erro = False Then If rsComissões("Filial") <> Filial Then Erro = True
      If Erro = False Then If rsComissões("Sequência") <> Mov Then Erro = True
      If Erro = False Then
        rsComissões.Delete
      End If
    Loop Until Erro = True
  End If
  
  
  ' Atualiza arquivo de Resumo de Clientes
  ' se for Compra
  Erro = False
  rsResumo_Clientes.Index = "Sequência"
  Do
    rsResumo_Clientes.Seek ">=", Filial, Mov
    If rsResumo_Clientes.NoMatch Then Erro = True
    If Erro = False Then If rsResumo_Clientes("Filial") <> Filial Then Erro = True
    If Erro = False Then If rsResumo_Clientes("Sequência") <> Mov Then Erro = True
    If Erro = False Then
      rsResumo_Clientes.Delete
    End If
  Loop Until Erro = True
  
  
  
  ' Atualiza Caixa, se for o caso
  If rsOp_Entrada("Dinheiro") = True And (rsEntradas("Dinheiro Caixa") <> 0 Or rsEntradas("Cheque Caixa") <> 0) Then
    Erro = False
    Caixa_Novo = False
    Ordem = 0
       
    rsCaixa.Index = "Data"
    rsCaixa.Seek "<", Filial, rsEntradas("Caixa"), rsEntradas("Data"), 9999
    If rsCaixa.NoMatch Then Exit Function
    If rsCaixa("Caixa") <> rsEntradas("Caixa") Then Exit Function
    ' Neste ponto tem o último caixa no buffer
    ' Acha parcela a vista
    Ordem = rsCaixa("Ordem")
    Ordem = Ordem + 1
    Saldo_Ant = rsCaixa("Final")
    Tot_Dinheiro = rsCaixa("Total Dinheiro")
    Tot_Cheques = rsCaixa("Total Cheques")
    Tot_Cheques_Pre = rsCaixa("Total Cheques Pré")
    Tot_Cartões = rsCaixa("Total Cartões")
    Tot_Vales = rsCaixa("Total Vales")
    Tot_Parcelamento = rsCaixa("Total Parcelamento")
  
    With rsCaixa
      .AddNew
      .Fields("Filial") = Filial
      .Fields("Data") = rsEntradas("Data")
      .Fields("Hora") = Format(Time, "hh:mm:ss")
      .Fields("Caixa") = rsEntradas("Caixa")
      .Fields("Ordem") = Ordem
      .Fields("Descrição") = "Cancelamento entrada " & str(Mov)
      .Fields("Saldo Anterior") = Saldo_Ant
      
      '12/01/2010 - Andrea
      .Fields("Total Cheques Pré") = Tot_Cheques_Pre + rsEntradas("Cheque Caixa")
      '.Fields("Cheques") = rsEntradas("Cheque Caixa")
      .Fields("Cheques") = 0
      .Fields("Cheques Pré") = rsEntradas("Cheque Caixa")
      
      .Fields("Total Cartões") = Tot_Cartões
      .Fields("Total Vales") = Tot_Vales
      .Fields("Total Cheques") = Tot_Cheques
      .Fields("Total Parcelamento") = Tot_Parcelamento
      .Fields("Dinheiro") = (rsEntradas("Dinheiro Caixa") - rsEntradas("Troco"))
      .Fields("Total Dinheiro") = Tot_Dinheiro + rsEntradas("Dinheiro Caixa") - rsEntradas("Troco")
      .Fields("Final") = Tot_Dinheiro + rsEntradas("Dinheiro Caixa") + Tot_Cheques + Tot_Cheques_Pre + Tot_Cartões + Tot_Vales - rsEntradas("Troco") + rsEntradas("Cheque Caixa")
      
      .Update
    End With
    
  End If
  
  
  ' Faz Lancamento na conta bancária, se for o caso
  If rsEntradas("Valor Cheque") <> 0 Then
    ' Acha Saldo Anterior
    Saldo_Conta = 0
    rsLançamentos.Index = "Conta"
    rsLançamentos.Seek "<", rsEntradas("Conta"), rsEntradas("data"), 99999999#
    If Not rsLançamentos.NoMatch Then
      If rsLançamentos("Conta") = rsEntradas("Conta") Then
        Saldo_Conta = rsLançamentos("Saldo Atual")
      End If
    End If
    
    With rsLançamentos
      .AddNew
      .Fields("Conta") = rsEntradas("Conta")
      .Fields("Data") = rsEntradas("Data")
      .Fields("Descrição") = "Cancelamento entrada " + str(rsEntradas("Sequência"))
      .Fields("Cheque") = rsEntradas("Num Cheque")
      .Fields("Crédito") = rsEntradas("Valor Cheque")
      .Fields("Saldo Anterior") = Saldo_Conta
      .Fields("Saldo Atual") = Saldo_Conta + rsEntradas("Valor Cheque")
      .Update
    End With
  End If
  
  '----------------------------------------------------------------------------------------------------------------
  '08/01/2010 - Andrea
  'Atualiza contas a receber (cheques de clientes utilizados para pagar a compra)
  Erro = False
  Ordem = 0
  Aux_Int = 1
  
  rsMovi_Cheques.Index = "Ordem"
  Do
    rsMovi_Cheques.Seek ">", Filial, Mov, Ordem
    If rsMovi_Cheques.NoMatch Then Erro = True
    If Erro = False Then If rsMovi_Cheques("Filial") <> Filial Then Erro = True
    If Erro = False Then If rsMovi_Cheques("Sequência") <> Mov Then Erro = True

    If Erro = False Then

      strSQL = "SELECT * FROM [Contas a Receber]  WHERE Filial = " & Filial
      strSQL = strSQL & " AND Tipo='C' AND Banco = " & rsMovi_Cheques("Banco") & " AND Cheque = '" & rsMovi_Cheques("Cheque") & "'"
      '11/01/2010 - mpdea
      'Substitui virgula por ponto (formato decimal americano)
      strSQL = strSQL & " AND (Processado=True) AND Valor = " & Replace(rsMovi_Cheques("Valor"), ",", ".")
      
      Set rsContas_Receber = db.OpenRecordset(strSQL, dbOpenDynaset)
      If Not rsContas_Receber.BOF And Not rsContas_Receber.EOF Then
        
        rsContas_Receber.Edit

        rsContas_Receber("Processado") = 0
        rsContas_Receber("Data Recebimento") = 0
        rsContas_Receber("FornecedorCreditado") = 0
        rsContas_Receber("SequenciaEntrada") = 0

        rsContas_Receber.Update
      
      End If

      rsMovi_Cheques.Delete
      
    End If
    
    'Ordem = Ordem + 1
  Loop Until Erro = True
  '----------------------------------------------------------------------------------------------------------------

  
  
  ' Faz contas a pagar, se for o caso
  Erro = False
  Aux_Sequência = 0
  rsContas_Pagar.Index = "Sequência"
  Do
    rsContas_Pagar.Seek ">", Filial, Mov, Aux_Sequência
    If rsContas_Pagar.NoMatch Then Erro = True
    If Erro = False Then If rsContas_Pagar("Filial") <> Filial Then Erro = True
    If Erro = False Then If rsContas_Pagar("Sequência") <> Mov Then Erro = True
    
    If Erro = False Then
      rsContas_Pagar.Delete
    End If
  Loop Until Erro = True
  
  
  
  ' Atualiza Resumo Diário
  If rsOp_Entrada("Tipo") <> "P" Then
    rsResumo_Diário.Index = "Data"
    rsResumo_Diário.Seek "=", Filial, rsEntradas("Data")
    With rsResumo_Diário
      If .NoMatch Then
        .AddNew
        .Fields("Filial") = Filial
        .Fields("Data") = rsEntradas("Data")
      Else
        .Edit
      End If
      Select Case rsOp_Entrada("Tipo")
        Case "C"
          .Fields("Valor Compras") = CDbl(.Fields("Valor Compras")) - CDbl(rsEntradas("Total"))
        Case "T"
          .Fields("Valor T Entrada") = CDbl(.Fields("Valor T Entrada")) - CDbl(rsEntradas("Total"))
        Case "A"
          .Fields("Valor A Entrada") = CDbl(.Fields("Valor A Entrada")) - CDbl(rsEntradas("Total"))
        Case "G"
          .Fields("Valor G Entrada") = CDbl(.Fields("Valor G Entrada")) - CDbl(rsEntradas("Total"))
        Case "E"
          .Fields("Valor E Entrada") = CDbl(.Fields("Valor E Entrada")) - CDbl(rsEntradas("Total"))
        Case "D"
          .Fields("Valor Devolução") = CDbl(.Fields("Valor Devolução")) - CDbl(rsEntradas("Total"))
          
          '08/08/2003 - maikel
          '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolução) do quick store
          '.Fields("Valor Vendas") = CDbl(.Fields("Valor Vendas")) + CDbl(rsEntradas("Total"))
      End Select
      .Update
    End With
  End If
  
  
  ' Atualiza Resumo Diário Financeiro
  If rsOp_Entrada("Dinheiro") = True Then
    rsRes_Financeiro.Index = "Data"
    rsRes_Financeiro.Seek "=", Filial, rsEntradas("Data")
    With rsRes_Financeiro
      If .NoMatch Then
        .AddNew
        .Fields("Filial") = Filial
        .Fields("Data") = rsEntradas("Data")
      Else
        .Edit
      End If
      
      Select Case rsOp_Entrada("Tipo")
       Case "C"
         .Fields("Valor Compras") = CDbl(.Fields("Valor Compras")) - CDbl(rsEntradas("Total"))
       Case "T"
         .Fields("Valor T Entrada") = CDbl(.Fields("Valor T Entrada")) - CDbl(rsEntradas("Total"))
       Case "A"
         .Fields("Valor A Entrada") = CDbl(.Fields("Valor A Entrada")) - CDbl(rsEntradas("Total"))
       Case "G"
         .Fields("Valor G Entra") = CDbl(.Fields("Valor G Entra")) - CDbl(rsEntradas("Total"))
       Case "E"
         .Fields("Valor E Entra") = CDbl(.Fields("Valor E Entra")) - CDbl(rsEntradas("Total"))
       Case "D"
         .Fields("Valor Devolução") = CDbl(.Fields("Valor Devolução")) - CDbl(rsEntradas("Total"))
          
          '08/08/2003 - maikel
          '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolução) do quick store
         '.Fields("Valor Vendas") = CDbl(.Fields("Valor Vendas")) + CDbl(rsEntradas("Total"))
      End Select
      .Update
    End With
  End If
  
  
  ' Desfaz Empréstimos
  rsEmprestimos.Index = "Cliente"
  Erro = False
  Do
    rsEmprestimos.Seek ">", rsEntradas("Filial"), rsEntradas("Sequência"), 0, 0, 0, 0, 0, 0
    If rsEmprestimos.NoMatch Then Erro = True
    If Erro = False Then If rsEntradas("Filial") <> rsEmprestimos("Filial") Then Erro = True
    If Erro = False Then If rsEntradas("Sequência") <> rsEmprestimos("Sequência") Then Erro = True
    If Erro = False Then
      rsEmprestimos.Delete
    End If
  Loop Until Erro = True
  
  
  
  rsEntra_Prod.Index = "Sequência"
  Ordem = 0
Prox_Prod:
  rsEntra_Prod.Seek ">", Filial, Mov, Ordem
  
  If rsEntra_Prod.NoMatch Then GoTo Fim_Desefetiva
  If rsEntra_Prod("Filial") <> Filial Then GoTo Fim_Desefetiva
  If rsEntra_Prod("sequência") <> Mov Then GoTo Fim_Desefetiva
  
  Ordem = rsEntra_Prod("Linha")
  'Verifica se tem grade
  Cód = rsEntra_Prod("Código")
  Tamanho = 0
  Cor = 0
  
  rsProdutos.Index = "Código"
  
  Aux_Prod = rsEntra_Prod("Código")
  Acha_Produto Aux_Prod, Cód, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro
  If Aux_Erro <> 0 Then
   'Call StatusMsg("Produto não encontrado."
   GoTo Prox_Prod
  End If
  Cód = UCase(Cód)
     
  rsProdutos.Seek "=", Cód
  If rsProdutos.NoMatch Then
   GoTo Prox_Prod
  End If
  
  
  'Neste ponto CÓD tem o código do produto
  'Tamanho e Cor contém os respectivos dados
  'Agora grava arquivo do estoque
  
  '  Ajusta Estoque
  If rsOp_Entrada("Estoque") = True Then
    Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
    
'-------------------------------------------------------------------------------------
    '31/03/2004 - mpdea
    'Modificado parâmetro de abertura do recordset
    'dbOpenSnapshot (muito lento!? 8-|) para dbOpenDynaset com dbReadOnly
    'e modificado para que salve somente no final da atualização
    'de estoque o recordset
    '
    '10/10/2003 - Maikel
    '             Modificada a forma de analisar a tabela de estoque. Da forma antiga gerava erro 3022 ao efetuar movimentação com data retroativa.
    strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & Filial & _
             " AND Produto = '" & Cód & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " ORDER BY Data "
            
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        Estoque_Final = .Fields("Estoque Final")
      Else
        Estoque_Final = 0
      End If
      
      .Close
    End With
    Set rsEstoque = Nothing
    
    strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & Filial & _
             " AND Produto = '" & Cód & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " AND Data = #" & Format(rsEntradas("Data"), "mm/dd/yyyy") & "#"
             
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .Edit
      Else
        .AddNew
        .Fields("Filial").Value = Filial
        .Fields("Data").Value = rsEntradas("Data").Value
        .Fields("Produto").Value = Cód
        .Fields("Tamanho").Value = Tamanho
        .Fields("Cor").Value = Cor
        .Fields("Edição").Value = Edição
        .Fields("Classe").Value = rsProdutos("Classe").Value
        .Fields("Sub Classe").Value = rsProdutos("Sub Classe").Value
        .Fields("Estoque Anterior").Value = Estoque_Final
'        .Update
'        .Requery
      End If
    End With
'-------------------------------------------------------------------------------------

    
'    ' Encontra a posição do estoque
'    Criar_Registro = False
'    Estoque_Final = 0
'    rsEstoque.Index = "Produto"
'    rsEstoque.Seek "=", Filial, rsEntradas("Data"), Cód, Tamanho, Cor, Edição
'
'    With rsEstoque
'
'      If Not .NoMatch Then
'        Estoque_Final = .Fields("Estoque Final")
'      Else
'        .Index = "Data"
'        .Seek "<", Filial, Cód, Tamanho, Cor, Edição, rsEntradas("Data")
'        If .NoMatch Then Criar_Registro = True
'        If Not .NoMatch Then
'          If .Fields("Filial") = Filial And .Fields("Produto") = Cód And .Fields("Tamanho") = Tamanho And .Fields("Cor") = Cor And .Fields("Edição") = Edição Then
'            Criar_Registro = True
'            Estoque_Final = .Fields("Estoque Final")
'          End If
'        End If
'
'        .AddNew
'        .Fields("Filial") = Filial
'        .Fields("Data") = rsEntradas("Data")
'        .Fields("Produto") = Cód
'        .Fields("Tamanho") = Tamanho
'        .Fields("Cor") = Cor
'        .Fields("Edição") = Edição
'        .Fields("Classe") = rsProdutos("Classe")
'        .Fields("Sub Classe") = rsProdutos("Sub Classe")
'        .Fields("Estoque Anterior") = Estoque_Final
'        .Update
'
'        .Index = "Produto"
'        .Seek "=", Filial, rsEntradas("Data"), Cód, Tamanho, Cor, Edição
'      End If
      
'-------------------------------------------------------------------------------------
      
      ' neste ponto esta com o registro de estoque
      ' no buffer, agora soma com os valores da movimentação
    With rsEstoque
'      .Edit
      Select Case rsOp_Entrada("Tipo")
        Case "C"
          .Fields("Compras") = .Fields("Compras") - rsEntra_Prod("Qtde")
          .Fields("Valor Compras") = Format(.Fields("Valor Compras") - rsEntra_Prod("Preço Final"), "###########0.00")
        Case "T"
          .Fields("Transf Entra") = .Fields("Transf Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor T Entra") = Format(.Fields("Valor T Entra") - rsEntra_Prod("Preço Final"), "#############0.00")
        Case "A"
          .Fields("Ajuste Entra") = .Fields("Ajuste Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor Ajuste Entra") = Format(.Fields("Valor Ajuste Entra") - rsEntra_Prod("Preço Final"), "###############0.00")
        Case "G"
          .Fields("Grátis Entra") = .Fields("Grátis Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor Grátis Entra") = Format(.Fields("Valor Grátis Entra") - rsEntra_Prod("Preço Final"), "################0.00")
        Case "E"
          .Fields("Empre Entra") = .Fields("Empre Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor Empre Entra") = Format(.Fields("Valor Empre Entra") - rsEntra_Prod("Preço Final"), "###############0.00")
        Case "D"
          .Fields("Devolução") = .Fields("Devolução") - rsEntra_Prod("Qtde")
          .Fields("Valor Devolução") = Format(.Fields("Valor Devolução") - rsEntra_Prod("Preço Final"), "#############0.00")
          
          '08/08/2003 - maikel
          '             Comentadas as duas linhas abaixo para resolver o problema de estoque (referente a devolução) do quick store
'          .Fields("Vendas") = .Fields("Vendas") + rsEntra_Prod("Qtde")
'          .Fields("Valor Vendas") = Format(.Fields("Valor Vendas") + rsEntra_Prod("Preço Final"), "##############0.00")
      End Select
      
      Estoque2 = .Fields("Estoque Anterior")
      Estoque_Final = .Fields("Estoque Anterior") - .Fields("Vendas") + .Fields("Compras")
      Estoque_Final = Estoque_Final - .Fields("Transf Saída") + .Fields("Transf Entra")
      Estoque_Final = Estoque_Final - .Fields("Ajuste Saída") + .Fields("Ajuste Entra")
      Estoque_Final = Estoque_Final - .Fields("Grátis Saída") + .Fields("Grátis Entra")
      Estoque_Final = Estoque_Final - .Fields("Empre Saída") + .Fields("Empre Entra")
      
      '08/08/2003 - maikel
      '             Descomentada a soma da coluna devolução para resolver o problema de estoque
      Estoque_Final = Estoque_Final - .Fields("Quebras") + rsEstoque("Devolução")
      
      If rsProdutos("Estoque") = False Then
        Estoque_Final = 0
      End If
      
      .Fields("Estoque Final") = Estoque_Final
      .Update
      .Close
      Call Grava_Estoque_Final(rsEntradas("Filial"), Cód, Tamanho, Cor, Edição, CSng(Estoque_Final), rsEntradas("Data"))
    
    End With
    
  End If
  
  
     
  ' apaga etiquetas
  If rsEntra_Prod("Etiqueta") = True Then
   rsEtiquetas.Index = "Funcionário"
   rsEtiquetas.Seek "=", rsEntradas("Digitador"), Cód, Tamanho, Cor
   If Not rsEtiquetas.NoMatch Then
      rsEtiquetas.Edit
      rsEtiquetas("Qtde") = rsEtiquetas("Qtde") - rsEntra_Prod("Qtde")
   End If
  End If
  
  GoTo Prox_Prod
  
  
Fim_Desefetiva:
  Desefetiva_Entrada = 0
  Call StatusMsg("")
  
  '---------------------------------------------------------------------------------
  '31/03/2004 - mpdea
  '
  'Incluído o fechamento dos recordsets abertos e suas desassociações
  '---------------------------------------------------------------------------------
  rsEntradas.Close
  rsEntra_Prod.Close
  rsOp_Entrada.Close
  rsCliFor.Close
  rsContas_Pagar.Close
  rsProdutos.Close
  rsResumo_Diário.Close
  rsEtiquetas.Close
  rsCaixa.Close
  rsRes_Financeiro.Close
  rsPreços.Close
  rsResumo_Clientes.Close
  rsEmprestimos.Close
  rsMovi_Parcelas.Close
  rsLançamentos.Close
  rsComissões.Close
  
  Set rsEntradas = Nothing
  Set rsEntra_Prod = Nothing
  Set rsOp_Entrada = Nothing
  Set rsCliFor = Nothing
  Set rsContas_Pagar = Nothing
  Set rsProdutos = Nothing
  Set rsResumo_Diário = Nothing
  Set rsEtiquetas = Nothing
  Set rsCaixa = Nothing
  Set rsRes_Financeiro = Nothing
  Set rsEstoque = Nothing
  Set rsPreços = Nothing
  Set rsResumo_Clientes = Nothing
  Set rsEmprestimos = Nothing
  Set rsMovi_Parcelas = Nothing
  Set rsLançamentos = Nothing
  Set rsComissões = Nothing
  '---------------------------------------------------------------------------------
  
  Exit Function
  
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case Err.Number
    Case 3186, 3197, 3218, 3260 'Registro bloqueado
      If nRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        nRepeatUpdateLocked = nRepeatUpdateLocked + 1
        Call WaitSeconds(1) 'Aguarda um segundo
        Resume
      Else
        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
          "uma nova tentativa.", vbExclamation + vbOKCancel, "Desefetiva Entrada") = vbOK Then
          nRepeatUpdateLocked = 0
          Resume
        Else
          Desefetiva_Entrada = -1 'Ação cancelada
          Exit Function
        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Desefetiva Entrada")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Desefetiva_Entrada = -1 'Ação cancelada
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select
  
End Function

Public Function Efetiva_Entrada(ByVal Filial As Integer, ByVal Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  'Inserido os recordsets que estavam a nível modular sem necessidade,
  'ocupando mais memória
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
  Dim rsParametros As Recordset
  Dim rsOp_Entrada As Recordset
  Dim rsContas_Pagar As Recordset
  Dim rsResumo_Diário As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
  Dim rsEstoque_Final As Recordset
  Dim rsPreços As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
  Dim rsGrade As Recordset
  Dim rsEntradas As Recordset
  Dim rsEntra_Prod As Recordset
  Dim rsEntraProd As Recordset
  Dim rsMovi_Parcelas As Recordset
  Dim rsLançamentos As Recordset
  Dim rsComissões As Recordset
  Dim rsFuncionarios As Recordset
  Dim rsTabelas As Recordset
  '07/01/2009 - Andrea
  Dim rsMovi_Cheques As Recordset
  Dim rsContas_Receber As Recordset
  '---------------------------------------------------------------------------------
  
  Dim Erro As Integer
  Dim Ordem As Integer
  Dim Aux_Int As Integer
  Dim Saldo_Ant As Double
  Dim Saldo As Double
  Dim Caixa_Novo As Integer
  Dim Tot_Dinheiro As Double
  Dim Tot_Cheques As Double
  Dim Tot_Cheques_Pre As Double
  Dim Tot_Cartões As Double
  Dim Tot_Vales As Double
  Dim Tot_Parcela As Double
  Dim Cód As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edição As Long
  Dim Estoque_Final As Double
  Dim Estoque2 As Double
  Dim Custo_Médio As Double
  Dim Criar_Registro As Integer
  Dim Saldo_Conta As Double
  Dim Comissão As Double
  Dim Tipo As Integer
  Dim Emp_Existe As Boolean
  Dim Saldo_Emp As Single
  Dim Ordem_Emp As Integer

  Dim nQtde As Single
  Dim nPrecoFinal As Currency
  Dim nPreco As Currency
  Dim bEtiqueta As Boolean
  Dim sCodProd As String
  Dim nRow As Long
  
  Dim nCaixa As Integer
  Dim dtData As Date
  Dim dtDataEmissao As Date
  Dim dtDataAcerto As Date
  Dim nCodDigitador As Integer
  Dim nCodFornecedor As Long
  Dim nCxDinheiro As Currency
  Dim nCxCheque As Currency
  Dim nCxChequePre As Currency
  Dim nConta As Integer
  Dim dtBomPara As Date
  Dim sDescricao As String
  Dim sNumCheque As String
  Dim nValCheque As Currency
  Dim sNF As String
  Dim nTotal As Currency
  Dim sTabela As String
   
  Dim strSQL As String
   
  'Variável de Tratamento de Erro
  Dim nRepeatUpdateLocked As Integer
  
  'Variáveis WEB
  Dim blnWEB_Sale As Boolean
  Dim blnWEBSynchronize As Boolean
  
  
  '-------------------------------------------
  '16/05/2006 - mpdea
  'Preço para cálculo de Custo
  Dim dblPrecoFinalCusto As Double
  '-------------------------------------------
  
  
  '14/06/2006 - mpdea
  'Frete
  Dim sngPercFrete As Single
  Dim dblValorFrete As Double
  Dim dblValorTotal As Double
  Dim dblValorIPI As Double
  
  '30/08/2007 - Anderson
  'Implementação do campo para automatização do preço de custo.
  'Solicitante: Candy Clean
  Dim dblPrecoCustoCalculado As Double
  
  
  On Error GoTo Processa_Erro
  
  
  '11/07/2002 - mpdea
  'Verifica a existência da movimentação e seta o registro de entrada principal
  strSQL = "SELECT * FROM Entradas WHERE Filial = " & Filial & _
           " AND Sequência = " & Mov
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rsEntradas
    If .BOF And .EOF Then
      'Movimentação não existe
      Efetiva_Entrada = -99
      .Close
      Set rsEntradas = Nothing
      Exit Function
    End If
    
    dtData = .Fields("Data").Value
    If IsNull(.Fields("Data Emissão").Value) Then
      dtDataEmissao = dtData
    Else
      dtDataEmissao = .Fields("Data Emissão").Value
    End If
    If IsNull(.Fields("Data Acerto Empréstimo").Value) Then
      dtDataAcerto = vbNull
    Else
      dtDataAcerto = .Fields("Data Acerto Empréstimo").Value
    End If
    nCodDigitador = .Fields("Digitador").Value
    nCodFornecedor = .Fields("Fornecedor").Value
    
    nCxDinheiro = .Fields("Dinheiro Caixa").Value
    
    '08/01/2010 - Andrea
    nCxDinheiro = nCxDinheiro + (.Fields("Troco").Value * -1)
    
    '12/01/2010 - Andrea
    'O valor pago em cheques (.Fields("Cheque Caixa").Value), será
    'de cheques pré-datados, pq os cheques a vista nao aparecem no grid
    'de cheques por já estarem processados.
    nCxChequePre = .Fields("Cheque Caixa").Value
    'nCxCheque = .Fields("Cheque Caixa").Value
    nCxCheque = 0
        
    nConta = .Fields("Conta").Value
    If IsNull(.Fields("Bom Para").Value) Then
      dtBomPara = vbNull
    Else
      dtBomPara = .Fields("Bom Para").Value
    End If
    sDescricao = .Fields("Descrição").Value & ""
    sNumCheque = .Fields("Num Cheque").Value & ""
    nValCheque = .Fields("Valor Cheque").Value
    sNF = .Fields("Nota Fiscal").Value & ""
    nTotal = .Fields("Total").Value
  End With
  
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  'Implementado verificação de venda do tipo WEB - blnWEB_Sale
  'blnWEB_Sale inibe erro ao verificar cliente e digitador
  '---------------------------------------------------------------------------------
  blnWEB_Sale = CLng("0" & rsEntradas.Fields("WebOrderFormID").Value) > 0
  
  
  Rem Encontra a tabela de operações
  Set rsOp_Entrada = db.OpenRecordset("Operações Entrada", , dbReadOnly)
  rsOp_Entrada.Index = "Código"
  rsOp_Entrada.Seek "=", rsEntradas.Fields("Operação").Value
  If rsOp_Entrada.NoMatch Then
    Efetiva_Entrada = 1
    rsOp_Entrada.Close
    Set rsOp_Entrada = Nothing
    Exit Function
  End If
  
  Rem Acha Funcionários
  Set rsFuncionarios = db.OpenRecordset("Funcionários")
  rsFuncionarios.Index = ("Código")
  rsFuncionarios.Seek "=", nCodDigitador
  If rsFuncionarios.NoMatch And Not blnWEB_Sale Then  '-> blnWEB_Sale inibe erro
    Efetiva_Entrada = 2
    rsFuncionarios.Close
    Set rsFuncionarios = Nothing
    Exit Function
  End If
  
  Rem Encontra Fornecedor
  Set rsCliFor = db.OpenRecordset("Cli_For")
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", nCodFornecedor
  If rsCliFor.NoMatch And Not blnWEB_Sale Then  '-> blnWEB_Sale inibe erro
    Efetiva_Entrada = 3
    rsCliFor.Close
    Set rsCliFor = Nothing
    Exit Function
  End If
  
  If rsOp_Entrada("Tipo") = "C" Then
    rsCliFor.Edit
    rsCliFor("Última Compra") = dtData
    rsCliFor("Data Alteração") = Format(Date, "dd/mm/yyyy")
    rsCliFor.Update
  End If
  
  Set rsContas_Pagar = db.OpenRecordset("Contas a Pagar")
  Set rsProdutos = db.OpenRecordset("Produtos")
  Set rsResumo_Diário = db.OpenRecordset("Resumo Diário")
  Set rsEtiquetas = db.OpenRecordset("Etiquetas")
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsRes_Financeiro = db.OpenRecordset("Resumo Diário Financeiro")
'  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsPreços = db.OpenRecordset("Preços")
  Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
  Set rsEmprestimos = db.OpenRecordset("Consignação Entrada")
  Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsLançamentos = db.OpenRecordset("Lançamentos Bancários")
  Set rsComissões = db.OpenRecordset("Comissão")
'  Set rsTabelas = db.OpenRecordset("Tabela de Preços")
  
  '07/01/2010 - Andrea
  Set rsMovi_Cheques = db.OpenRecordset("Movimento - Cheques")
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber")

  Rem Atualiza Caixa, se for o caso
  If nCxDinheiro <> 0 Or nCxCheque <> 0 Or nCxChequePre <> 0 Then
   
    nCaixa = rsEntradas.Fields("Caixa").Value
    
    Erro = False
    Caixa_Novo = False
    Ordem = 0
    
    rsCaixa.Index = "Data"
    rsCaixa.Seek "<", Filial, nCaixa, dtData, 9999
    If rsCaixa.NoMatch Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Filial") <> Filial Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Data") <> dtData Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Caixa") <> nCaixa Then Caixa_Novo = True
    
    If Caixa_Novo = True Then 'Começa o Caixa do dia
      Erro = False
      rsCaixa.Seek "<", Filial, nCaixa, dtData, 0
      If rsCaixa.NoMatch Then Erro = True
      If Not Erro Then If rsCaixa("Filial") <> Filial Then Erro = True
      If Not Erro Then If rsCaixa("Caixa") <> nCaixa Then Erro = True
      If Erro = True Then  'Não existe dia anterior
        rsCaixa.AddNew
        rsCaixa("Filial") = Filial
        rsCaixa("Data") = dtData
        rsCaixa("Hora") = Format(Time, "hh:mm:ss")
        rsCaixa("Funcionário") = nCodDigitador
        rsCaixa("Caixa") = nCaixa
        
        Ordem = 1
        rsCaixa("Ordem") = Ordem
        rsCaixa("Saldo Anterior") = 0
        rsCaixa("Final") = 0
        rsCaixa("Descrição") = "Início do dia"
        rsCaixa.Update
      Else
        Ordem = 1
        Saldo_Ant = rsCaixa("Final")
        Tot_Dinheiro = rsCaixa("Total Dinheiro")
        Tot_Cheques = rsCaixa("Total Cheques")
        Tot_Cheques_Pre = rsCaixa("Total Cheques Pré")
        Tot_Cartões = rsCaixa("Total Cartões")
        Tot_Vales = rsCaixa("Total Vales")
        Tot_Parcela = rsCaixa("Total Parcelamento")
                          
        rsCaixa.AddNew
        rsCaixa("Filial") = Filial
        rsCaixa("Data") = dtData
        rsCaixa("Hora") = Format(Time, "hh:mm:ss")
        rsCaixa("Funcionário") = nCodDigitador
        rsCaixa("Caixa") = nCaixa
        rsCaixa("Ordem") = Ordem
        rsCaixa("Descrição") = "Início do dia"
        rsCaixa("Saldo Anterior") = Saldo_Ant
        rsCaixa("Dinheiro") = Tot_Dinheiro
        rsCaixa("Cheques") = Tot_Cheques
        rsCaixa("Cheques Pré") = Tot_Cheques_Pre
        rsCaixa("Cartões") = Tot_Cartões
        rsCaixa("Vales") = Tot_Vales
        rsCaixa("Total Dinheiro") = Tot_Dinheiro
        rsCaixa("Total Cheques") = Tot_Cheques
        rsCaixa("Total Cheques Pré") = Tot_Cheques_Pre
        rsCaixa("Total Cartões") = Tot_Cartões
        rsCaixa("Total Vales") = Tot_Vales
        rsCaixa("Parcelamento") = Tot_Parcela
        rsCaixa("Total Parcelamento") = Tot_Parcela
        rsCaixa("Final") = Saldo_Ant
        rsCaixa.Update
      End If
       
      rsCaixa.Seek "<", Filial, nCaixa, dtData, 9999
    End If
  
     
     Rem Neste ponto tem o último caixa no buffer
     Rem Acha parcela a vista
     Ordem = rsCaixa("Ordem")
     Ordem = Ordem + 1
     Saldo_Ant = rsCaixa("Final")
     Tot_Dinheiro = rsCaixa("Total Dinheiro")
     Tot_Cheques = rsCaixa("Total Cheques")
     Tot_Cheques_Pre = rsCaixa("Total Cheques Pré")
     Tot_Cartões = rsCaixa("Total Cartões")
     Tot_Vales = rsCaixa("Total Vales")
     Tot_Parcela = rsCaixa("Total Parcelamento")
  
      rsCaixa.AddNew
        rsCaixa("Filial") = Filial
        rsCaixa("Data") = dtData
        rsCaixa("Hora") = Format(Time, "hh:mm:ss")
        rsCaixa("Funcionário") = nCodDigitador
        rsCaixa("Caixa") = nCaixa
        rsCaixa("Ordem") = Ordem
        rsCaixa("Descrição") = "Entrada número " & str(Mov)
        rsCaixa("Saldo Anterior") = Saldo_Ant
        rsCaixa("Total Cheques Pré") = Tot_Cheques_Pre - nCxChequePre
        rsCaixa("Total Cartões") = Tot_Cartões
        rsCaixa("Total Vales") = Tot_Vales
        rsCaixa("Cheques") = -(nCxCheque)
        rsCaixa("Cheques Pré") = -(nCxChequePre)
        rsCaixa("Total Cheques") = Tot_Cheques - nCxCheque '- nCxChequePre
        rsCaixa("Dinheiro") = -(nCxDinheiro)
        rsCaixa("Total Dinheiro") = Tot_Dinheiro - nCxDinheiro
        rsCaixa("Total Parcelamento") = Tot_Parcela
        rsCaixa("Final") = Tot_Dinheiro - nCxCheque - nCxChequePre - nCxDinheiro + Tot_Cheques + Tot_Cheques_Pre + Tot_Cartões + Tot_Vales
      rsCaixa.Update
  End If



  Rem Faz Lancamento na conta bancária, se for o caso
  If nValCheque <> 0 Then
    Rem Acha Saldo Anterior
    Saldo_Conta = 0
    rsLançamentos.Index = "Conta"
    rsLançamentos.Seek "<", nConta, dtBomPara, 99999999#
    If Not rsLançamentos.NoMatch Then
      If rsLançamentos("Conta") = nConta Then
        Saldo_Conta = rsLançamentos("Saldo Atual")
      End If
    End If
    
    rsLançamentos.AddNew
    rsLançamentos("Conta") = nConta
    rsLançamentos("Data") = dtBomPara
    rsLançamentos("Descrição") = sDescricao
    rsLançamentos("Cheque") = sNumCheque
    rsLançamentos("Débito") = nValCheque
    rsLançamentos("Saldo Anterior") = Saldo_Conta
    rsLançamentos("Saldo Atual") = Saldo_Conta - nValCheque
    rsLançamentos.Update
  End If
  
  
 
  '----------------------------------------------------------------------------------------------------------------
  '07/01/2010 - Andrea
  'Atualiza contas a receber (cheques de clientes utilizados para pagar a compra)
  Erro = False
  Ordem = 0
  Aux_Int = 1
  
 
  rsMovi_Cheques.Index = "Ordem"
  Do
    rsMovi_Cheques.Seek ">", Filial, Mov, Ordem
    If rsMovi_Cheques.NoMatch Then Erro = True
    If Erro = False Then If rsMovi_Cheques("Filial") <> Filial Then Erro = True
    If Erro = False Then If rsMovi_Cheques("Sequência") <> Mov Then Erro = True

    If Erro = False Then

      strSQL = "SELECT * FROM [Contas a Receber]  WHERE Filial = " & Filial
      strSQL = strSQL & " AND Tipo='C' AND Banco = " & rsMovi_Cheques("Banco") & " AND Cheque = '" & rsMovi_Cheques("Cheque") & "'"
      '11/01/2010 - mpdea
      'Substitui virgula por ponto (formato decimal americano)
      strSQL = strSQL & " AND (Processado = 0) AND Valor = " & Replace(rsMovi_Cheques("Valor"), ",", ".")
      
      Set rsContas_Receber = db.OpenRecordset(strSQL, dbOpenDynaset)
      If Not rsContas_Receber.BOF And Not rsContas_Receber.EOF Then
        
        rsContas_Receber.Edit

        rsContas_Receber("Processado") = 1
        rsContas_Receber("Data Recebimento") = Format(Date, "dd/mm/yyyy")
        rsContas_Receber("FornecedorCreditado") = nCodFornecedor
        rsContas_Receber("SequenciaEntrada") = Mov

        rsContas_Receber.Update
      
      End If

    End If
    Ordem = Ordem + 1
  Loop Until Erro = True

  '----------------------------------------------------------------------------------------------------------------
  Rem Faz contas a pagar, se for o caso
  Erro = False
  Ordem = 0
  Aux_Int = 1
  rsMovi_Parcelas.Index = "Ordem"
  
  Do
    rsMovi_Parcelas.Seek ">", Filial, Mov, Ordem
    If rsMovi_Parcelas.NoMatch Then Erro = True
    If Erro = False Then If rsMovi_Parcelas("Filial") <> Filial Then Erro = True
    If Erro = False Then If rsMovi_Parcelas("Sequência") <> Mov Then Erro = True
    
    If Erro = False Then
      Ordem = rsMovi_Parcelas("Ordem")
      If rsMovi_Parcelas("Bom") >= dtData Then
        rsContas_Pagar.AddNew
        rsContas_Pagar("Filial") = Filial
        rsContas_Pagar("Fornecedor") = nCodFornecedor
        rsContas_Pagar("Data Emissão") = dtDataEmissao
        rsContas_Pagar("Descrição") = "Parcela " & str(Aux_Int)
        rsContas_Pagar("Vencimento") = rsMovi_Parcelas("Bom")
        rsContas_Pagar("Valor") = rsMovi_Parcelas("Valor")
        rsContas_Pagar("Sequência") = Mov
        rsContas_Pagar("Nota") = sNF
        '30/09/2004 - Daniel
        'Tratamento para Consignações da Resultado
        If frmImpressaoNFPrestacao.gbConsignacaoResultado Then
          rsContas_Pagar("Centro de Custo") = IIf(IsNumeric(Trim(frmImpressaoNFPrestacao.cboCodigoCC.Text)), Trim(frmImpressaoNFPrestacao.cboCodigoCC.Text), 1)
        Else
          rsContas_Pagar("Centro de Custo") = IIf(IsNumeric(Trim(frmEntrada.cboCodigoCC.Text)), Trim(frmEntrada.cboCodigoCC.Text), 1)
        End If
        rsContas_Pagar("Data Alteração") = Format(Date, "dd/mm/yyyy")
        rsContas_Pagar.Update
        Aux_Int = Aux_Int + 1
      End If
    End If
  Loop Until Erro = True
 
  Rem Atualiza Resumo Diário
  If rsOp_Entrada("Tipo") <> "P" Then
    rsResumo_Diário.Index = "Data"
    rsResumo_Diário.Seek "=", Filial, dtData
    If rsResumo_Diário.NoMatch Then
      rsResumo_Diário.AddNew
      rsResumo_Diário("Filial") = Filial
      rsResumo_Diário("Data") = dtData
    Else
      rsResumo_Diário.Edit
    End If
    If rsOp_Entrada("Tipo") = "C" Then rsResumo_Diário("Valor Compras") = CDbl(rsResumo_Diário("Valor Compras")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "T" Then rsResumo_Diário("Valor T Entrada") = CDbl(rsResumo_Diário("Valor T Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "A" Then rsResumo_Diário("Valor A Entrada") = CDbl(rsResumo_Diário("Valor A Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "G" Then rsResumo_Diário("Valor G Entrada") = CDbl(rsResumo_Diário("Valor G Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "E" Then rsResumo_Diário("Valor E Entrada") = CDbl(rsResumo_Diário("Valor E Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "D" Then
      rsResumo_Diário("Valor Devolução") = CDbl(rsResumo_Diário("Valor Devolução")) + CDbl(nTotal)
      '08/08/2003 - maikel
      '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolução) do quick store
      'rsResumo_Diário("Valor Vendas") = CDbl(rsResumo_Diário("Valor Vendas")) - CDbl(nTotal)
    End If
    rsResumo_Diário.Update
  End If
  
 
  Rem Atualiza Resumo Diário Financeiro
  If rsOp_Entrada("Dinheiro") = True Then
    rsRes_Financeiro.Index = "Data"
    rsRes_Financeiro.Seek "=", Filial, dtData
    If rsRes_Financeiro.NoMatch Then
      rsRes_Financeiro.AddNew
      rsRes_Financeiro("Filial") = Filial
      rsRes_Financeiro("Data") = dtData
    Else
      rsRes_Financeiro.Edit
    End If
    
    If rsOp_Entrada("Tipo") = "C" Then rsRes_Financeiro("Valor Compras") = CDbl(rsRes_Financeiro("Valor Compras")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "T" Then rsRes_Financeiro("Valor T Entrada") = CDbl(rsRes_Financeiro("Valor T Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "A" Then rsRes_Financeiro("Valor A Entrada") = CDbl(rsRes_Financeiro("Valor A Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "G" Then rsRes_Financeiro("Valor G Entra") = CDbl(rsRes_Financeiro("Valor G Entra")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "E" Then rsRes_Financeiro("Valor E Entra") = CDbl(rsRes_Financeiro("Valor E Entra")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "D" Then
      rsRes_Financeiro("Valor Devolução") = CDbl(rsRes_Financeiro("Valor Devolução")) + CDbl(nTotal)
      '08/08/2003 - maikel
      '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolução) do quick store
      'rsRes_Financeiro("Valor Vendas") = CDbl(rsRes_Financeiro("Valor Vendas")) - CDbl(nTotal)
    End If
    rsRes_Financeiro.Update
  End If
  

  '--------------------------------------------------------------------------------
  '14/06/2006 - mpdea
  'Percentual de Frete para cálculos de custo
  If rsOp_Entrada.Fields("SomarFreteCustoProduto").Value Then
    'Valor do frete
    Call IsDataType(dtDouble, rsEntradas.Fields("Frete").Value, dblValorFrete)
    'Valor total
    Call IsDataType(dtDouble, rsEntradas.Fields("Total").Value, dblValorTotal)
    'Verifica se o frete soma no total
    If rsOp_Entrada.Fields("Somar Frete ao Total").Value Then
      dblValorTotal = dblValorTotal - dblValorFrete
    End If
    'Verifica se calcula IPI somente para o Total
    If rsOp_Entrada.Fields("IPI TOT").Value Then
      Call IsDataType(dtDouble, rsEntradas.Fields("IPI").Value, dblValorIPI)
      dblValorTotal = dblValorTotal - dblValorIPI
    End If
    'Percentual de frete
    If dblValorTotal > 0 Then
      sngPercFrete = dblValorFrete / dblValorTotal
    End If
  End If
  '--------------------------------------------------------------------------------


  rsProdutos.Index = "Código"
  
  strSQL = "SELECT * FROM [Entradas - Produtos] WHERE Filial = " & Filial & _
           " AND Sequência = " & Mov
  Set rsEntraProd = db.OpenRecordset(strSQL, dbOpenSnapshot)
  If rsEntraProd.BOF And rsEntraProd.EOF Then
    'Movimentação de produtos não existe
    Efetiva_Entrada = -98
    rsEntraProd.Close
    Set rsEntraProd = Nothing
    Exit Function
  End If
  
  Do Until rsEntraProd.EOF
    
    blnWEBSynchronize = False
    
    sCodProd = rsEntraProd.Fields("Código").Value
    If Len(sCodProd) > 0 Then
      
      With rsEntraProd
        nQtde = .Fields("Qtde").Value
        nPrecoFinal = .Fields("Preço Final").Value
        nPreco = .Fields("Preço").Value
        bEtiqueta = .Fields("Etiqueta").Value
      End With
      
      Ordem = nRow + 1
      'Verifica se tem grade
      Cód = ""
      Tamanho = 0
      Cor = 0
      Edição = 0
      Tipo = 0
      Erro = 0
      
      Call Acha_Produto(sCodProd, Cód, Tamanho, Cor, Edição, Tipo, Erro)
      
      If Erro = 0 Then
      
        Cód = UCase(Cód)
        
        rsProdutos.Seek "=", Cód
         
        'Neste ponto CÓD tem o código do produto
        'Tamanho e Cor contém os respectivos dados
        'Agora grava arquivo do estoque
        
        Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
      
        Rem  Ajusta Estoque
        If rsOp_Entrada("Estoque") = True And Not rsEntraProd("InGeradoViaConsig") Then
        
'-------------------------------------------------------------------------------------
    '04/03/2004 - mpdea
    'Modificado parâmetro de abertura do recordset
    'dbOpenSnapshot (muito lento!? 8-|) para dbOpenDynaset com dbReadOnly
    'e modificado para que salve somente no final da atualização
    'de estoque o recordset
    '
    '10/10/2003 - Maikel
    '             Modificada a forma de analisar a tabela de estoque. Da forma antiga gerava erro 3022 ao efetuar movimentação com data retroativa.
'''    strSQL = "SELECT * FROM Estoque WHERE "

    strSQL = "SELECT * FROM [Estoque Final] WHERE " & _
             " Filial = " & Filial & _
             " AND Produto = '" & Cód & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição    ''' & " ORDER BY Data"
             
'    'LOG ESPECIFICO PARA MARE MANSA
'    Dim sSQL_Aux As String
'    Dim sCod_mare As String
'    Dim iTam_mare As String
'    Dim iCor_mare As String
'
'    If IsNull(Cód) Then
'      sCod_mare = "N"
'    Else
'      sCod_mare = Cód
'    End If
'    If IsNull(Tamanho) Then
'      iTam_mare = "N"
'    Else
'      iTam_mare = Tamanho
'    End If
'    If IsNull(Cor) Then
'      iCor_mare = "N"
'    Else
'      iCor_mare = Cor
'    End If
'
'    sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'    sSQL_Aux = sSQL_Aux & Left("1ENTRADA - " & Filial & " : " & sCod_mare & " : " & iTam_mare & " : " & iCor_mare & " : " & Edição, 80) & "', 'VENDENDO MARE')"
'    db.Execute sSQL_Aux, dbFailOnError
'    'fim
            
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
'        .MoveFirst
'''        .MoveLast
'''        Estoque_Final = .Fields("Estoque Final")
        Estoque_Final = .Fields("Estoque Atual")
      Else
        Estoque_Final = 0
      End If
      
'      'LOG ESPECIFICO PARA MARE MANSA
'      Dim sData_mare As String
'      If Not (.BOF And .EOF) Then
'        sData_mare = rsEstoque.Fields("Data")
'      Else
'        sData_mare = "01/01/2030"
'      End If
'
'      sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'      sSQL_Aux = sSQL_Aux & Left("2ENTRADA - " & sCod_mare & " : " & Estoque_Final & " : " & Format(sData_mare, "mm/dd/yyyy"), 80) & "', 'ENTRADA MARE')"
'      db.Execute sSQL_Aux, dbFailOnError
'      'fim
      
      .Close
    End With
    Set rsEstoque = Nothing
    
    strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & Filial & _
             " AND Produto = '" & Cód & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edição = " & Edição & _
             " AND Data = #" & Format(Now, "mm/dd/yyyy") & "#"
             '''" AND Data = #" & Format(rsEntradas("Data"), "mm/dd/yyyy") & "#"
            
            
            
'    'LOG ESPECIFICO PARA MARE MANSA
'    If IsNull(Cód) Then
'      sCod_mare = "N"
'    Else
'      sCod_mare = Cód
'    End If
'    If IsNull(Tamanho) Then
'      iTam_mare = "N"
'    Else
'      iTam_mare = Tamanho
'    End If
'    If IsNull(Cor) Then
'      iCor_mare = "N"
'    Else
'      iCor_mare = Cor
'    End If
'    sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'    sSQL_Aux = sSQL_Aux & Left("3ENTRADA - " & Filial & " : " & sCod_mare & " : " & iTam_mare & " : " & iCor_mare & " : " & Edição & " : " & Format(rsEntradas("Data"), "mm/dd/yyyy"), 80) & "', 'VENDENDO MARE')"
'    db.Execute sSQL_Aux, dbFailOnError
'    'fim
    
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .Edit
        
'        'LOG ESPECIFICO PARA MARE MANSA
'        If IsNull(Cód) Then
'            sCod_mare = "N"
'        Else
'            sCod_mare = Cód
'        End If
'
'        sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'        sSQL_Aux = sSQL_Aux & Left("4BENTRADA - " & sCod_mare & " : UPDATE NA LINHA NA TABELA DE ESTOQUE", 80) & "', 'VENDENDO MARE')"
'        db.Execute sSQL_Aux, dbFailOnError
'        'fim
      Else
        .AddNew
        .Fields("Filial").Value = Filial
        .Fields("Data").Value = Format(Now, "dd/mm/yyyy")
        '''.Fields("Data").Value = rsEntradas("Data").Value
        .Fields("Produto").Value = Cód
        .Fields("Tamanho").Value = Tamanho
        .Fields("Cor").Value = Cor
        .Fields("Edição").Value = Edição
        .Fields("Classe").Value = rsProdutos("Classe").Value
        .Fields("Sub Classe").Value = rsProdutos("Sub Classe").Value
        .Fields("Estoque Anterior").Value = Estoque_Final
'        .Update
'        .Requery

'        'LOG ESPECIFICO PARA MARE MANSA
'        If Not (.BOF And .EOF) Then
'          sData_mare = rsEstoque.Fields("Data")
'        Else
'          sData_mare = "01/01/2030"
'        End If
'
'        sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'        sSQL_Aux = sSQL_Aux & Left("4AENTRADA - " & sCod_mare & " : " & sData_mare & " : " & Estoque_Final & " : " & "NOVA LINHA NA TABELA DE ESTOQUE", 80) & "', 'VENDENDO MARE')"
'        db.Execute sSQL_Aux, dbFailOnError
'        'fim

      End If
    End With
'-------------------------------------------------------------------------------------

'          Rem Encontra a posição do estoque
'          Criar_Registro = False
'          Estoque_Final = 0
'          rsEstoque.Index = "Produto"
'          rsEstoque.Seek "=", Filial, dtData, Cód, Tamanho, Cor, Edição
'
'          If Not rsEstoque.NoMatch Then
'            Estoque_Final = rsEstoque("Estoque Final")
'            Estoque2 = rsEstoque("Estoque Final")
'          End If
'          If rsEstoque.NoMatch Then
'            rsEstoque.Index = "Data"
'            rsEstoque.Seek "<", Filial, Cód, Tamanho, Cor, Edição, dtData
'            If rsEstoque.NoMatch Then Criar_Registro = True
'            If Not rsEstoque.NoMatch Then
'              If rsEstoque("Filial") = Filial And rsEstoque("Produto") = Cód And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edição") = Edição Then
'                Criar_Registro = True
'                Estoque_Final = rsEstoque("Estoque Final")
'                Estoque2 = Estoque_Final
'              End If
'            End If
'
'            rsEstoque.AddNew
'            rsEstoque("Filial") = Filial
'            rsEstoque("Data") = dtData
'            rsEstoque("Produto") = Cód
'            rsEstoque("Tamanho") = Tamanho
'            rsEstoque("Cor") = Cor
'            rsEstoque("Edição") = Edição
'            rsEstoque("Classe") = rsProdutos("Classe")
'            rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
'            rsEstoque("Estoque Anterior") = Estoque_Final
'            rsEstoque.Update
'
'            rsEstoque.Index = "Produto"
'            rsEstoque.Seek "=", Filial, dtData, Cód, Tamanho, Cor, Edição
'          End If
      
'-------------------------------------------------------------------------------------
      
          Rem neste ponto esta com o registro de estoque
          Rem no buffer, agora soma com os valores da movimentação
'          rsEstoque.Edit
          Select Case rsOp_Entrada("Tipo")
            Case "C"
              rsEstoque("Compras") = rsEstoque("Compras") + nQtde
              rsEstoque("Valor Compras") = Format(rsEstoque("Valor Compras") + nPrecoFinal, "############0.00")
            Case "T"
              rsEstoque("Transf Entra") = rsEstoque("Transf Entra") + nQtde
              rsEstoque("Valor T Entra") = Format(rsEstoque("Valor T Entra") + nPrecoFinal, "############0.00")
            Case "A"
              rsEstoque("Ajuste Entra") = rsEstoque("Ajuste Entra") + nQtde
              rsEstoque("Valor Ajuste Entra") = Format(rsEstoque("Valor Ajuste Entra") + nPrecoFinal, "############0.00")
            Case "G"
              rsEstoque("Grátis Entra") = rsEstoque("Grátis Entra") + nQtde
              rsEstoque("Valor Grátis Entra") = Format(rsEstoque("Valor Grátis Entra") + nPrecoFinal, "############0.00")
            Case "E"
              rsEstoque("Empre Entra") = rsEstoque("Empre Entra") + nQtde
              rsEstoque("Valor Empre Entra") = Format(rsEstoque("Valor Empre Entra") + nPrecoFinal, "############0.00")
            Case "D"
              rsEstoque("Devolução") = rsEstoque("Devolução") + nQtde
              rsEstoque("Valor Devolução") = Format(rsEstoque("Valor Devolução") + nPrecoFinal, "############0.00")
              
              '08/08/2003 - maikel
              '             Comentadas as duas linhas abaixo para resolver o problema de estoque (referente a devolução) do quick store
'              rsEstoque("Vendas") = rsEstoque("Vendas") - nQtde
'              rsEstoque("Valor Vendas") = Format(rsEstoque("Valor Vendas") - nPrecoFinal, "############0.00")
          End Select
      
          Estoque2 = rsEstoque("Estoque Anterior")
          Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
          Estoque_Final = Estoque_Final - rsEstoque("Transf Saída") + rsEstoque("Transf Entra")
          Estoque_Final = Estoque_Final - rsEstoque("Ajuste Saída") + rsEstoque("Ajuste Entra")
          Estoque_Final = Estoque_Final - rsEstoque("Grátis Saída") + rsEstoque("Grátis Entra")
          Estoque_Final = Estoque_Final - rsEstoque("Empre Saída") + rsEstoque("Empre Entra")
          
          '08/08/2003 - maikel
          '             Descomentada a soma da coluna desconto para resolver o problema de estoque
          Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolução")
    
          If rsProdutos("Estoque") = False Then
            Estoque_Final = 0
          End If
    
          rsEstoque("Estoque Final") = Estoque_Final
          rsEstoque.Update
          
          Rem Arruma Estoque Final
          '''Grava_Estoque_Final gnCodFilial, Cód, Tamanho, Cor, Edição, CSng(Estoque_Final), dtData
          Grava_Estoque_Final gnCodFilial, Cód, Tamanho, Cor, Edição, CSng(Estoque_Final), Format(Now, "dd/mm/yyyy")
          
        End If
        
        '---------------------------------------------------------------------------
        '17/05/2005 - mpdea
        'Cálculo do custo com aplicação do valor de ICMS Retido
        Dim p_tmp_ValorIcmsRetido As Double
        If IsNumeric(rsEntraProd.Fields("ValorIcmsRetido").Value) Then
          p_tmp_ValorIcmsRetido = CDbl(rsEntraProd.Fields("ValorIcmsRetido").Value)
        Else
          p_tmp_ValorIcmsRetido = 0
        End If
        
        dblPrecoFinalCusto = CDbl(nPrecoFinal) + p_tmp_ValorIcmsRetido
        
        '14/06/2006 - mpdea
        'Aplica frete no custo se o percentual calculado for maior do que zero
        'e não houver valor de icms retido, pois o frete já foi adicionado anteriormente
        If sngPercFrete > 0 And p_tmp_ValorIcmsRetido = 0 Then
          dblPrecoFinalCusto = dblPrecoFinalCusto * (1 + sngPercFrete)
        End If
        
        
        'Calcula Custo Médio e Grava
        If rsOp_Entrada("Tipo") = "C" Then
          
          Custo_Médio = IIf(IsNull(rsProdutos("Custo Médio")), 0, rsProdutos("Custo Médio")) * Estoque2
          
          '26/08/2003 - mpdea
          'Calcula com o preço final (IPI, desconto)
          Custo_Médio = Custo_Médio + dblPrecoFinalCusto 'nQtde * nPreco
          
          If (Estoque2 + nQtde) <> 0 Then
            Custo_Médio = Custo_Médio / (Estoque2 + nQtde)
          End If
          If (Estoque2 + nQtde) = 0 Then
            Custo_Médio = nPreco
          End If
          If Estoque2 < 0 Then
            Custo_Médio = nPreco
          End If
          
          With rsProdutos
            .Edit
            .Fields("Última Compra").Value = Format$(dtData, "dd/mm/yyyy")
            .Fields("Último Custo").Value = Format(dblPrecoFinalCusto / CDbl(nQtde), FORMAT_VALUE)
            .Fields("Custo Médio").Value = Format(Custo_Médio, FORMAT_VALUE)
            .Fields("Último Fornecedor").Value = nCodFornecedor
            '16/11/2004 - Daniel
            'Adicionado tratamento caso a moeda
            'seja nula do produto, colocaremos para
            'igual a 1 (Real)
            'Case: Nazareno, não foi identificado como mas
            'alguns produtos estavam com moeda = 0
            If .Fields("Moeda").Value = 0 Then .Fields("Moeda").Value = 1
            
            .Update
          End With
          blnWEBSynchronize = True
        
        End If
        
        '22/09/2005 - mpdea
        'Grava Custo para Preço de lista sem IPI
        'Utilizado na pasta Cálculos do Produto
        If rsOp_Entrada.Fields("GravaCustoPrecoListaSemIPI").Value Then
          With rsProdutos
            .Edit
            '.Fields("Custo Preço Valor").Value = Format(nPrecoFinal / CDbl(nQtde), FORMAT_VALUE)
            .Fields("Custo Preço Valor").Value = Format(dblPrecoFinalCusto / CDbl(nQtde), FORMAT_VALUE)
            .Update
          End With
          blnWEBSynchronize = True
        End If
        
        '30/08/2007 - Anderson
        'Implementação do campo para automatização do preço de custo.
        'Solicitante: Candy Clean
        If rsOp_Entrada.Fields("PrecoCustoCalculado").Value Then
          Calcula_Custo dblPrecoCustoCalculado, rsProdutos("Custo Desconto Fixo").Value, rsProdutos("Custo Desconto Valor").Value, rsProdutos("Custo Desconto Perc").Value, rsProdutos("Custo Preço Valor").Value, rsProdutos("Custo Frete Fixo").Value, rsProdutos("Custo Frete Valor").Value, rsProdutos("Custo Frete Perc").Value, rsProdutos("Custo ICM Compra Fixo").Value, rsProdutos("Custo ICM Compra Valor").Value, rsProdutos("Custo ICM Compra Perc").Value, rsProdutos("Custo IPI Compra Fixo").Value, rsProdutos("Custo IPI Compra Valor").Value, rsProdutos("Custo IPI Compra Perc").Value, rsProdutos("Custo Custo Finan Fixo").Value, rsProdutos("Custo Custo Finan Valor").Value, rsProdutos("Custo Custo Finan Perc").Value, rsProdutos("Custo Outros Compra Fixo").Value, rsProdutos("Custo Outros Compra Valor").Value, rsProdutos("Custo Outros Compra Perc").Value
        End If
        '---------------------------------------------------------------------------
        
      
        Rem Diminui comissão se for devolução
        If rsOp_Entrada("Tipo") = "D" Then
          If rsOp_Entrada("Comissão") = True Then
            Comissão = rsFuncionarios("Comissão")
            If rsProdutos("Comissão Sobrepõe") = True Then
              Comissão = rsProdutos("Comissão")
            End If
'              Comissão = Comissão * rsTabelas("Multiplicador Comissão")
            Comissão = Format(Comissão, "#############0.00")
            rsComissões.AddNew
            rsComissões("Data") = dtData
            rsComissões("Vendedor") = nCodDigitador
            rsComissões("Produto") = Cód
            rsComissões("Tamanho") = Tamanho
            rsComissões("Cor") = Cor
            '31/03/2005 - Daniel
            'Antiga linha abaixo, estava sem o sinal de menos (-)
            'rsComissões("Qtde") = nQtde
            'A partir da 6.52.0.28 contemplou esta alteração
            rsComissões("Qtde") = -nQtde
            rsComissões("Valor") = -nPrecoFinal
            rsComissões("Sequência") = Mov
            '14/02/2005 - Daniel
            '
            'Solicitante: Daring - RJ
            '
            'Se ocorre devolução e esta devolução implica em abatimento de
            'comissão do vendedor, o Quick estava descontando erroneamente
            'da comissão para casos em que a venda possuia descontos.
            If (Len(frmEntrada.gsTabelaVenda) & "") > 0 Then 'Foi preenchida a var global...
              Dim dblValorDoCadastroProduto As Double
              Dim strCodProdSemGrade        As String
              Dim rstTabelaPrecos           As Recordset
              Dim sngPercentComisDesconto   As Single
              Dim dblValorComissao          As Double
              
              Set rstTabelaPrecos = db.OpenRecordset("SELECT PercentualComissaoDesconto FROM [Tabela de Preços] WHERE Tabela = '" & frmEntrada.gsTabelaVenda & "'", dbOpenSnapshot)
              
              If rstTabelaPrecos.RecordCount > 0 Then
                
                Call ReduzirComissao(frmEntrada.gsTabelaVenda & "", Cód & "", dblValorDoCadastroProduto)
              
                If Not IsNull(rstTabelaPrecos.Fields("PercentualComissaoDesconto")) Then
                  sngPercentComisDesconto = rstTabelaPrecos.Fields("PercentualComissaoDesconto")
                Else
                  sngPercentComisDesconto = 0
                End If
              
                'Se for diferente ocorre o abatimento pela metade ou percentual da comissão
                If dblValorDoCadastroProduto <> Format((nPrecoFinal / nQtde), FORMAT_VALUE) Then
                  
                  dblValorComissao = (rsComissões("Valor") * Comissão / 100)
                  dblValorComissao = dblValorComissao * ((100 - sngPercentComisDesconto) / 100)
                  
                  rsComissões("Comissão") = Truncate(dblValorComissao, 6)
                  
                Else
                  rsComissões("Comissão") = CCur(Format((rsComissões("Valor") * Comissão / 100), "###########0.00"))
                End If
                
              Else
                rsComissões("Comissão") = CCur(Format((rsComissões("Valor") * Comissão / 100), "###########0.00"))
              End If
            
              rstTabelaPrecos.Close
              Set rstTabelaPrecos = Nothing
              
            Else
              rsComissões("Comissão") = CCur(Format((rsComissões("Valor") * Comissão / 100), "###########0.00"))
            End If
            rsComissões("Filial") = gnCodFilial
            rsComissões("Cliente") = nCodFornecedor
            rsComissões.Update
          End If
        End If
        
      
        Rem Grava etiquetas
        If bEtiqueta = True Then
          rsEtiquetas.Index = "Funcionário"
          rsEtiquetas.Seek "=", nCodDigitador, Cód, Tamanho, Cor
          If rsEtiquetas.NoMatch Then
             rsEtiquetas.AddNew
          Else
             rsEtiquetas.Edit
          End If
          rsEtiquetas("Funcionário") = nCodDigitador
          rsEtiquetas("Produto") = Cód
          rsEtiquetas("Tamanho") = Tamanho
          rsEtiquetas("Cor") = Cor
          rsEtiquetas("Qtde") = rsEtiquetas("Qtde") + nQtde
          rsEtiquetas("Sequência") = Mov
          rsEtiquetas.Update
        End If
      
      
        'Atualiza preço de custo na tabela  CUSTO
        'quando for Compra
        rsPreços.Index = "Tabela"
        If rsOp_Entrada("Gravar Custo") = True Then
          rsPreços.Seek "=", "CUSTO", rsProdutos("Código")
          If rsPreços.NoMatch Then
             rsPreços.AddNew
             rsPreços("Tabela") = "CUSTO"
             rsPreços("Produto") = rsProdutos("Código")
          Else
             rsPreços.Edit
          End If
      
          '03/11/2004 - Daniel
          'Tratamento para quando o produto tiver
          'preço em Dólar
          If MoedaReal(rsProdutos("Moeda").Value) Then
            rsPreços("Preço") = CSng(dblPrecoFinalCusto) / nQtde
            
            '30/08/2007 - Anderson
            'Implementação do campo para automatização do preço de custo.
            'Solicitante: Candy Clean
            If rsOp_Entrada.Fields("PrecoCustoCalculado").Value Then
              rsPreços("Preço") = dblPrecoCustoCalculado
            End If

            rsPreços.Update
          Else 'Produto em Dólar ou em outra moeda
            Dim dblCotacao As Double
            
            Call BuscarUltimaCotacao(rsProdutos("Moeda").Value, dblCotacao)
            
            '07/06/2005 - Daniel
            'Correção do bug: Impossível divisão por zero
            'Adicionado tratamento para evitar erro na divisão por zero
            'Este erro aparecia quando não havia cotação alguma cadastrada
            If dblCotacao > 0 Then
              rsPreços("Preço") = Format(CSng((CSng(dblPrecoFinalCusto) / nQtde / dblCotacao)), FORMAT_VALUE)
              
              '30/08/2007 - Anderson
              'Implementação do campo para automatização do preço de custo.
              'Solicitante: Candy Clean
              If rsOp_Entrada.Fields("PrecoCustoCalculado").Value Then
                rsPreços("Preço") = Format(CSng((CSng(dblPrecoCustoCalculado) / dblCotacao)), FORMAT_VALUE)
              End If

            End If
            
            rsPreços.Update
            
          End If
          
          '-------------------------------------------------------------------------
          '05/08/2003 - mpdea
          'Grava informações de custo para operações diferentes de Compra
          'Obs.: Operação de Compra já possui tratamento próprio
          If rsOp_Entrada.Fields("Tipo").Value <> "C" Then
            
            Custo_Médio = IIf(IsNull(rsProdutos("Custo Médio")), 0, rsProdutos("Custo Médio")) * Estoque2
            
            '26/08/2003 - mpdea
            'Calcula com o preço final (IPI, desconto)
            Custo_Médio = Custo_Médio + nPrecoFinal 'nQtde * nPreco
            
            If (Estoque2 + nQtde) <> 0 Then
              Custo_Médio = Custo_Médio / (Estoque2 + nQtde)
            End If
            If (Estoque2 + nQtde) = 0 Then
              Custo_Médio = nPreco
            End If
            If Estoque2 < 0 Then
              Custo_Médio = nPreco
            End If
            
            With rsProdutos
              .Edit
              .Fields("Último Custo").Value = Format(nPrecoFinal / CDbl(nQtde), FORMAT_VALUE)
              .Fields("Custo Médio").Value = Format(Custo_Médio, FORMAT_VALUE)
              .Update
            End With
          End If
          '-------------------------------------------------------------------------
          
          
          blnWEBSynchronize = True
          
        End If
      
        
        Rem Atualiza arquivo de Resumo de Clientes
        Rem se for Compra
        If rsOp_Entrada("Tipo") = "C" Or rsOp_Entrada("Tipo") = "D" Then
          
          rsResumo_Clientes.Index = "Cliente"
          rsResumo_Clientes.Seek "=", nCodFornecedor, dtData, Cód, Tamanho, Cor, Edição, Mov
          If rsResumo_Clientes.NoMatch Then
             rsResumo_Clientes.AddNew
          Else
             rsResumo_Clientes.Edit
          End If
           
          rsResumo_Clientes("Dia") = dtData
          rsResumo_Clientes("Cliente") = nCodFornecedor
          rsResumo_Clientes("Produto") = Cód
          rsResumo_Clientes("Tamanho") = Tamanho
          rsResumo_Clientes("Cor") = Cor
          rsResumo_Clientes("Edição") = Edição
          rsResumo_Clientes("Tipo") = "F"
          
          If rsOp_Entrada("Tipo") = "C" Then
            rsResumo_Clientes("Qtde") = nQtde
            rsResumo_Clientes("Valor Total") = Format(nPrecoFinal, "############0.00")
          End If
          If rsOp_Entrada("Tipo") = "D" Then
            rsResumo_Clientes("Qtde") = -nQtde
            rsResumo_Clientes("Valor Total") = -Format(nPrecoFinal, "############0.00")
            rsResumo_Clientes("Tipo") = "C"
          End If
          
          rsResumo_Clientes("Sequência") = Mov
          rsResumo_Clientes("Filial") = gnCodFilial
  
          rsResumo_Clientes.Update
          
        End If
      
      
      
        Rem Atualiza arquivo de Empréstimos
        If rsOp_Entrada("Tipo") = "E" And Not rsEntraProd("InGeradoViaConsig") Then
        
          Rem Saldo Emprestado = 0 para este empréstimo
          rsEmprestimos.Index = "Cliente"
          Saldo_Emp = 0
          Ordem_Emp = Ordem
          Emp_Existe = False
                   
          rsEmprestimos.Seek "<", gnCodFilial, Mov, nCodFornecedor, Cód, Tamanho, Cor, Edição, 999999
          If Not rsEmprestimos.NoMatch Then
            If rsEmprestimos("Filial") = gnCodFilial Then
              If rsEmprestimos("Sequência") = Mov Then
                If rsEmprestimos("Fornecedor") = nCodFornecedor Then
                  If rsEmprestimos("Produto") = Cód Then
                    If rsEmprestimos("Tamanho") = Tamanho Then
                      If rsEmprestimos("Cor") = Cor Then
                        If rsEmprestimos("Edição") = Edição Then
                           Ordem_Emp = rsEmprestimos("Ordem")
                           Emp_Existe = True
                           Saldo_Emp = rsEmprestimos("Saldo Atual")
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
                   
          If Emp_Existe = True Then
            rsEmprestimos.Edit
          Else
            rsEmprestimos.AddNew
            rsEmprestimos("Filial") = gnCodFilial
            rsEmprestimos("Sequência") = Mov
            rsEmprestimos("Fornecedor") = nCodFornecedor
            rsEmprestimos("Produto") = Cód
            rsEmprestimos("Tamanho") = Tamanho
            rsEmprestimos("Cor") = Cor
            rsEmprestimos("Edição") = Edição
            rsEmprestimos("Ordem") = Ordem_Emp
          End If
              
          rsEmprestimos("Saldo Anterior") = Saldo_Emp
          rsEmprestimos("Empréstimo Recebido") = nQtde
          rsEmprestimos("Saldo Atual") = Saldo_Emp + nQtde
          rsEmprestimos("Preço Unitário") = (nPrecoFinal / nQtde)
          rsEmprestimos("Data Operação") = dtData
          rsEmprestimos("Data Alteração") = Format(Date, "dd/mm/yyyy")
          If dtDataAcerto <> vbNull Then
            rsEmprestimos("Data Cobrança") = dtDataAcerto
          End If
    
          rsEmprestimos.Update
             
        End If
        
        If blnWEBSynchronize Then
          'Atualiza o sincronismo para o produto WEB alterado
          Call WEB_SynchronizeProduct(rsProdutos("Código").Value)
        End If
        
      End If  'Erro = 0
    End If
    
    'Próximo produto
    rsEntraProd.MoveNext
    
  Loop

  With rsEntradas
    .Edit
    .Fields("Efetivada").Value = True
    .Update
  End With
  
  Efetiva_Entrada = 0
  Call StatusMsg("")
   
  rsEntradas.Close
  rsEntraProd.Close
  rsOp_Entrada.Close
  rsFuncionarios.Close
  rsCliFor.Close
  rsContas_Pagar.Close
  rsProdutos.Close
  rsResumo_Diário.Close
  rsEtiquetas.Close
  rsCaixa.Close
  rsRes_Financeiro.Close
  
  '07/01/2010 - Andrea
  rsContas_Receber.Close
  rsMovi_Cheques.Close
  
  '27/06/2003 - mpdea
  'Verifica se o objeto foi criado antes de fechá-lo
  If Not rsEstoque Is Nothing Then
    rsEstoque.Close
  End If
  
  rsPreços.Close
  rsResumo_Clientes.Close
  rsEmprestimos.Close
  rsMovi_Parcelas.Close
  rsLançamentos.Close
  rsComissões.Close
  
  Set rsEntradas = Nothing
  Set rsEntraProd = Nothing
  Set rsOp_Entrada = Nothing
  Set rsFuncionarios = Nothing
  Set rsCliFor = Nothing
  Set rsContas_Pagar = Nothing
  Set rsProdutos = Nothing
  Set rsResumo_Diário = Nothing
  Set rsEtiquetas = Nothing
  Set rsCaixa = Nothing
  Set rsRes_Financeiro = Nothing
  Set rsEstoque = Nothing
  Set rsPreços = Nothing
  Set rsResumo_Clientes = Nothing
  Set rsEmprestimos = Nothing
  Set rsMovi_Parcelas = Nothing
  Set rsLançamentos = Nothing
  Set rsComissões = Nothing
   
  '07/01/2010 - Andrea
  Set rsContas_Receber = Nothing
  Set rsMovi_Cheques = Nothing
   
  Exit Function

Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case Err.Number
    Case 3186, 3197, 3187, 3218, 3260 'Registro bloqueado
      If nRepeatUpdateLocked < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - nRepeatUpdateLocked)
        nRepeatUpdateLocked = nRepeatUpdateLocked + 1
        Call WaitSeconds(1) 'Aguarda um segundo
        Resume
      Else
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          nRepeatUpdateLocked = 0
          Screen.MousePointer = vbHourglass
          Resume
        Else
          Efetiva_Entrada = -1 'Ação cancelada
          Exit Function
        End If
      
'        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Efetiva Entrada") = vbOK Then
'          nRepeatUpdateLocked = 0
'          Resume
'        Else
'          Efetiva_Entrada = -1 'Ação cancelada
'          Exit Function
'        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Efetiva Entrada")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Efetiva_Entrada = -1 'Ação cancelada
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select

End Function

Private Function MoedaReal(ByVal Moeda As Byte) As Boolean
  '03/11/2004 - Daniel
  'Verificação do tipo de moeda que implicará no cálculo
  'da tabela CUSTO
  Dim rstMoedas As Recordset
  Dim strMoeda  As String
  
  Set rstMoedas = db.OpenRecordset("SELECT * FROM Moedas WHERE Código = " & Moeda, dbOpenDynaset)
  
  With rstMoedas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      strMoeda = .Fields("Nome").Value & ""
      strMoeda = UCase(strMoeda)
      
      If strMoeda = "REAL" Then MoedaReal = True
      
    End If
    .Close
  End With
  
  Set rstMoedas = Nothing
  
End Function

Private Sub BuscarUltimaCotacao(ByVal Moeda As Byte, ByRef Cotacao As Double)
  Dim rstCotacoes As Recordset
  
  Cotacao = 0
  
  Set rstCotacoes = db.OpenRecordset("SELECT * FROM Cotações WHERE Moeda = " & Moeda & " ORDER BY Data ", dbOpenDynaset)

  With rstCotacoes
    If Not (.BOF And .EOF) Then
      .MoveLast
      
      Cotacao = Format(.Fields("Cotação").Value, FORMAT_VALUE)
      
    End If
    .Close
  End With
  
  Set rstCotacoes = Nothing

End Sub

Private Sub ReduzirComissao(ByVal Tabela As String, ByVal Produto As String, ByRef dblValorDoCadastroProduto As Double)
  '14/02/2005 - Daniel
  'Problema levantado pela Daring
  'Se ocorre devolução e esta devolução implica em abatimento de
  'comissão do vendedor, o Quick estava descontando erroneamente
  'da comissão para casos em que a venda possuia descontos.
  Dim rstPrecos As Recordset
  Dim strSQL    As String
  
  strSQL = "SELECT Preço FROM Preços "
  strSQL = strSQL & " WHERE Produto = '" & Produto & "'"
  strSQL = strSQL & " AND Tabela = '" & Tabela & "'"
  
  Set rstPrecos = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstPrecos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      dblValorDoCadastroProduto = .Fields("Preço").Value
    End If
    .Close
  End With
  
  Set rstPrecos = Nothing
  
End Sub
