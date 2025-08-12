Attribute VB_Name = "modEfetivaEntrada"
Option Explicit

Public Function Desefetiva_Entrada(Filial As Integer, Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  'Inserido os recordsets que estavam a n�vel modular sem necessidade,
  'ocupando mais mem�ria
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
'  Dim rsParametros As Recordset
  Dim rsOp_Entrada As Recordset
  Dim rsContas_Pagar As Recordset
  Dim rsResumo_Di�rio As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
  Dim rsEstoque_Final As Recordset
  Dim rsPre�os As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
'  Dim rsGrade As Recordset
  Dim rsEntradas As Recordset
  Dim rsEntra_Prod As Recordset
  Dim rsEntraProd As Recordset
  Dim rsMovi_Parcelas As Recordset
  Dim rsLan�amentos As Recordset
  Dim rsComiss�es As Recordset
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
  Dim Tot_Cart�es As Double
  Dim Tot_Vales As Double
  Dim Tot_Parcelamento As Double
  Dim C�d As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Aux_Prod As String
  Dim Aux_Tipo As Integer
  Dim Aux_Erro As Integer
  Dim Edi��o As Long
  Dim Estoque_Final As Double
  Dim Estoque2 As Double
  Dim Custo_M�dio As Double
  Dim Criar_Registro As Integer
  Dim Mensagem As String
  Dim Saldo_Conta As Double
  Dim Aux_Sequ�ncia As Long
  
  'Vari�vel de Tratamento de Erro
  Dim nRepeatUpdateLocked As Integer
  
  Dim strSQL As String
  
  On Error GoTo Processa_Erro
  
  Screen.MousePointer = vbHourglass
  
  Set rsEntradas = db.OpenRecordset("Entradas")
  Set rsContas_Pagar = db.OpenRecordset("Contas a Pagar")
  Set rsProdutos = db.OpenRecordset("Produtos")
'  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  Set rsOp_Entrada = db.OpenRecordset("Opera��es Entrada", , dbReadOnly)
  Set rsResumo_Di�rio = db.OpenRecordset("Resumo Di�rio")
  Set rsEtiquetas = db.OpenRecordset("Etiquetas")
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsRes_Financeiro = db.OpenRecordset("Resumo Di�rio Financeiro")
'  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsPre�os = db.OpenRecordset("Pre�os")
  Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
  Set rsEmprestimos = db.OpenRecordset("Consigna��o Entrada")
  Set rsCliFor = db.OpenRecordset("Cli_For")
'  Set rsGrade = db.OpenRecordset("C�digos da Grade", , dbReadOnly)
  Set rsEntra_Prod = db.OpenRecordset("Entradas - Produtos", , dbReadOnly)
  Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsLan�amentos = db.OpenRecordset("Lan�amentos Banc�rios")
  Set rsComiss�es = db.OpenRecordset("Comiss�o")
  
  '08/01/2010 - Andrea
  Set rsMovi_Cheques = db.OpenRecordset("Movimento - Cheques")
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
  
  Screen.MousePointer = vbDefault
  
  rsEntradas.Index = "Sequ�ncia"
  rsEntradas.Seek "=", Filial, Mov
  If rsEntradas.NoMatch Then
    Desefetiva_Entrada = 1
    Exit Function
  End If
  
  
  ' Encontrou Movimenta��o
  ' Separa mes e ano
  Aux_Str = Format$(rsEntradas("Data"), "dd/mm/yyyy")
  Ano_Atual = Val(Right(Aux_Str, 4))
  Mes_Atual = Val(Mid(Aux_Str, 4, 2))
  
  ' Encontra a tabela de opera��es
  rsOp_Entrada.Index = "C�digo"
  rsOp_Entrada.Seek "=", rsEntradas("Opera��o")
  If rsOp_Entrada.NoMatch Then
    Desefetiva_Entrada = 2
    Exit Function
  End If
  
  ' Encontra Fornecedor
  rsCliFor.Index = "C�digo"
  rsCliFor.Seek "=", rsEntradas("Fornecedor")
  If rsCliFor.NoMatch Then
    Desefetiva_Entrada = 3
    Exit Function
  End If
  
  
  ' Diminui comiss�o se for devolu��o
  If rsOp_Entrada("Tipo") = "D" Then
    Aux_Sequ�ncia = 0
    Erro = False
    rsComiss�es.Index = "Sequ�ncia"
    Do
      rsComiss�es.Seek ">", Filial, Mov, 0
      If rsComiss�es.NoMatch Then Erro = True
      If Erro = False Then If rsComiss�es("Filial") <> Filial Then Erro = True
      If Erro = False Then If rsComiss�es("Sequ�ncia") <> Mov Then Erro = True
      If Erro = False Then
        rsComiss�es.Delete
      End If
    Loop Until Erro = True
  End If
  
  
  ' Atualiza arquivo de Resumo de Clientes
  ' se for Compra
  Erro = False
  rsResumo_Clientes.Index = "Sequ�ncia"
  Do
    rsResumo_Clientes.Seek ">=", Filial, Mov
    If rsResumo_Clientes.NoMatch Then Erro = True
    If Erro = False Then If rsResumo_Clientes("Filial") <> Filial Then Erro = True
    If Erro = False Then If rsResumo_Clientes("Sequ�ncia") <> Mov Then Erro = True
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
    ' Neste ponto tem o �ltimo caixa no buffer
    ' Acha parcela a vista
    Ordem = rsCaixa("Ordem")
    Ordem = Ordem + 1
    Saldo_Ant = rsCaixa("Final")
    Tot_Dinheiro = rsCaixa("Total Dinheiro")
    Tot_Cheques = rsCaixa("Total Cheques")
    Tot_Cheques_Pre = rsCaixa("Total Cheques Pr�")
    Tot_Cart�es = rsCaixa("Total Cart�es")
    Tot_Vales = rsCaixa("Total Vales")
    Tot_Parcelamento = rsCaixa("Total Parcelamento")
  
    With rsCaixa
      .AddNew
      .Fields("Filial") = Filial
      .Fields("Data") = rsEntradas("Data")
      .Fields("Hora") = Format(Time, "hh:mm:ss")
      .Fields("Caixa") = rsEntradas("Caixa")
      .Fields("Ordem") = Ordem
      .Fields("Descri��o") = "Cancelamento entrada " & str(Mov)
      .Fields("Saldo Anterior") = Saldo_Ant
      
      '12/01/2010 - Andrea
      .Fields("Total Cheques Pr�") = Tot_Cheques_Pre + rsEntradas("Cheque Caixa")
      '.Fields("Cheques") = rsEntradas("Cheque Caixa")
      .Fields("Cheques") = 0
      .Fields("Cheques Pr�") = rsEntradas("Cheque Caixa")
      
      .Fields("Total Cart�es") = Tot_Cart�es
      .Fields("Total Vales") = Tot_Vales
      .Fields("Total Cheques") = Tot_Cheques
      .Fields("Total Parcelamento") = Tot_Parcelamento
      .Fields("Dinheiro") = (rsEntradas("Dinheiro Caixa") - rsEntradas("Troco"))
      .Fields("Total Dinheiro") = Tot_Dinheiro + rsEntradas("Dinheiro Caixa") - rsEntradas("Troco")
      .Fields("Final") = Tot_Dinheiro + rsEntradas("Dinheiro Caixa") + Tot_Cheques + Tot_Cheques_Pre + Tot_Cart�es + Tot_Vales - rsEntradas("Troco") + rsEntradas("Cheque Caixa")
      
      .Update
    End With
    
  End If
  
  
  ' Faz Lancamento na conta banc�ria, se for o caso
  If rsEntradas("Valor Cheque") <> 0 Then
    ' Acha Saldo Anterior
    Saldo_Conta = 0
    rsLan�amentos.Index = "Conta"
    rsLan�amentos.Seek "<", rsEntradas("Conta"), rsEntradas("data"), 99999999#
    If Not rsLan�amentos.NoMatch Then
      If rsLan�amentos("Conta") = rsEntradas("Conta") Then
        Saldo_Conta = rsLan�amentos("Saldo Atual")
      End If
    End If
    
    With rsLan�amentos
      .AddNew
      .Fields("Conta") = rsEntradas("Conta")
      .Fields("Data") = rsEntradas("Data")
      .Fields("Descri��o") = "Cancelamento entrada " + str(rsEntradas("Sequ�ncia"))
      .Fields("Cheque") = rsEntradas("Num Cheque")
      .Fields("Cr�dito") = rsEntradas("Valor Cheque")
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
    If Erro = False Then If rsMovi_Cheques("Sequ�ncia") <> Mov Then Erro = True

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
  Aux_Sequ�ncia = 0
  rsContas_Pagar.Index = "Sequ�ncia"
  Do
    rsContas_Pagar.Seek ">", Filial, Mov, Aux_Sequ�ncia
    If rsContas_Pagar.NoMatch Then Erro = True
    If Erro = False Then If rsContas_Pagar("Filial") <> Filial Then Erro = True
    If Erro = False Then If rsContas_Pagar("Sequ�ncia") <> Mov Then Erro = True
    
    If Erro = False Then
      rsContas_Pagar.Delete
    End If
  Loop Until Erro = True
  
  
  
  ' Atualiza Resumo Di�rio
  If rsOp_Entrada("Tipo") <> "P" Then
    rsResumo_Di�rio.Index = "Data"
    rsResumo_Di�rio.Seek "=", Filial, rsEntradas("Data")
    With rsResumo_Di�rio
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
          .Fields("Valor Devolu��o") = CDbl(.Fields("Valor Devolu��o")) - CDbl(rsEntradas("Total"))
          
          '08/08/2003 - maikel
          '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolu��o) do quick store
          '.Fields("Valor Vendas") = CDbl(.Fields("Valor Vendas")) + CDbl(rsEntradas("Total"))
      End Select
      .Update
    End With
  End If
  
  
  ' Atualiza Resumo Di�rio Financeiro
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
         .Fields("Valor Devolu��o") = CDbl(.Fields("Valor Devolu��o")) - CDbl(rsEntradas("Total"))
          
          '08/08/2003 - maikel
          '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolu��o) do quick store
         '.Fields("Valor Vendas") = CDbl(.Fields("Valor Vendas")) + CDbl(rsEntradas("Total"))
      End Select
      .Update
    End With
  End If
  
  
  ' Desfaz Empr�stimos
  rsEmprestimos.Index = "Cliente"
  Erro = False
  Do
    rsEmprestimos.Seek ">", rsEntradas("Filial"), rsEntradas("Sequ�ncia"), 0, 0, 0, 0, 0, 0
    If rsEmprestimos.NoMatch Then Erro = True
    If Erro = False Then If rsEntradas("Filial") <> rsEmprestimos("Filial") Then Erro = True
    If Erro = False Then If rsEntradas("Sequ�ncia") <> rsEmprestimos("Sequ�ncia") Then Erro = True
    If Erro = False Then
      rsEmprestimos.Delete
    End If
  Loop Until Erro = True
  
  
  
  rsEntra_Prod.Index = "Sequ�ncia"
  Ordem = 0
Prox_Prod:
  rsEntra_Prod.Seek ">", Filial, Mov, Ordem
  
  If rsEntra_Prod.NoMatch Then GoTo Fim_Desefetiva
  If rsEntra_Prod("Filial") <> Filial Then GoTo Fim_Desefetiva
  If rsEntra_Prod("sequ�ncia") <> Mov Then GoTo Fim_Desefetiva
  
  Ordem = rsEntra_Prod("Linha")
  'Verifica se tem grade
  C�d = rsEntra_Prod("C�digo")
  Tamanho = 0
  Cor = 0
  
  rsProdutos.Index = "C�digo"
  
  Aux_Prod = rsEntra_Prod("C�digo")
  Acha_Produto Aux_Prod, C�d, Tamanho, Cor, Edi��o, Aux_Tipo, Aux_Erro
  If Aux_Erro <> 0 Then
   'Call StatusMsg("Produto n�o encontrado."
   GoTo Prox_Prod
  End If
  C�d = UCase(C�d)
     
  rsProdutos.Seek "=", C�d
  If rsProdutos.NoMatch Then
   GoTo Prox_Prod
  End If
  
  
  'Neste ponto C�D tem o c�digo do produto
  'Tamanho e Cor cont�m os respectivos dados
  'Agora grava arquivo do estoque
  
  '  Ajusta Estoque
  If rsOp_Entrada("Estoque") = True Then
    Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
    
'-------------------------------------------------------------------------------------
    '31/03/2004 - mpdea
    'Modificado par�metro de abertura do recordset
    'dbOpenSnapshot (muito lento!? 8-|) para dbOpenDynaset com dbReadOnly
    'e modificado para que salve somente no final da atualiza��o
    'de estoque o recordset
    '
    '10/10/2003 - Maikel
    '             Modificada a forma de analisar a tabela de estoque. Da forma antiga gerava erro 3022 ao efetuar movimenta��o com data retroativa.
    strSQL = "SELECT * FROM Estoque WHERE " & _
             " Filial = " & Filial & _
             " AND Produto = '" & C�d & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edi��o = " & Edi��o & _
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
             " AND Produto = '" & C�d & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edi��o = " & Edi��o & _
             " AND Data = #" & Format(rsEntradas("Data"), "mm/dd/yyyy") & "#"
             
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .Edit
      Else
        .AddNew
        .Fields("Filial").Value = Filial
        .Fields("Data").Value = rsEntradas("Data").Value
        .Fields("Produto").Value = C�d
        .Fields("Tamanho").Value = Tamanho
        .Fields("Cor").Value = Cor
        .Fields("Edi��o").Value = Edi��o
        .Fields("Classe").Value = rsProdutos("Classe").Value
        .Fields("Sub Classe").Value = rsProdutos("Sub Classe").Value
        .Fields("Estoque Anterior").Value = Estoque_Final
'        .Update
'        .Requery
      End If
    End With
'-------------------------------------------------------------------------------------

    
'    ' Encontra a posi��o do estoque
'    Criar_Registro = False
'    Estoque_Final = 0
'    rsEstoque.Index = "Produto"
'    rsEstoque.Seek "=", Filial, rsEntradas("Data"), C�d, Tamanho, Cor, Edi��o
'
'    With rsEstoque
'
'      If Not .NoMatch Then
'        Estoque_Final = .Fields("Estoque Final")
'      Else
'        .Index = "Data"
'        .Seek "<", Filial, C�d, Tamanho, Cor, Edi��o, rsEntradas("Data")
'        If .NoMatch Then Criar_Registro = True
'        If Not .NoMatch Then
'          If .Fields("Filial") = Filial And .Fields("Produto") = C�d And .Fields("Tamanho") = Tamanho And .Fields("Cor") = Cor And .Fields("Edi��o") = Edi��o Then
'            Criar_Registro = True
'            Estoque_Final = .Fields("Estoque Final")
'          End If
'        End If
'
'        .AddNew
'        .Fields("Filial") = Filial
'        .Fields("Data") = rsEntradas("Data")
'        .Fields("Produto") = C�d
'        .Fields("Tamanho") = Tamanho
'        .Fields("Cor") = Cor
'        .Fields("Edi��o") = Edi��o
'        .Fields("Classe") = rsProdutos("Classe")
'        .Fields("Sub Classe") = rsProdutos("Sub Classe")
'        .Fields("Estoque Anterior") = Estoque_Final
'        .Update
'
'        .Index = "Produto"
'        .Seek "=", Filial, rsEntradas("Data"), C�d, Tamanho, Cor, Edi��o
'      End If
      
'-------------------------------------------------------------------------------------
      
      ' neste ponto esta com o registro de estoque
      ' no buffer, agora soma com os valores da movimenta��o
    With rsEstoque
'      .Edit
      Select Case rsOp_Entrada("Tipo")
        Case "C"
          .Fields("Compras") = .Fields("Compras") - rsEntra_Prod("Qtde")
          .Fields("Valor Compras") = Format(.Fields("Valor Compras") - rsEntra_Prod("Pre�o Final"), "###########0.00")
        Case "T"
          .Fields("Transf Entra") = .Fields("Transf Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor T Entra") = Format(.Fields("Valor T Entra") - rsEntra_Prod("Pre�o Final"), "#############0.00")
        Case "A"
          .Fields("Ajuste Entra") = .Fields("Ajuste Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor Ajuste Entra") = Format(.Fields("Valor Ajuste Entra") - rsEntra_Prod("Pre�o Final"), "###############0.00")
        Case "G"
          .Fields("Gr�tis Entra") = .Fields("Gr�tis Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor Gr�tis Entra") = Format(.Fields("Valor Gr�tis Entra") - rsEntra_Prod("Pre�o Final"), "################0.00")
        Case "E"
          .Fields("Empre Entra") = .Fields("Empre Entra") - rsEntra_Prod("Qtde")
          .Fields("Valor Empre Entra") = Format(.Fields("Valor Empre Entra") - rsEntra_Prod("Pre�o Final"), "###############0.00")
        Case "D"
          .Fields("Devolu��o") = .Fields("Devolu��o") - rsEntra_Prod("Qtde")
          .Fields("Valor Devolu��o") = Format(.Fields("Valor Devolu��o") - rsEntra_Prod("Pre�o Final"), "#############0.00")
          
          '08/08/2003 - maikel
          '             Comentadas as duas linhas abaixo para resolver o problema de estoque (referente a devolu��o) do quick store
'          .Fields("Vendas") = .Fields("Vendas") + rsEntra_Prod("Qtde")
'          .Fields("Valor Vendas") = Format(.Fields("Valor Vendas") + rsEntra_Prod("Pre�o Final"), "##############0.00")
      End Select
      
      Estoque2 = .Fields("Estoque Anterior")
      Estoque_Final = .Fields("Estoque Anterior") - .Fields("Vendas") + .Fields("Compras")
      Estoque_Final = Estoque_Final - .Fields("Transf Sa�da") + .Fields("Transf Entra")
      Estoque_Final = Estoque_Final - .Fields("Ajuste Sa�da") + .Fields("Ajuste Entra")
      Estoque_Final = Estoque_Final - .Fields("Gr�tis Sa�da") + .Fields("Gr�tis Entra")
      Estoque_Final = Estoque_Final - .Fields("Empre Sa�da") + .Fields("Empre Entra")
      
      '08/08/2003 - maikel
      '             Descomentada a soma da coluna devolu��o para resolver o problema de estoque
      Estoque_Final = Estoque_Final - .Fields("Quebras") + rsEstoque("Devolu��o")
      
      If rsProdutos("Estoque") = False Then
        Estoque_Final = 0
      End If
      
      .Fields("Estoque Final") = Estoque_Final
      .Update
      .Close
      Call Grava_Estoque_Final(rsEntradas("Filial"), C�d, Tamanho, Cor, Edi��o, CSng(Estoque_Final), rsEntradas("Data"))
    
    End With
    
  End If
  
  
     
  ' apaga etiquetas
  If rsEntra_Prod("Etiqueta") = True Then
   rsEtiquetas.Index = "Funcion�rio"
   rsEtiquetas.Seek "=", rsEntradas("Digitador"), C�d, Tamanho, Cor
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
  'Inclu�do o fechamento dos recordsets abertos e suas desassocia��es
  '---------------------------------------------------------------------------------
  rsEntradas.Close
  rsEntra_Prod.Close
  rsOp_Entrada.Close
  rsCliFor.Close
  rsContas_Pagar.Close
  rsProdutos.Close
  rsResumo_Di�rio.Close
  rsEtiquetas.Close
  rsCaixa.Close
  rsRes_Financeiro.Close
  rsPre�os.Close
  rsResumo_Clientes.Close
  rsEmprestimos.Close
  rsMovi_Parcelas.Close
  rsLan�amentos.Close
  rsComiss�es.Close
  
  Set rsEntradas = Nothing
  Set rsEntra_Prod = Nothing
  Set rsOp_Entrada = Nothing
  Set rsCliFor = Nothing
  Set rsContas_Pagar = Nothing
  Set rsProdutos = Nothing
  Set rsResumo_Di�rio = Nothing
  Set rsEtiquetas = Nothing
  Set rsCaixa = Nothing
  Set rsRes_Financeiro = Nothing
  Set rsEstoque = Nothing
  Set rsPre�os = Nothing
  Set rsResumo_Clientes = Nothing
  Set rsEmprestimos = Nothing
  Set rsMovi_Parcelas = Nothing
  Set rsLan�amentos = Nothing
  Set rsComiss�es = Nothing
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
        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
          "uma nova tentativa.", vbExclamation + vbOKCancel, "Desefetiva Entrada") = vbOK Then
          nRepeatUpdateLocked = 0
          Resume
        Else
          Desefetiva_Entrada = -1 'A��o cancelada
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
          Desefetiva_Entrada = -1 'A��o cancelada
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select
  
End Function

Public Function Efetiva_Entrada(ByVal Filial As Integer, ByVal Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  'Inserido os recordsets que estavam a n�vel modular sem necessidade,
  'ocupando mais mem�ria
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
  Dim rsParametros As Recordset
  Dim rsOp_Entrada As Recordset
  Dim rsContas_Pagar As Recordset
  Dim rsResumo_Di�rio As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
  Dim rsEstoque_Final As Recordset
  Dim rsPre�os As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
  Dim rsGrade As Recordset
  Dim rsEntradas As Recordset
  Dim rsEntra_Prod As Recordset
  Dim rsEntraProd As Recordset
  Dim rsMovi_Parcelas As Recordset
  Dim rsLan�amentos As Recordset
  Dim rsComiss�es As Recordset
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
  Dim Tot_Cart�es As Double
  Dim Tot_Vales As Double
  Dim Tot_Parcela As Double
  Dim C�d As String
  Dim Tamanho As Integer
  Dim Cor As Integer
  Dim Edi��o As Long
  Dim Estoque_Final As Double
  Dim Estoque2 As Double
  Dim Custo_M�dio As Double
  Dim Criar_Registro As Integer
  Dim Saldo_Conta As Double
  Dim Comiss�o As Double
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
   
  'Vari�vel de Tratamento de Erro
  Dim nRepeatUpdateLocked As Integer
  
  'Vari�veis WEB
  Dim blnWEB_Sale As Boolean
  Dim blnWEBSynchronize As Boolean
  
  
  '-------------------------------------------
  '16/05/2006 - mpdea
  'Pre�o para c�lculo de Custo
  Dim dblPrecoFinalCusto As Double
  '-------------------------------------------
  
  
  '14/06/2006 - mpdea
  'Frete
  Dim sngPercFrete As Single
  Dim dblValorFrete As Double
  Dim dblValorTotal As Double
  Dim dblValorIPI As Double
  
  '30/08/2007 - Anderson
  'Implementa��o do campo para automatiza��o do pre�o de custo.
  'Solicitante: Candy Clean
  Dim dblPrecoCustoCalculado As Double
  
  
  On Error GoTo Processa_Erro
  
  
  '11/07/2002 - mpdea
  'Verifica a exist�ncia da movimenta��o e seta o registro de entrada principal
  strSQL = "SELECT * FROM Entradas WHERE Filial = " & Filial & _
           " AND Sequ�ncia = " & Mov
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rsEntradas
    If .BOF And .EOF Then
      'Movimenta��o n�o existe
      Efetiva_Entrada = -99
      .Close
      Set rsEntradas = Nothing
      Exit Function
    End If
    
    dtData = .Fields("Data").Value
    If IsNull(.Fields("Data Emiss�o").Value) Then
      dtDataEmissao = dtData
    Else
      dtDataEmissao = .Fields("Data Emiss�o").Value
    End If
    If IsNull(.Fields("Data Acerto Empr�stimo").Value) Then
      dtDataAcerto = vbNull
    Else
      dtDataAcerto = .Fields("Data Acerto Empr�stimo").Value
    End If
    nCodDigitador = .Fields("Digitador").Value
    nCodFornecedor = .Fields("Fornecedor").Value
    
    nCxDinheiro = .Fields("Dinheiro Caixa").Value
    
    '08/01/2010 - Andrea
    nCxDinheiro = nCxDinheiro + (.Fields("Troco").Value * -1)
    
    '12/01/2010 - Andrea
    'O valor pago em cheques (.Fields("Cheque Caixa").Value), ser�
    'de cheques pr�-datados, pq os cheques a vista nao aparecem no grid
    'de cheques por j� estarem processados.
    nCxChequePre = .Fields("Cheque Caixa").Value
    'nCxCheque = .Fields("Cheque Caixa").Value
    nCxCheque = 0
        
    nConta = .Fields("Conta").Value
    If IsNull(.Fields("Bom Para").Value) Then
      dtBomPara = vbNull
    Else
      dtBomPara = .Fields("Bom Para").Value
    End If
    sDescricao = .Fields("Descri��o").Value & ""
    sNumCheque = .Fields("Num Cheque").Value & ""
    nValCheque = .Fields("Valor Cheque").Value
    sNF = .Fields("Nota Fiscal").Value & ""
    nTotal = .Fields("Total").Value
  End With
  
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  'Implementado verifica��o de venda do tipo WEB - blnWEB_Sale
  'blnWEB_Sale inibe erro ao verificar cliente e digitador
  '---------------------------------------------------------------------------------
  blnWEB_Sale = CLng("0" & rsEntradas.Fields("WebOrderFormID").Value) > 0
  
  
  Rem Encontra a tabela de opera��es
  Set rsOp_Entrada = db.OpenRecordset("Opera��es Entrada", , dbReadOnly)
  rsOp_Entrada.Index = "C�digo"
  rsOp_Entrada.Seek "=", rsEntradas.Fields("Opera��o").Value
  If rsOp_Entrada.NoMatch Then
    Efetiva_Entrada = 1
    rsOp_Entrada.Close
    Set rsOp_Entrada = Nothing
    Exit Function
  End If
  
  Rem Acha Funcion�rios
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios")
  rsFuncionarios.Index = ("C�digo")
  rsFuncionarios.Seek "=", nCodDigitador
  If rsFuncionarios.NoMatch And Not blnWEB_Sale Then  '-> blnWEB_Sale inibe erro
    Efetiva_Entrada = 2
    rsFuncionarios.Close
    Set rsFuncionarios = Nothing
    Exit Function
  End If
  
  Rem Encontra Fornecedor
  Set rsCliFor = db.OpenRecordset("Cli_For")
  rsCliFor.Index = "C�digo"
  rsCliFor.Seek "=", nCodFornecedor
  If rsCliFor.NoMatch And Not blnWEB_Sale Then  '-> blnWEB_Sale inibe erro
    Efetiva_Entrada = 3
    rsCliFor.Close
    Set rsCliFor = Nothing
    Exit Function
  End If
  
  If rsOp_Entrada("Tipo") = "C" Then
    rsCliFor.Edit
    rsCliFor("�ltima Compra") = dtData
    rsCliFor("Data Altera��o") = Format(Date, "dd/mm/yyyy")
    rsCliFor.Update
  End If
  
  Set rsContas_Pagar = db.OpenRecordset("Contas a Pagar")
  Set rsProdutos = db.OpenRecordset("Produtos")
  Set rsResumo_Di�rio = db.OpenRecordset("Resumo Di�rio")
  Set rsEtiquetas = db.OpenRecordset("Etiquetas")
  Set rsCaixa = db.OpenRecordset("Caixa")
  Set rsRes_Financeiro = db.OpenRecordset("Resumo Di�rio Financeiro")
'  Set rsEstoque = db.OpenRecordset("Estoque")
  Set rsPre�os = db.OpenRecordset("Pre�os")
  Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
  Set rsEmprestimos = db.OpenRecordset("Consigna��o Entrada")
  Set rsMovi_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsLan�amentos = db.OpenRecordset("Lan�amentos Banc�rios")
  Set rsComiss�es = db.OpenRecordset("Comiss�o")
'  Set rsTabelas = db.OpenRecordset("Tabela de Pre�os")
  
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
    
    If Caixa_Novo = True Then 'Come�a o Caixa do dia
      Erro = False
      rsCaixa.Seek "<", Filial, nCaixa, dtData, 0
      If rsCaixa.NoMatch Then Erro = True
      If Not Erro Then If rsCaixa("Filial") <> Filial Then Erro = True
      If Not Erro Then If rsCaixa("Caixa") <> nCaixa Then Erro = True
      If Erro = True Then  'N�o existe dia anterior
        rsCaixa.AddNew
        rsCaixa("Filial") = Filial
        rsCaixa("Data") = dtData
        rsCaixa("Hora") = Format(Time, "hh:mm:ss")
        rsCaixa("Funcion�rio") = nCodDigitador
        rsCaixa("Caixa") = nCaixa
        
        Ordem = 1
        rsCaixa("Ordem") = Ordem
        rsCaixa("Saldo Anterior") = 0
        rsCaixa("Final") = 0
        rsCaixa("Descri��o") = "In�cio do dia"
        rsCaixa.Update
      Else
        Ordem = 1
        Saldo_Ant = rsCaixa("Final")
        Tot_Dinheiro = rsCaixa("Total Dinheiro")
        Tot_Cheques = rsCaixa("Total Cheques")
        Tot_Cheques_Pre = rsCaixa("Total Cheques Pr�")
        Tot_Cart�es = rsCaixa("Total Cart�es")
        Tot_Vales = rsCaixa("Total Vales")
        Tot_Parcela = rsCaixa("Total Parcelamento")
                          
        rsCaixa.AddNew
        rsCaixa("Filial") = Filial
        rsCaixa("Data") = dtData
        rsCaixa("Hora") = Format(Time, "hh:mm:ss")
        rsCaixa("Funcion�rio") = nCodDigitador
        rsCaixa("Caixa") = nCaixa
        rsCaixa("Ordem") = Ordem
        rsCaixa("Descri��o") = "In�cio do dia"
        rsCaixa("Saldo Anterior") = Saldo_Ant
        rsCaixa("Dinheiro") = Tot_Dinheiro
        rsCaixa("Cheques") = Tot_Cheques
        rsCaixa("Cheques Pr�") = Tot_Cheques_Pre
        rsCaixa("Cart�es") = Tot_Cart�es
        rsCaixa("Vales") = Tot_Vales
        rsCaixa("Total Dinheiro") = Tot_Dinheiro
        rsCaixa("Total Cheques") = Tot_Cheques
        rsCaixa("Total Cheques Pr�") = Tot_Cheques_Pre
        rsCaixa("Total Cart�es") = Tot_Cart�es
        rsCaixa("Total Vales") = Tot_Vales
        rsCaixa("Parcelamento") = Tot_Parcela
        rsCaixa("Total Parcelamento") = Tot_Parcela
        rsCaixa("Final") = Saldo_Ant
        rsCaixa.Update
      End If
       
      rsCaixa.Seek "<", Filial, nCaixa, dtData, 9999
    End If
  
     
     Rem Neste ponto tem o �ltimo caixa no buffer
     Rem Acha parcela a vista
     Ordem = rsCaixa("Ordem")
     Ordem = Ordem + 1
     Saldo_Ant = rsCaixa("Final")
     Tot_Dinheiro = rsCaixa("Total Dinheiro")
     Tot_Cheques = rsCaixa("Total Cheques")
     Tot_Cheques_Pre = rsCaixa("Total Cheques Pr�")
     Tot_Cart�es = rsCaixa("Total Cart�es")
     Tot_Vales = rsCaixa("Total Vales")
     Tot_Parcela = rsCaixa("Total Parcelamento")
  
      rsCaixa.AddNew
        rsCaixa("Filial") = Filial
        rsCaixa("Data") = dtData
        rsCaixa("Hora") = Format(Time, "hh:mm:ss")
        rsCaixa("Funcion�rio") = nCodDigitador
        rsCaixa("Caixa") = nCaixa
        rsCaixa("Ordem") = Ordem
        rsCaixa("Descri��o") = "Entrada n�mero " & str(Mov)
        rsCaixa("Saldo Anterior") = Saldo_Ant
        rsCaixa("Total Cheques Pr�") = Tot_Cheques_Pre - nCxChequePre
        rsCaixa("Total Cart�es") = Tot_Cart�es
        rsCaixa("Total Vales") = Tot_Vales
        rsCaixa("Cheques") = -(nCxCheque)
        rsCaixa("Cheques Pr�") = -(nCxChequePre)
        rsCaixa("Total Cheques") = Tot_Cheques - nCxCheque '- nCxChequePre
        rsCaixa("Dinheiro") = -(nCxDinheiro)
        rsCaixa("Total Dinheiro") = Tot_Dinheiro - nCxDinheiro
        rsCaixa("Total Parcelamento") = Tot_Parcela
        rsCaixa("Final") = Tot_Dinheiro - nCxCheque - nCxChequePre - nCxDinheiro + Tot_Cheques + Tot_Cheques_Pre + Tot_Cart�es + Tot_Vales
      rsCaixa.Update
  End If



  Rem Faz Lancamento na conta banc�ria, se for o caso
  If nValCheque <> 0 Then
    Rem Acha Saldo Anterior
    Saldo_Conta = 0
    rsLan�amentos.Index = "Conta"
    rsLan�amentos.Seek "<", nConta, dtBomPara, 99999999#
    If Not rsLan�amentos.NoMatch Then
      If rsLan�amentos("Conta") = nConta Then
        Saldo_Conta = rsLan�amentos("Saldo Atual")
      End If
    End If
    
    rsLan�amentos.AddNew
    rsLan�amentos("Conta") = nConta
    rsLan�amentos("Data") = dtBomPara
    rsLan�amentos("Descri��o") = sDescricao
    rsLan�amentos("Cheque") = sNumCheque
    rsLan�amentos("D�bito") = nValCheque
    rsLan�amentos("Saldo Anterior") = Saldo_Conta
    rsLan�amentos("Saldo Atual") = Saldo_Conta - nValCheque
    rsLan�amentos.Update
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
    If Erro = False Then If rsMovi_Cheques("Sequ�ncia") <> Mov Then Erro = True

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
    If Erro = False Then If rsMovi_Parcelas("Sequ�ncia") <> Mov Then Erro = True
    
    If Erro = False Then
      Ordem = rsMovi_Parcelas("Ordem")
      If rsMovi_Parcelas("Bom") >= dtData Then
        rsContas_Pagar.AddNew
        rsContas_Pagar("Filial") = Filial
        rsContas_Pagar("Fornecedor") = nCodFornecedor
        rsContas_Pagar("Data Emiss�o") = dtDataEmissao
        rsContas_Pagar("Descri��o") = "Parcela " & str(Aux_Int)
        rsContas_Pagar("Vencimento") = rsMovi_Parcelas("Bom")
        rsContas_Pagar("Valor") = rsMovi_Parcelas("Valor")
        rsContas_Pagar("Sequ�ncia") = Mov
        rsContas_Pagar("Nota") = sNF
        '30/09/2004 - Daniel
        'Tratamento para Consigna��es da Resultado
        If frmImpressaoNFPrestacao.gbConsignacaoResultado Then
          rsContas_Pagar("Centro de Custo") = IIf(IsNumeric(Trim(frmImpressaoNFPrestacao.cboCodigoCC.Text)), Trim(frmImpressaoNFPrestacao.cboCodigoCC.Text), 1)
        Else
          rsContas_Pagar("Centro de Custo") = IIf(IsNumeric(Trim(frmEntrada.cboCodigoCC.Text)), Trim(frmEntrada.cboCodigoCC.Text), 1)
        End If
        rsContas_Pagar("Data Altera��o") = Format(Date, "dd/mm/yyyy")
        rsContas_Pagar.Update
        Aux_Int = Aux_Int + 1
      End If
    End If
  Loop Until Erro = True
 
  Rem Atualiza Resumo Di�rio
  If rsOp_Entrada("Tipo") <> "P" Then
    rsResumo_Di�rio.Index = "Data"
    rsResumo_Di�rio.Seek "=", Filial, dtData
    If rsResumo_Di�rio.NoMatch Then
      rsResumo_Di�rio.AddNew
      rsResumo_Di�rio("Filial") = Filial
      rsResumo_Di�rio("Data") = dtData
    Else
      rsResumo_Di�rio.Edit
    End If
    If rsOp_Entrada("Tipo") = "C" Then rsResumo_Di�rio("Valor Compras") = CDbl(rsResumo_Di�rio("Valor Compras")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "T" Then rsResumo_Di�rio("Valor T Entrada") = CDbl(rsResumo_Di�rio("Valor T Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "A" Then rsResumo_Di�rio("Valor A Entrada") = CDbl(rsResumo_Di�rio("Valor A Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "G" Then rsResumo_Di�rio("Valor G Entrada") = CDbl(rsResumo_Di�rio("Valor G Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "E" Then rsResumo_Di�rio("Valor E Entrada") = CDbl(rsResumo_Di�rio("Valor E Entrada")) + CDbl(nTotal)
    If rsOp_Entrada("Tipo") = "D" Then
      rsResumo_Di�rio("Valor Devolu��o") = CDbl(rsResumo_Di�rio("Valor Devolu��o")) + CDbl(nTotal)
      '08/08/2003 - maikel
      '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolu��o) do quick store
      'rsResumo_Di�rio("Valor Vendas") = CDbl(rsResumo_Di�rio("Valor Vendas")) - CDbl(nTotal)
    End If
    rsResumo_Di�rio.Update
  End If
  
 
  Rem Atualiza Resumo Di�rio Financeiro
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
      rsRes_Financeiro("Valor Devolu��o") = CDbl(rsRes_Financeiro("Valor Devolu��o")) + CDbl(nTotal)
      '08/08/2003 - maikel
      '             Comentada a linha abaixo para resolver o problema de estoque (referente a devolu��o) do quick store
      'rsRes_Financeiro("Valor Vendas") = CDbl(rsRes_Financeiro("Valor Vendas")) - CDbl(nTotal)
    End If
    rsRes_Financeiro.Update
  End If
  

  '--------------------------------------------------------------------------------
  '14/06/2006 - mpdea
  'Percentual de Frete para c�lculos de custo
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


  rsProdutos.Index = "C�digo"
  
  strSQL = "SELECT * FROM [Entradas - Produtos] WHERE Filial = " & Filial & _
           " AND Sequ�ncia = " & Mov
  Set rsEntraProd = db.OpenRecordset(strSQL, dbOpenSnapshot)
  If rsEntraProd.BOF And rsEntraProd.EOF Then
    'Movimenta��o de produtos n�o existe
    Efetiva_Entrada = -98
    rsEntraProd.Close
    Set rsEntraProd = Nothing
    Exit Function
  End If
  
  Do Until rsEntraProd.EOF
    
    blnWEBSynchronize = False
    
    sCodProd = rsEntraProd.Fields("C�digo").Value
    If Len(sCodProd) > 0 Then
      
      With rsEntraProd
        nQtde = .Fields("Qtde").Value
        nPrecoFinal = .Fields("Pre�o Final").Value
        nPreco = .Fields("Pre�o").Value
        bEtiqueta = .Fields("Etiqueta").Value
      End With
      
      Ordem = nRow + 1
      'Verifica se tem grade
      C�d = ""
      Tamanho = 0
      Cor = 0
      Edi��o = 0
      Tipo = 0
      Erro = 0
      
      Call Acha_Produto(sCodProd, C�d, Tamanho, Cor, Edi��o, Tipo, Erro)
      
      If Erro = 0 Then
      
        C�d = UCase(C�d)
        
        rsProdutos.Seek "=", C�d
         
        'Neste ponto C�D tem o c�digo do produto
        'Tamanho e Cor cont�m os respectivos dados
        'Agora grava arquivo do estoque
        
        Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
      
        Rem  Ajusta Estoque
        If rsOp_Entrada("Estoque") = True And Not rsEntraProd("InGeradoViaConsig") Then
        
'-------------------------------------------------------------------------------------
    '04/03/2004 - mpdea
    'Modificado par�metro de abertura do recordset
    'dbOpenSnapshot (muito lento!? 8-|) para dbOpenDynaset com dbReadOnly
    'e modificado para que salve somente no final da atualiza��o
    'de estoque o recordset
    '
    '10/10/2003 - Maikel
    '             Modificada a forma de analisar a tabela de estoque. Da forma antiga gerava erro 3022 ao efetuar movimenta��o com data retroativa.
'''    strSQL = "SELECT * FROM Estoque WHERE "

    strSQL = "SELECT * FROM [Estoque Final] WHERE " & _
             " Filial = " & Filial & _
             " AND Produto = '" & C�d & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edi��o = " & Edi��o    ''' & " ORDER BY Data"
             
'    'LOG ESPECIFICO PARA MARE MANSA
'    Dim sSQL_Aux As String
'    Dim sCod_mare As String
'    Dim iTam_mare As String
'    Dim iCor_mare As String
'
'    If IsNull(C�d) Then
'      sCod_mare = "N"
'    Else
'      sCod_mare = C�d
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
'    sSQL_Aux = sSQL_Aux & Left("1ENTRADA - " & Filial & " : " & sCod_mare & " : " & iTam_mare & " : " & iCor_mare & " : " & Edi��o, 80) & "', 'VENDENDO MARE')"
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
             " AND Produto = '" & C�d & "'" & _
             " AND Tamanho = " & Tamanho & _
             " AND Cor = " & Cor & _
             " AND Edi��o = " & Edi��o & _
             " AND Data = #" & Format(Now, "mm/dd/yyyy") & "#"
             '''" AND Data = #" & Format(rsEntradas("Data"), "mm/dd/yyyy") & "#"
            
            
            
'    'LOG ESPECIFICO PARA MARE MANSA
'    If IsNull(C�d) Then
'      sCod_mare = "N"
'    Else
'      sCod_mare = C�d
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
'    sSQL_Aux = sSQL_Aux & Left("3ENTRADA - " & Filial & " : " & sCod_mare & " : " & iTam_mare & " : " & iCor_mare & " : " & Edi��o & " : " & Format(rsEntradas("Data"), "mm/dd/yyyy"), 80) & "', 'VENDENDO MARE')"
'    db.Execute sSQL_Aux, dbFailOnError
'    'fim
    
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .Edit
        
'        'LOG ESPECIFICO PARA MARE MANSA
'        If IsNull(C�d) Then
'            sCod_mare = "N"
'        Else
'            sCod_mare = C�d
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
        .Fields("Produto").Value = C�d
        .Fields("Tamanho").Value = Tamanho
        .Fields("Cor").Value = Cor
        .Fields("Edi��o").Value = Edi��o
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

'          Rem Encontra a posi��o do estoque
'          Criar_Registro = False
'          Estoque_Final = 0
'          rsEstoque.Index = "Produto"
'          rsEstoque.Seek "=", Filial, dtData, C�d, Tamanho, Cor, Edi��o
'
'          If Not rsEstoque.NoMatch Then
'            Estoque_Final = rsEstoque("Estoque Final")
'            Estoque2 = rsEstoque("Estoque Final")
'          End If
'          If rsEstoque.NoMatch Then
'            rsEstoque.Index = "Data"
'            rsEstoque.Seek "<", Filial, C�d, Tamanho, Cor, Edi��o, dtData
'            If rsEstoque.NoMatch Then Criar_Registro = True
'            If Not rsEstoque.NoMatch Then
'              If rsEstoque("Filial") = Filial And rsEstoque("Produto") = C�d And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edi��o") = Edi��o Then
'                Criar_Registro = True
'                Estoque_Final = rsEstoque("Estoque Final")
'                Estoque2 = Estoque_Final
'              End If
'            End If
'
'            rsEstoque.AddNew
'            rsEstoque("Filial") = Filial
'            rsEstoque("Data") = dtData
'            rsEstoque("Produto") = C�d
'            rsEstoque("Tamanho") = Tamanho
'            rsEstoque("Cor") = Cor
'            rsEstoque("Edi��o") = Edi��o
'            rsEstoque("Classe") = rsProdutos("Classe")
'            rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
'            rsEstoque("Estoque Anterior") = Estoque_Final
'            rsEstoque.Update
'
'            rsEstoque.Index = "Produto"
'            rsEstoque.Seek "=", Filial, dtData, C�d, Tamanho, Cor, Edi��o
'          End If
      
'-------------------------------------------------------------------------------------
      
          Rem neste ponto esta com o registro de estoque
          Rem no buffer, agora soma com os valores da movimenta��o
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
              rsEstoque("Gr�tis Entra") = rsEstoque("Gr�tis Entra") + nQtde
              rsEstoque("Valor Gr�tis Entra") = Format(rsEstoque("Valor Gr�tis Entra") + nPrecoFinal, "############0.00")
            Case "E"
              rsEstoque("Empre Entra") = rsEstoque("Empre Entra") + nQtde
              rsEstoque("Valor Empre Entra") = Format(rsEstoque("Valor Empre Entra") + nPrecoFinal, "############0.00")
            Case "D"
              rsEstoque("Devolu��o") = rsEstoque("Devolu��o") + nQtde
              rsEstoque("Valor Devolu��o") = Format(rsEstoque("Valor Devolu��o") + nPrecoFinal, "############0.00")
              
              '08/08/2003 - maikel
              '             Comentadas as duas linhas abaixo para resolver o problema de estoque (referente a devolu��o) do quick store
'              rsEstoque("Vendas") = rsEstoque("Vendas") - nQtde
'              rsEstoque("Valor Vendas") = Format(rsEstoque("Valor Vendas") - nPrecoFinal, "############0.00")
          End Select
      
          Estoque2 = rsEstoque("Estoque Anterior")
          Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
          Estoque_Final = Estoque_Final - rsEstoque("Transf Sa�da") + rsEstoque("Transf Entra")
          Estoque_Final = Estoque_Final - rsEstoque("Ajuste Sa�da") + rsEstoque("Ajuste Entra")
          Estoque_Final = Estoque_Final - rsEstoque("Gr�tis Sa�da") + rsEstoque("Gr�tis Entra")
          Estoque_Final = Estoque_Final - rsEstoque("Empre Sa�da") + rsEstoque("Empre Entra")
          
          '08/08/2003 - maikel
          '             Descomentada a soma da coluna desconto para resolver o problema de estoque
          Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolu��o")
    
          If rsProdutos("Estoque") = False Then
            Estoque_Final = 0
          End If
    
          rsEstoque("Estoque Final") = Estoque_Final
          rsEstoque.Update
          
          Rem Arruma Estoque Final
          '''Grava_Estoque_Final gnCodFilial, C�d, Tamanho, Cor, Edi��o, CSng(Estoque_Final), dtData
          Grava_Estoque_Final gnCodFilial, C�d, Tamanho, Cor, Edi��o, CSng(Estoque_Final), Format(Now, "dd/mm/yyyy")
          
        End If
        
        '---------------------------------------------------------------------------
        '17/05/2005 - mpdea
        'C�lculo do custo com aplica��o do valor de ICMS Retido
        Dim p_tmp_ValorIcmsRetido As Double
        If IsNumeric(rsEntraProd.Fields("ValorIcmsRetido").Value) Then
          p_tmp_ValorIcmsRetido = CDbl(rsEntraProd.Fields("ValorIcmsRetido").Value)
        Else
          p_tmp_ValorIcmsRetido = 0
        End If
        
        dblPrecoFinalCusto = CDbl(nPrecoFinal) + p_tmp_ValorIcmsRetido
        
        '14/06/2006 - mpdea
        'Aplica frete no custo se o percentual calculado for maior do que zero
        'e n�o houver valor de icms retido, pois o frete j� foi adicionado anteriormente
        If sngPercFrete > 0 And p_tmp_ValorIcmsRetido = 0 Then
          dblPrecoFinalCusto = dblPrecoFinalCusto * (1 + sngPercFrete)
        End If
        
        
        'Calcula Custo M�dio e Grava
        If rsOp_Entrada("Tipo") = "C" Then
          
          Custo_M�dio = IIf(IsNull(rsProdutos("Custo M�dio")), 0, rsProdutos("Custo M�dio")) * Estoque2
          
          '26/08/2003 - mpdea
          'Calcula com o pre�o final (IPI, desconto)
          Custo_M�dio = Custo_M�dio + dblPrecoFinalCusto 'nQtde * nPreco
          
          If (Estoque2 + nQtde) <> 0 Then
            Custo_M�dio = Custo_M�dio / (Estoque2 + nQtde)
          End If
          If (Estoque2 + nQtde) = 0 Then
            Custo_M�dio = nPreco
          End If
          If Estoque2 < 0 Then
            Custo_M�dio = nPreco
          End If
          
          With rsProdutos
            .Edit
            .Fields("�ltima Compra").Value = Format$(dtData, "dd/mm/yyyy")
            .Fields("�ltimo Custo").Value = Format(dblPrecoFinalCusto / CDbl(nQtde), FORMAT_VALUE)
            .Fields("Custo M�dio").Value = Format(Custo_M�dio, FORMAT_VALUE)
            .Fields("�ltimo Fornecedor").Value = nCodFornecedor
            '16/11/2004 - Daniel
            'Adicionado tratamento caso a moeda
            'seja nula do produto, colocaremos para
            'igual a 1 (Real)
            'Case: Nazareno, n�o foi identificado como mas
            'alguns produtos estavam com moeda = 0
            If .Fields("Moeda").Value = 0 Then .Fields("Moeda").Value = 1
            
            .Update
          End With
          blnWEBSynchronize = True
        
        End If
        
        '22/09/2005 - mpdea
        'Grava Custo para Pre�o de lista sem IPI
        'Utilizado na pasta C�lculos do Produto
        If rsOp_Entrada.Fields("GravaCustoPrecoListaSemIPI").Value Then
          With rsProdutos
            .Edit
            '.Fields("Custo Pre�o Valor").Value = Format(nPrecoFinal / CDbl(nQtde), FORMAT_VALUE)
            .Fields("Custo Pre�o Valor").Value = Format(dblPrecoFinalCusto / CDbl(nQtde), FORMAT_VALUE)
            .Update
          End With
          blnWEBSynchronize = True
        End If
        
        '30/08/2007 - Anderson
        'Implementa��o do campo para automatiza��o do pre�o de custo.
        'Solicitante: Candy Clean
        If rsOp_Entrada.Fields("PrecoCustoCalculado").Value Then
          Calcula_Custo dblPrecoCustoCalculado, rsProdutos("Custo Desconto Fixo").Value, rsProdutos("Custo Desconto Valor").Value, rsProdutos("Custo Desconto Perc").Value, rsProdutos("Custo Pre�o Valor").Value, rsProdutos("Custo Frete Fixo").Value, rsProdutos("Custo Frete Valor").Value, rsProdutos("Custo Frete Perc").Value, rsProdutos("Custo ICM Compra Fixo").Value, rsProdutos("Custo ICM Compra Valor").Value, rsProdutos("Custo ICM Compra Perc").Value, rsProdutos("Custo IPI Compra Fixo").Value, rsProdutos("Custo IPI Compra Valor").Value, rsProdutos("Custo IPI Compra Perc").Value, rsProdutos("Custo Custo Finan Fixo").Value, rsProdutos("Custo Custo Finan Valor").Value, rsProdutos("Custo Custo Finan Perc").Value, rsProdutos("Custo Outros Compra Fixo").Value, rsProdutos("Custo Outros Compra Valor").Value, rsProdutos("Custo Outros Compra Perc").Value
        End If
        '---------------------------------------------------------------------------
        
      
        Rem Diminui comiss�o se for devolu��o
        If rsOp_Entrada("Tipo") = "D" Then
          If rsOp_Entrada("Comiss�o") = True Then
            Comiss�o = rsFuncionarios("Comiss�o")
            If rsProdutos("Comiss�o Sobrep�e") = True Then
              Comiss�o = rsProdutos("Comiss�o")
            End If
'              Comiss�o = Comiss�o * rsTabelas("Multiplicador Comiss�o")
            Comiss�o = Format(Comiss�o, "#############0.00")
            rsComiss�es.AddNew
            rsComiss�es("Data") = dtData
            rsComiss�es("Vendedor") = nCodDigitador
            rsComiss�es("Produto") = C�d
            rsComiss�es("Tamanho") = Tamanho
            rsComiss�es("Cor") = Cor
            '31/03/2005 - Daniel
            'Antiga linha abaixo, estava sem o sinal de menos (-)
            'rsComiss�es("Qtde") = nQtde
            'A partir da 6.52.0.28 contemplou esta altera��o
            rsComiss�es("Qtde") = -nQtde
            rsComiss�es("Valor") = -nPrecoFinal
            rsComiss�es("Sequ�ncia") = Mov
            '14/02/2005 - Daniel
            '
            'Solicitante: Daring - RJ
            '
            'Se ocorre devolu��o e esta devolu��o implica em abatimento de
            'comiss�o do vendedor, o Quick estava descontando erroneamente
            'da comiss�o para casos em que a venda possuia descontos.
            If (Len(frmEntrada.gsTabelaVenda) & "") > 0 Then 'Foi preenchida a var global...
              Dim dblValorDoCadastroProduto As Double
              Dim strCodProdSemGrade        As String
              Dim rstTabelaPrecos           As Recordset
              Dim sngPercentComisDesconto   As Single
              Dim dblValorComissao          As Double
              
              Set rstTabelaPrecos = db.OpenRecordset("SELECT PercentualComissaoDesconto FROM [Tabela de Pre�os] WHERE Tabela = '" & frmEntrada.gsTabelaVenda & "'", dbOpenSnapshot)
              
              If rstTabelaPrecos.RecordCount > 0 Then
                
                Call ReduzirComissao(frmEntrada.gsTabelaVenda & "", C�d & "", dblValorDoCadastroProduto)
              
                If Not IsNull(rstTabelaPrecos.Fields("PercentualComissaoDesconto")) Then
                  sngPercentComisDesconto = rstTabelaPrecos.Fields("PercentualComissaoDesconto")
                Else
                  sngPercentComisDesconto = 0
                End If
              
                'Se for diferente ocorre o abatimento pela metade ou percentual da comiss�o
                If dblValorDoCadastroProduto <> Format((nPrecoFinal / nQtde), FORMAT_VALUE) Then
                  
                  dblValorComissao = (rsComiss�es("Valor") * Comiss�o / 100)
                  dblValorComissao = dblValorComissao * ((100 - sngPercentComisDesconto) / 100)
                  
                  rsComiss�es("Comiss�o") = Truncate(dblValorComissao, 6)
                  
                Else
                  rsComiss�es("Comiss�o") = CCur(Format((rsComiss�es("Valor") * Comiss�o / 100), "###########0.00"))
                End If
                
              Else
                rsComiss�es("Comiss�o") = CCur(Format((rsComiss�es("Valor") * Comiss�o / 100), "###########0.00"))
              End If
            
              rstTabelaPrecos.Close
              Set rstTabelaPrecos = Nothing
              
            Else
              rsComiss�es("Comiss�o") = CCur(Format((rsComiss�es("Valor") * Comiss�o / 100), "###########0.00"))
            End If
            rsComiss�es("Filial") = gnCodFilial
            rsComiss�es("Cliente") = nCodFornecedor
            rsComiss�es.Update
          End If
        End If
        
      
        Rem Grava etiquetas
        If bEtiqueta = True Then
          rsEtiquetas.Index = "Funcion�rio"
          rsEtiquetas.Seek "=", nCodDigitador, C�d, Tamanho, Cor
          If rsEtiquetas.NoMatch Then
             rsEtiquetas.AddNew
          Else
             rsEtiquetas.Edit
          End If
          rsEtiquetas("Funcion�rio") = nCodDigitador
          rsEtiquetas("Produto") = C�d
          rsEtiquetas("Tamanho") = Tamanho
          rsEtiquetas("Cor") = Cor
          rsEtiquetas("Qtde") = rsEtiquetas("Qtde") + nQtde
          rsEtiquetas("Sequ�ncia") = Mov
          rsEtiquetas.Update
        End If
      
      
        'Atualiza pre�o de custo na tabela  CUSTO
        'quando for Compra
        rsPre�os.Index = "Tabela"
        If rsOp_Entrada("Gravar Custo") = True Then
          rsPre�os.Seek "=", "CUSTO", rsProdutos("C�digo")
          If rsPre�os.NoMatch Then
             rsPre�os.AddNew
             rsPre�os("Tabela") = "CUSTO"
             rsPre�os("Produto") = rsProdutos("C�digo")
          Else
             rsPre�os.Edit
          End If
      
          '03/11/2004 - Daniel
          'Tratamento para quando o produto tiver
          'pre�o em D�lar
          If MoedaReal(rsProdutos("Moeda").Value) Then
            rsPre�os("Pre�o") = CSng(dblPrecoFinalCusto) / nQtde
            
            '30/08/2007 - Anderson
            'Implementa��o do campo para automatiza��o do pre�o de custo.
            'Solicitante: Candy Clean
            If rsOp_Entrada.Fields("PrecoCustoCalculado").Value Then
              rsPre�os("Pre�o") = dblPrecoCustoCalculado
            End If

            rsPre�os.Update
          Else 'Produto em D�lar ou em outra moeda
            Dim dblCotacao As Double
            
            Call BuscarUltimaCotacao(rsProdutos("Moeda").Value, dblCotacao)
            
            '07/06/2005 - Daniel
            'Corre��o do bug: Imposs�vel divis�o por zero
            'Adicionado tratamento para evitar erro na divis�o por zero
            'Este erro aparecia quando n�o havia cota��o alguma cadastrada
            If dblCotacao > 0 Then
              rsPre�os("Pre�o") = Format(CSng((CSng(dblPrecoFinalCusto) / nQtde / dblCotacao)), FORMAT_VALUE)
              
              '30/08/2007 - Anderson
              'Implementa��o do campo para automatiza��o do pre�o de custo.
              'Solicitante: Candy Clean
              If rsOp_Entrada.Fields("PrecoCustoCalculado").Value Then
                rsPre�os("Pre�o") = Format(CSng((CSng(dblPrecoCustoCalculado) / dblCotacao)), FORMAT_VALUE)
              End If

            End If
            
            rsPre�os.Update
            
          End If
          
          '-------------------------------------------------------------------------
          '05/08/2003 - mpdea
          'Grava informa��es de custo para opera��es diferentes de Compra
          'Obs.: Opera��o de Compra j� possui tratamento pr�prio
          If rsOp_Entrada.Fields("Tipo").Value <> "C" Then
            
            Custo_M�dio = IIf(IsNull(rsProdutos("Custo M�dio")), 0, rsProdutos("Custo M�dio")) * Estoque2
            
            '26/08/2003 - mpdea
            'Calcula com o pre�o final (IPI, desconto)
            Custo_M�dio = Custo_M�dio + nPrecoFinal 'nQtde * nPreco
            
            If (Estoque2 + nQtde) <> 0 Then
              Custo_M�dio = Custo_M�dio / (Estoque2 + nQtde)
            End If
            If (Estoque2 + nQtde) = 0 Then
              Custo_M�dio = nPreco
            End If
            If Estoque2 < 0 Then
              Custo_M�dio = nPreco
            End If
            
            With rsProdutos
              .Edit
              .Fields("�ltimo Custo").Value = Format(nPrecoFinal / CDbl(nQtde), FORMAT_VALUE)
              .Fields("Custo M�dio").Value = Format(Custo_M�dio, FORMAT_VALUE)
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
          rsResumo_Clientes.Seek "=", nCodFornecedor, dtData, C�d, Tamanho, Cor, Edi��o, Mov
          If rsResumo_Clientes.NoMatch Then
             rsResumo_Clientes.AddNew
          Else
             rsResumo_Clientes.Edit
          End If
           
          rsResumo_Clientes("Dia") = dtData
          rsResumo_Clientes("Cliente") = nCodFornecedor
          rsResumo_Clientes("Produto") = C�d
          rsResumo_Clientes("Tamanho") = Tamanho
          rsResumo_Clientes("Cor") = Cor
          rsResumo_Clientes("Edi��o") = Edi��o
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
          
          rsResumo_Clientes("Sequ�ncia") = Mov
          rsResumo_Clientes("Filial") = gnCodFilial
  
          rsResumo_Clientes.Update
          
        End If
      
      
      
        Rem Atualiza arquivo de Empr�stimos
        If rsOp_Entrada("Tipo") = "E" And Not rsEntraProd("InGeradoViaConsig") Then
        
          Rem Saldo Emprestado = 0 para este empr�stimo
          rsEmprestimos.Index = "Cliente"
          Saldo_Emp = 0
          Ordem_Emp = Ordem
          Emp_Existe = False
                   
          rsEmprestimos.Seek "<", gnCodFilial, Mov, nCodFornecedor, C�d, Tamanho, Cor, Edi��o, 999999
          If Not rsEmprestimos.NoMatch Then
            If rsEmprestimos("Filial") = gnCodFilial Then
              If rsEmprestimos("Sequ�ncia") = Mov Then
                If rsEmprestimos("Fornecedor") = nCodFornecedor Then
                  If rsEmprestimos("Produto") = C�d Then
                    If rsEmprestimos("Tamanho") = Tamanho Then
                      If rsEmprestimos("Cor") = Cor Then
                        If rsEmprestimos("Edi��o") = Edi��o Then
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
            rsEmprestimos("Sequ�ncia") = Mov
            rsEmprestimos("Fornecedor") = nCodFornecedor
            rsEmprestimos("Produto") = C�d
            rsEmprestimos("Tamanho") = Tamanho
            rsEmprestimos("Cor") = Cor
            rsEmprestimos("Edi��o") = Edi��o
            rsEmprestimos("Ordem") = Ordem_Emp
          End If
              
          rsEmprestimos("Saldo Anterior") = Saldo_Emp
          rsEmprestimos("Empr�stimo Recebido") = nQtde
          rsEmprestimos("Saldo Atual") = Saldo_Emp + nQtde
          rsEmprestimos("Pre�o Unit�rio") = (nPrecoFinal / nQtde)
          rsEmprestimos("Data Opera��o") = dtData
          rsEmprestimos("Data Altera��o") = Format(Date, "dd/mm/yyyy")
          If dtDataAcerto <> vbNull Then
            rsEmprestimos("Data Cobran�a") = dtDataAcerto
          End If
    
          rsEmprestimos.Update
             
        End If
        
        If blnWEBSynchronize Then
          'Atualiza o sincronismo para o produto WEB alterado
          Call WEB_SynchronizeProduct(rsProdutos("C�digo").Value)
        End If
        
      End If  'Erro = 0
    End If
    
    'Pr�ximo produto
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
  rsResumo_Di�rio.Close
  rsEtiquetas.Close
  rsCaixa.Close
  rsRes_Financeiro.Close
  
  '07/01/2010 - Andrea
  rsContas_Receber.Close
  rsMovi_Cheques.Close
  
  '27/06/2003 - mpdea
  'Verifica se o objeto foi criado antes de fech�-lo
  If Not rsEstoque Is Nothing Then
    rsEstoque.Close
  End If
  
  rsPre�os.Close
  rsResumo_Clientes.Close
  rsEmprestimos.Close
  rsMovi_Parcelas.Close
  rsLan�amentos.Close
  rsComiss�es.Close
  
  Set rsEntradas = Nothing
  Set rsEntraProd = Nothing
  Set rsOp_Entrada = Nothing
  Set rsFuncionarios = Nothing
  Set rsCliFor = Nothing
  Set rsContas_Pagar = Nothing
  Set rsProdutos = Nothing
  Set rsResumo_Di�rio = Nothing
  Set rsEtiquetas = Nothing
  Set rsCaixa = Nothing
  Set rsRes_Financeiro = Nothing
  Set rsEstoque = Nothing
  Set rsPre�os = Nothing
  Set rsResumo_Clientes = Nothing
  Set rsEmprestimos = Nothing
  Set rsMovi_Parcelas = Nothing
  Set rsLan�amentos = Nothing
  Set rsComiss�es = Nothing
   
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
          Efetiva_Entrada = -1 'A��o cancelada
          Exit Function
        End If
      
'        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
'          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Efetiva Entrada") = vbOK Then
'          nRepeatUpdateLocked = 0
'          Resume
'        Else
'          Efetiva_Entrada = -1 'A��o cancelada
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
          Efetiva_Entrada = -1 'A��o cancelada
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select

End Function

Private Function MoedaReal(ByVal Moeda As Byte) As Boolean
  '03/11/2004 - Daniel
  'Verifica��o do tipo de moeda que implicar� no c�lculo
  'da tabela CUSTO
  Dim rstMoedas As Recordset
  Dim strMoeda  As String
  
  Set rstMoedas = db.OpenRecordset("SELECT * FROM Moedas WHERE C�digo = " & Moeda, dbOpenDynaset)
  
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
  
  Set rstCotacoes = db.OpenRecordset("SELECT * FROM Cota��es WHERE Moeda = " & Moeda & " ORDER BY Data ", dbOpenDynaset)

  With rstCotacoes
    If Not (.BOF And .EOF) Then
      .MoveLast
      
      Cotacao = Format(.Fields("Cota��o").Value, FORMAT_VALUE)
      
    End If
    .Close
  End With
  
  Set rstCotacoes = Nothing

End Sub

Private Sub ReduzirComissao(ByVal Tabela As String, ByVal Produto As String, ByRef dblValorDoCadastroProduto As Double)
  '14/02/2005 - Daniel
  'Problema levantado pela Daring
  'Se ocorre devolu��o e esta devolu��o implica em abatimento de
  'comiss�o do vendedor, o Quick estava descontando erroneamente
  'da comiss�o para casos em que a venda possuia descontos.
  Dim rstPrecos As Recordset
  Dim strSQL    As String
  
  strSQL = "SELECT Pre�o FROM Pre�os "
  strSQL = strSQL & " WHERE Produto = '" & Produto & "'"
  strSQL = strSQL & " AND Tabela = '" & Tabela & "'"
  
  Set rstPrecos = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstPrecos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      dblValorDoCadastroProduto = .Fields("Pre�o").Value
    End If
    .Close
  End With
  
  Set rstPrecos = Nothing
  
End Sub
