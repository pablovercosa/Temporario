Attribute VB_Name = "modEfetivaSaida"
Option Explicit

'20/12/2005 - mpdea
'Quick Fiscal
'Variáveis para impressão de dados em cupom de parcelamento
'Bematech
'CASE: Margarete Parizoto ME (QS71277-474)
Public g_str_nome_cliente As String
Public g_lng_nr_sequencia As Long
Private objComissao As clsComissao

Public Type tpPaymentType
  dblDinheiro As Double
  dblCheque As Double
  dblChequePre As Double
  dblCartao As Double
  dblVale As Double
  dblParcelamento As Double
End Type

Public Function Desefetiva_Saída(ByVal Filial As Integer, ByVal Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '07/08/2002 - mpdea
  'Inserido os recordsets que estavam a nível modular sem necessidade,
  'ocupando mais memória
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
'  Dim rsParametros As Recordset
  Dim rsOp_Saída As Recordset
'  Dim rsContas_Receber As Recordset
  Dim rsResumo_Diário As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
'  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
'  Dim rsEstoque_Final As Recordset
'  Dim rsPreços As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
'  Dim rsGrade As Recordset
  Dim rsSaidas As Recordset
  Dim rsSaidas_Prod As Recordset
'  Dim rsSaidas_Serv As Recordset
  Dim rsSaída_Cheques As Recordset
  Dim rsSaída_Parcelas As Recordset
  Dim rsComissões As Recordset
  Dim rsComissões_Serv As Recordset
'  Dim rsFuncionarios As Recordset
'  Dim rsTabelas As Recordset
  Dim rsConta_Cli As Recordset
'  Dim rsCartoes As Recordset
'  Dim rsBancos As Recordset
'  Dim rsEdicoes As Recordset
'  Dim rsServicos As Recordset
    
  '11/12/2009 - Andrea
  Dim rsSaída_Cartoes As Recordset
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
 Dim Aux_Prod As String
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim Edição As Long
 Dim Aux_Tipo As Integer
 Dim Aux_Erro As Integer
 Dim Estoque_Final As Double
 Dim Criar_Registro As Integer
 Dim Mensagem As String
 Dim Saldo_Conta As Double
 
 Dim Aux_Val_Produto As Double
 Dim Aux_Val_Serviço As Double
 
 Dim Val_Cheques As Double
 Dim Val_Cheques_Pré As Double
 Dim Comissão As Double
 
  'Variável de Tratamento de Erro
  Dim intRepeatUpdateLocked As Integer
 
  Dim strSQL As String
 
 On Error GoTo Processa_Erro

  Screen.MousePointer = vbHourglass
  
 Set rsSaidas = db.OpenRecordset("Saídas")
' Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
' Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 Set rsOp_Saída = db.OpenRecordset("Operações Saída", , dbReadOnly)
 Set rsResumo_Diário = db.OpenRecordset("Resumo Diário")
 Set rsEtiquetas = db.OpenRecordset("Etiquetas")
 Set rsCaixa = db.OpenRecordset("Caixa")
 Set rsRes_Financeiro = db.OpenRecordset("Resumo Diário Financeiro")
' Set rsEstoque = db.OpenRecordset("Estoque")
' Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
 Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
 Set rsEmprestimos = db.OpenRecordset("Consignação Saída")
' Set rsGrade = db.OpenRecordset("Códigos da Grade")
 Set rsSaidas_Prod = db.OpenRecordset("Saídas - Produtos", , dbReadOnly)
 Set rsSaída_Cheques = db.OpenRecordset("Movimento - Cheques", , dbReadOnly)
 Set rsSaída_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
 Set rsComissões = db.OpenRecordset("Comissão")
 Set rsComissões_Serv = db.OpenRecordset("Comissão Serviços")
 Set rsConta_Cli = db.OpenRecordset("Conta Cliente")
' Set rsCartoes = db.OpenRecordset("Cartões", , dbReadOnly)
' Set rsBancos = db.OpenRecordset("Bancos", , dbReadOnly)

 '11/12/2009 - Andrea
 Set rsSaída_Cartoes = db.OpenRecordset("Movimento - Cartoes", , dbReadOnly)
  
  Screen.MousePointer = vbDefault
  
 rsSaidas.Index = "Sequência"
 rsSaidas.Seek "=", Filial, Mov
 If rsSaidas.NoMatch Then
   Desefetiva_Saída = 1
   Exit Function
 End If
 
 
 Rem Encontra a tabela de operações
 rsOp_Saída.Index = "Código"
 rsOp_Saída.Seek "=", rsSaidas("Operação")
 If rsOp_Saída.NoMatch Then
    Desefetiva_Saída = 2
    Exit Function
 End If


  Screen.MousePointer = vbHourglass
  
 Rem Atualiza Caixa, se for o caso
 'frmEntradas.Percent.Value = 4
 If rsOp_Saída("Dinheiro") = True Then
 ' If rsSaidas("Recebe - Dinheiro") <> 0 Or rsSaidas("Recebe - Cartão") <> 0 Or rsSaidas("Recebe - Vale") <> 0 Then
    Erro = False
    Caixa_Novo = False
    Ordem = 0
      
    rsCaixa.Index = "Data"
    rsCaixa.Seek "<", Filial, rsSaidas("Caixa"), CDate(rsSaidas("Data")), 9999
    If rsCaixa.NoMatch Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Filial") <> Filial Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Data") <> rsSaidas("Data") Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Caixa") <> rsSaidas("Caixa") Then Caixa_Novo = True
   
    If Caixa_Novo = True Then 'Começa o Caixa do dia
       Desefetiva_Saída = 55
       Exit Function
    End If

    
    Rem Neste ponto tem o último caixa no buffer
    Rem Acha cheques
    Val_Cheques = 0
    rsSaída_Cheques.Index = "Ordem"
    Ordem = 0
    Erro = False
    Do
     rsSaída_Cheques.Seek ">", Filial, Mov, Ordem
     If rsSaída_Cheques.NoMatch Then Erro = True
     If Erro = False Then If rsSaída_Cheques("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSaída_Cheques("Sequência") <> Mov Then Erro = True

     If Erro = False Then
       If rsSaída_Cheques("Bom") = rsSaidas("Data") Then
         Val_Cheques = Val_Cheques + rsSaída_Cheques("Valor")
       End If
       If rsSaída_Cheques("Bom") <> rsSaidas("Data") Then
         Val_Cheques_Pré = Val_Cheques_Pré + rsSaída_Cheques("Valor")
       End If
       Ordem = rsSaída_Cheques("Ordem")
     End If
    Loop Until Erro = True

        
    Ordem = rsCaixa("Ordem")
    Ordem = Ordem + 1
    Saldo_Ant = rsCaixa("Final")
    Tot_Dinheiro = rsCaixa("Total Dinheiro")
    Tot_Cheques = rsCaixa("Total Cheques")
    Tot_Cheques_Pre = rsCaixa("Total Cheques Pré")
    Tot_Cartões = rsCaixa("Total Cartões")
    Tot_Vales = rsCaixa("Total Vales")
    If Not rsCaixa("Total Parcelamento") = Null Then
    Tot_Parcela = rsCaixa("Total Parcelamento")
    Else
    rsCaixa.Edit
    rsCaixa("Total Parcelamento") = 0
    rsCaixa.Update
    End If

     rsCaixa.AddNew
       rsCaixa("Filial") = Filial
       rsCaixa("Data") = rsSaidas("Data")
       rsCaixa("Hora") = Format(Time, "hh:mm:ss")
       rsCaixa("Caixa") = rsSaidas("Caixa")
       rsCaixa("Ordem") = Ordem
       rsCaixa("Descrição") = "Cancelada Saída número " & str(Mov)
       rsCaixa("Saldo Anterior") = Saldo_Ant
       'rsCaixa("Total Cheques Pré") = Tot_Cheques_Pre
       rsCaixa("Cartões") = -rsSaidas("Recebe - Cartão")
       rsCaixa("Total Cartões") = Tot_Cartões - rsSaidas("Recebe - Cartão")
       rsCaixa("Vales") = -rsSaidas("Recebe - Vale")
       rsCaixa("Total Vales") = Tot_Vales - rsSaidas("Recebe - Vale")
       rsCaixa("Cheques") = -Val_Cheques
       rsCaixa("Total Cheques") = Tot_Cheques - Val_Cheques
       rsCaixa("Cheques Pré") = -Val_Cheques_Pré
       rsCaixa("Total Cheques Pré") = Tot_Cheques_Pre - Val_Cheques_Pré
       rsCaixa("Parcelamento") = -rsSaidas("Total Prazo")
       rsCaixa("Total Parcelamento") = Tot_Parcela - rsSaidas("Total Prazo")
       rsCaixa("Dinheiro") = -rsSaidas("Recebe - Dinheiro")
       rsCaixa("Total Dinheiro") = Tot_Dinheiro - rsSaidas("Recebe - Dinheiro")
       rsCaixa("Final") = Tot_Dinheiro - rsSaidas("Recebe - Cartão") - rsSaidas("Recebe - Vale") - rsSaidas("Recebe - Dinheiro") - Val_Cheques - Val_Cheques_Pré + Tot_Cartões + Tot_Vales + Tot_Cheques + Tot_Cheques_Pre
     rsCaixa.Update
 ' End If
 End If
 
 
  '---------------------------------------------------------------------------------
  '20/05/2002 - mpdea
  '
  'Otimizado a exclusão dos registros da tabela de Contas a Receber
  '---------------------------------------------------------------------------------
 
 
' Rem Apagar Lançamentos em Controle de Cheques
' rsContas_Receber.Index = "Contas"
' Erro = False
'Lp1_Cheque1:
' rsContas_Receber.Seek ">", "C", Filial, Mov, 0
' If rsContas_Receber.NoMatch Then Erro = True
' If Erro = False Then If rsContas_Receber("Sequência") <> Mov Then Erro = True
' If Erro = False Then If rsContas_Receber("Filial") <> Filial Then Erro = True
' If Erro = False Then If rsContas_Receber("Tipo") <> "C" Then Erro = True
' If Erro = False Then
'   rsContas_Receber.Delete
'   GoTo Lp1_Cheque1
' End If
'
'
'
' Rem Apagar Lançamentos em controle de cartões, se for o caso
' rsContas_Receber.Index = "Contas"
' Erro = False
' Do While True
'    rsContas_Receber.Seek ">", "O", Filial, Mov, 0
'    If rsContas_Receber.NoMatch Then Erro = True
'    If Erro = False Then If rsContas_Receber("Sequência") <> Mov Then Erro = True
'    If Erro = False Then If rsContas_Receber("Filial") <> Filial Then Erro = True
'    If Erro = False Then If rsContas_Receber("Tipo") <> "O" Then Erro = True
'    If Erro = False Then
'      rsContas_Receber.Delete
'    Else
'      Exit Do
'    End If
' Loop
'
' Rem Apaga contas a receber, se for o caso
' rsContas_Receber.Index = "Contas"
' Erro = False
'Lp1_Receber1:
' rsContas_Receber.Seek ">", "R", Filial, Mov, 0
' If rsContas_Receber.NoMatch Then Erro = True
' If Erro = False Then If rsContas_Receber("Sequência") <> Mov Then Erro = True
' If Erro = False Then If rsContas_Receber("Filial") <> Filial Then Erro = True
' If Erro = False Then If rsContas_Receber("Tipo") <> "R" Then Erro = True
' If Erro = False Then
'   rsContas_Receber.Delete
'   GoTo Lp1_Receber1
' End If
  
  
  db.Execute "DELETE * FROM [Contas a Receber] WHERE Filial = " & Filial & _
    " AND Sequência = " & Mov, dbFailOnError
  '10/09/2007 - Anderson
  'Gera arquivo log do sistema
  If g_bolSystemLog Then
    SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
    "DELETE * FROM [Contas a Receber] WHERE Filial = " & Filial & " AND Sequência = " & Mov, _
    "modEfetivaSaida_Desefetiva_Saída", _
    "Contas a Receber", g_strArquivoSystemLog
  End If
  
  '---------------------------------------------------------------------------------
 
 
 Aux_Val_Produto = CDbl(rsSaidas("Total")) - CDbl(rsSaidas("Serviços"))
 Aux_Val_Serviço = CDbl(rsSaidas("Serviços"))

 Rem Atualiza Resumo Diário
 If rsOp_Saída("Tipo") <> "O" Then
   rsResumo_Diário.Index = "Data"
   rsResumo_Diário.Seek "=", Filial, rsSaidas("Data")
   If rsResumo_Diário.NoMatch Then
     rsResumo_Diário.AddNew
     rsResumo_Diário("Filial") = Filial
     rsResumo_Diário("Data") = rsSaidas("Data")
   Else
     rsResumo_Diário.Edit
   End If
   
   If rsOp_Saída("Tipo") = "V" Then
       rsResumo_Diário("Valor Vendas") = CDbl(rsResumo_Diário("Valor Vendas")) - Aux_Val_Produto
       rsResumo_Diário("Valor Serviços") = CDbl(rsResumo_Diário("Valor Serviços")) - Aux_Val_Serviço
   End If
   
   If rsOp_Saída("Tipo") = "T" Then rsResumo_Diário("Valor T Saída") = CDbl(rsResumo_Diário("Valor T Saída")) - CDbl(rsSaidas("Total"))
   If rsOp_Saída("Tipo") = "A" Then rsResumo_Diário("Valor A Saída") = CDbl(rsResumo_Diário("Valor A Saída")) - CDbl(rsSaidas("Total"))
   If rsOp_Saída("Tipo") = "G" Then rsResumo_Diário("Valor G Saída") = CDbl(rsResumo_Diário("Valor G Saída")) - CDbl(rsSaidas("Total"))
   If rsOp_Saída("Tipo") = "E" Then rsResumo_Diário("Valor E Saída") = CDbl(rsResumo_Diário("Valor E Saída")) - CDbl(rsSaidas("Total"))

  rsResumo_Diário.Update
End If

 
  Rem Atualiza Resumo Diário Financeiro
  If rsOp_Saída("Dinheiro") = True Then
    rsRes_Financeiro.Index = "Data"
    rsRes_Financeiro.Seek "=", Filial, rsSaidas("Data")
    If rsRes_Financeiro.NoMatch Then
       rsRes_Financeiro.AddNew
       rsRes_Financeiro("Filial") = Filial
       rsRes_Financeiro("Data") = rsSaidas("Data")
    Else
       rsRes_Financeiro.Edit
    End If
    
    If rsOp_Saída("Tipo") = "V" Then
       rsRes_Financeiro("Valor Vendas") = CDbl(rsRes_Financeiro("Valor Vendas")) - Aux_Val_Produto
       rsRes_Financeiro("Valor Serviços") = CDbl(rsRes_Financeiro("Valor Serviços")) - Aux_Val_Serviço
    End If
    If rsOp_Saída("Tipo") = "T" Then rsRes_Financeiro("Valor T Saída") = CDbl(rsRes_Financeiro("Valor T Saída")) - CDbl(rsSaidas("Total"))
    If rsOp_Saída("Tipo") = "A" Then rsRes_Financeiro("Valor A Saída") = CDbl(rsRes_Financeiro("Valor A Saída")) - CDbl(rsSaidas("Total"))
    If rsOp_Saída("Tipo") = "G" Then rsRes_Financeiro("Valor G Saída") = CDbl(rsRes_Financeiro("Valor G Saída")) - CDbl(rsSaidas("Total"))
    If rsOp_Saída("Tipo") = "E" Then rsRes_Financeiro("Valor E Saída") = CDbl(rsRes_Financeiro("Valor E Saída")) - CDbl(rsSaidas("Total"))


    rsRes_Financeiro.Update
  End If


  Rem Apaga conta do cliente
  rsConta_Cli.Index = "Sequência"
  Erro = False
Lp1_Conta_Cli1:
  rsConta_Cli.Seek ">", Filial, Mov, 0
  If rsConta_Cli.NoMatch Then Erro = True
  If Erro = False Then If rsConta_Cli("Sequência") <> Mov Then Erro = True
  If Erro = False Then If rsConta_Cli("Filial") <> Filial Then Erro = True
  If Erro = False Then
    rsConta_Cli.Delete
    GoTo Lp1_Conta_Cli1
  End If

  Rem Apaga comissões
  rsComissões.Index = "Sequência"
  Erro = False
Lp1_Comissão1:
  rsComissões.Seek ">", Filial, Mov, 0
  If rsComissões.NoMatch Then Erro = True
  If Erro = False Then If rsComissões("Sequência") <> Mov Then Erro = True
  If Erro = False Then If rsComissões("Filial") <> Filial Then Erro = True
  If Erro = False Then
    rsComissões.Delete
    GoTo Lp1_Comissão1
  End If


  Rem Apaga Resumo Clientes
  rsResumo_Clientes.Index = "Sequência"
  Erro = False
Lp1_ResumoCli1:
  rsResumo_Clientes.Seek ">=", Filial, Mov
  If rsResumo_Clientes.NoMatch Then Erro = True
  If Erro = False Then If rsResumo_Clientes("Sequência") <> Mov Then Erro = True
  If Erro = False Then If rsResumo_Clientes("Filial") <> Filial Then Erro = True
  If Erro = False Then
    rsResumo_Clientes.Delete
    GoTo Lp1_ResumoCli1
  End If






  Rem Loop dos Produtos
  rsSaidas_Prod.Index = "Sequência"
  Ordem = 0
  rsProdutos.Index = "Código"
Prox_Prod:
  rsSaidas_Prod.Seek ">", Filial, Mov, Ordem
  
  If rsSaidas_Prod.NoMatch Then GoTo Ve_Serv
  If rsSaidas_Prod("Filial") <> Filial Then GoTo Ve_Serv
  If rsSaidas_Prod("Sequência") <> Mov Then GoTo Ve_Serv
  
  Ordem = rsSaidas_Prod("Linha")
  'Verifica se tem grade
  Cód = rsSaidas_Prod("Código")
  Tamanho = 0
  Cor = 0
  Aux_Prod = Cód
  
  Acha_Produto Aux_Prod, Cód, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro
  If Aux_Erro <> 0 Then
    GoTo Prox_Prod
  End If
   
  Cód = UCase(Cód)
   
  'Neste ponto CÓD tem o código do produto
  'Tamanho e Cor contém os respectivos dados
  'Agora grava arquivo do estoque
  rsProdutos.Seek "=", Cód
  
  Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))

  Rem  Ajusta Estoque
  If rsOp_Saída("Estoque") = True Then
  
'-------------------------------------------------------------------------------------
    '16/11/2003 - mpdea
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
             " AND Data = #" & Format(rsSaidas("Data"), "mm/dd/yyyy") & "#"
            
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .Edit
      Else
        .AddNew
        .Fields("Filial").Value = Filial
        .Fields("Data").Value = rsSaidas("Data").Value
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


'    Rem Encontra a posição do estoque
'    Criar_Registro = False
'    Estoque_Final = 0
'    rsEstoque.Index = "Produto"
'    rsEstoque.Seek "=", Filial, rsSaidas("Data"), Cód, Tamanho, Cor, Edição
'
'    If Not rsEstoque.NoMatch Then Estoque_Final = rsEstoque("Estoque Final")
'    If rsEstoque.NoMatch Then
'       rsEstoque.Index = "Data"
'       rsEstoque.Seek "<", Filial, Cód, Tamanho, Cor, Edição, rsSaidas("Data")
'       If rsEstoque.NoMatch Then Criar_Registro = True
'       If Not rsEstoque.NoMatch Then
'          If rsEstoque("Filial") = Filial And rsEstoque("Produto") = Cód And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edição") = Edição Then
'             Criar_Registro = True
'             Estoque_Final = rsEstoque("Estoque Final")
'           End If
'       End If
'
'       rsEstoque.AddNew
'       rsEstoque("Filial") = Filial
'       rsEstoque("Data") = rsSaidas("Data")
'       rsEstoque("Produto") = Cód
'       rsEstoque("Tamanho") = Tamanho
'       rsEstoque("Cor") = Cor
'       rsEstoque("Edição") = Edição
'       rsEstoque("Classe") = rsProdutos("Classe")
'       rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
'       rsEstoque("Estoque Anterior") = Estoque_Final
'       rsEstoque.Update
'
'       rsEstoque.Index = "Produto"
'       rsEstoque.Seek "=", Filial, rsSaidas("Data"), Cód, Tamanho, Cor, Edição
'      End If

'-------------------------------------------------------------------------------------

      Rem neste ponto esta com o registro de estoque
      Rem no buffer, agora soma com os valores da movimentação
'      rsEstoque.Edit
      If rsOp_Saída("Tipo") = "V" Then
          rsEstoque("Vendas") = rsEstoque("Vendas") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Vendas") = Format(rsEstoque("Valor Vendas") - rsSaidas_Prod("Preço Final"), "############0.00")
      End If
            
      If rsOp_Saída("Tipo") = "T" Then
          rsEstoque("Transf Saída") = rsEstoque("Transf Saída") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor T Saída") = Format(rsEstoque("Valor T Saída") - rsSaidas_Prod("Preço Final"), "############0.00")
      End If
      
      If rsOp_Saída("Tipo") = "A" Then
          rsEstoque("Ajuste Saída") = rsEstoque("Ajuste Saída") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Ajuste Saída") = Format(rsEstoque("Valor Ajuste Saída") - rsSaidas_Prod("Preço Final"), "############0.00")
      End If
          
      If rsOp_Saída("Tipo") = "G" Then
          rsEstoque("Grátis Saída") = rsEstoque("Grátis Saída") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Grátis Saída") = Format(rsEstoque("Valor Grátis Saída") - rsSaidas_Prod("Preço Final"), "############0.00")
      End If
      
      If rsOp_Saída("Tipo") = "E" Then
          rsEstoque("Empre Saída") = rsEstoque("Empre Saída") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Empre Saída") = Format(rsEstoque("Valor Empre Saída") - rsSaidas_Prod("Preço Final"), "############0.00")
      End If

      Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
      Estoque_Final = Estoque_Final + rsEstoque("Transf Saída") + rsEstoque("Transf Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Ajuste Saída") + rsEstoque("Ajuste Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Grátis Saída") + rsEstoque("Grátis Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Empre Saída") + rsEstoque("Empre Entra")
      
      '08/08/2003 - maikel
      '             Descomentada a soma da coluna Devolução para resolver o problema de estoque
      Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolução")

      If rsProdutos("Estoque") = False Then
        Estoque_Final = 0
      End If

      rsEstoque("Estoque Final") = Estoque_Final
      rsEstoque.Update
      rsEstoque.Close
      
      Grava_Estoque_Final Filial, Cód, Tamanho, Cor, Edição, CSng(Estoque_Final), Date
      
  End If

  

  Rem Apaga etiquetas
  If rsSaidas_Prod("Etiqueta") = True Then
    rsEtiquetas.Index = "Funcionário"
    rsEtiquetas.Seek "=", rsSaidas("Digitador"), Cód, Tamanho, Cor
    If rsEtiquetas.NoMatch Then
       rsEtiquetas.AddNew
    Else
       rsEtiquetas.Edit
    End If
    rsEtiquetas("Funcionário") = rsSaidas("Digitador")
    rsEtiquetas("Produto") = Cód
    rsEtiquetas("Tamanho") = Tamanho
    rsEtiquetas("Cor") = Cor
    rsEtiquetas("Qtde") = rsEtiquetas("Qtde") - rsSaidas_Prod("Qtde")
    rsEtiquetas("Sequência") = Mov
    rsEtiquetas.Update
  End If





  Rem Atualiza arquivo de Empréstimos
  If rsOp_Saída("Tipo") = "E" Then
     rsEmprestimos.Index = "Cliente"
Lp_Emp1:
     rsEmprestimos.Seek ">", rsSaidas("Filial"), rsSaidas("Sequência"), rsSaidas("Cliente"), 0, 0, 0, 0, 0
     If Not rsEmprestimos.NoMatch Then
       If rsEmprestimos("Filial") = rsSaidas("Filial") Then
         If rsEmprestimos("Sequência") = rsSaidas("Sequência") Then
           rsEmprestimos.Delete
           GoTo Lp_Emp1
         End If
       End If
     End If
  End If
     

  GoTo Prox_Prod
  
  
Ve_Serv:
  rsComissões_Serv.Index = "Sequência"
  Ordem = 0
Prox_Serv:
  rsComissões_Serv.Seek ">", Filial, Mov, Ordem
    
  If rsComissões_Serv.NoMatch Then GoTo Fim_Desefetiva
  If rsComissões_Serv("Filial") <> Filial Then GoTo Fim_Desefetiva
  If rsComissões_Serv("Sequência") <> Mov Then GoTo Fim_Desefetiva
  
  rsComissões_Serv.Delete
  
  GoTo Prox_Serv
  
  
Fim_Desefetiva:
   Desefetiva_Saída = 0
  
  '---------------------------------------------------------------------------------
  '20/05/2002 - mpdea
  '
  'Incluído o fechamento dos recordsets abertos e suas desassociações
  '---------------------------------------------------------------------------------
  
  rsSaidas.Close
  rsProdutos.Close
  rsOp_Saída.Close
  rsResumo_Diário.Close
  rsEtiquetas.Close
  rsCaixa.Close
  rsRes_Financeiro.Close
'  If Not rsEstoque Is Nothing Then rsEstoque.Close
  rsResumo_Clientes.Close
  rsEmprestimos.Close
  rsSaidas_Prod.Close
  rsSaída_Cheques.Close
  rsSaída_Parcelas.Close
  rsComissões.Close
  rsComissões_Serv.Close
  rsConta_Cli.Close
  
  '11/12/2009 - Andrea
  rsSaída_Cartoes.Close
   
  Set rsSaidas = Nothing
  Set rsProdutos = Nothing
  Set rsOp_Saída = Nothing
  Set rsResumo_Diário = Nothing
  Set rsEtiquetas = Nothing
  Set rsCaixa = Nothing
  Set rsRes_Financeiro = Nothing
  Set rsEstoque = Nothing
  Set rsResumo_Clientes = Nothing
  Set rsEmprestimos = Nothing
  Set rsSaidas_Prod = Nothing
  Set rsSaída_Cheques = Nothing
  Set rsSaída_Parcelas = Nothing
  Set rsComissões = Nothing
  Set rsComissões_Serv = Nothing
  Set rsConta_Cli = Nothing
  Set rsSaída_Cartoes = Nothing
  '---------------------------------------------------------------------------------
   
   Screen.MousePointer = vbDefault
   Exit Function

Processa_Erro:

  Screen.MousePointer = vbDefault
  Select Case Err.Number
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        Screen.MousePointer = vbHourglass
        Call WaitSeconds(1) 'Aguarda um segundo
        Resume
      Else
        
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Screen.MousePointer = vbHourglass
          Resume
        Else
          Desefetiva_Saída = -1 'Ação cancelada
          Exit Function
        End If
        
'        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Desefetiva Saída") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        Else
'          Desefetiva_Saída = -1 'Ação cancelada
'          Exit Function
'        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Desefetiva Saída")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Desefetiva_Saída = -1 'Ação cancelada
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select

End Function

Public Function Efetiva_Saída(ByVal Filial As Integer, ByVal Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '07/08/2002 - mpdea
  'Inserido os recordsets que estavam a nível modular sem necessidade,
  'ocupando mais memória
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
  Dim rsParametros As Recordset
  Dim rsOp_Saída As Recordset
  Dim rsContas_Receber As Recordset
  Dim rsResumo_Diário As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
'  Dim rsEstoque_Final As Recordset
  Dim rsPreços As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
  Dim rsGrade As Recordset
  Dim rsSaidas As Recordset
  Dim rsSaidas_Prod As Recordset
  Dim rsSaidas_Serv As Recordset
  Dim rsSaída_Cheques As Recordset
  Dim rsSaída_Parcelas As Recordset
  Dim rsComissões As Recordset
  Dim rsComissões_Serv As Recordset
  Dim rsFuncionarios As Recordset
  Dim rsTabelas As Recordset
  Dim rsConta_Cli As Recordset
  Dim rsCartoes As Recordset
  Dim rsBancos As Recordset
  Dim rsEdicoes As Recordset
  Dim rsServicos As Recordset
  '10/12/2009 - Andrea
  Dim rsSaída_Cartoes As Recordset
  '---------------------------------------------------------------------------------
 
 Dim nI As Integer
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
 Dim Tot_Parcelas As Double
 Dim Cód As String
 Dim Cód_Serv As Integer
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim Edição As Long
 Dim Estoque_Final As Double
 Dim Criar_Registro As Integer
 Dim Mensagem As String
 Dim Saldo_Conta As Double
 Dim Val_Cheques As Double
 Dim Val_Cheques_Pré As Double
 Dim Comissão As Double
 Dim Aux_Código As String
 Dim Aux_Tipo As Integer
 Dim Aux_Erro As Integer
 Dim Saldo_Emp As Long
 Dim Ordem_Emp As Long
 Dim Comiss_Técnico As Single
 Dim Val_Vista As Double
 Dim Emp_Existe As Boolean
 Dim sDescrAdicional As String
 
 Dim strAuxiliar As String
 
  
 '10/12/2009 - Andrea
 Dim Val_Cartoes As Double
  
  
  'Variável de Tratamento de Erro
  Dim blnCaseCaixa As Boolean
  Dim intRepeatUpdateLocked As Integer
  Dim intRepeatUpdate3022 As Integer
  
  'Estrutura de formas de pagamentos
  Dim typTotalizadores As tpPaymentType
  
  'Variáveis WEB
  Dim lngWEB_ID As Long
  Dim strStatusShopper As String
  Dim strStatusAdmin As String
  Dim strListPrice As String
  Dim intCodOpVenda As Integer
  Dim intCodOpReserva As Integer
  
  Dim strSQL As String
 
  '29/10/2002 - mpdea
  'Código do cliente
  Dim lngCodCliente As Long
  
  
  '28/09/2005 - mpdea
  'Nome do cliente
  Dim strNomeCliente As String
  
 
  Dim blnDiminuiComissao  As Boolean
  Dim rstTabelaPrecos     As Recordset
  Dim dblValorComissao    As Double
 
  '16/10/2004 - Daniel
  'Adicionada a var PrecoVenda que será tratada em GeraAcertoConsignacao
  'Case: Resultado
  Dim dblPrecoVenda As Double
  
  '03/07/2006 - mpdea
  'Comissão com retenção
  'Case.....: Bem Me Quer
  'Projeto..: Retenção sobre comissões
  Dim dblVlPagoEmCartao As Double
  Dim dblVlPagoEmCartaoComRetencao As Double
  Dim sngTaxaRetencao As Single
  Dim rstCartoes As Recordset
  Dim dblAuxi    As Double
  
  '17/07/2006 - Andrea
  'Comissao com retencao - acerto dos calculos
  Dim dblTotalSaidas   As Double
  Dim dblRecebe_Cartao As Double
  Dim dblPercentualPagoEmCartao As Double
   
  On Error GoTo ErrHandler
  
 
  Screen.MousePointer = vbHourglass
  
  Set rsSaidas = db.OpenRecordset("Saídas")
  Set rsSaidas_Prod = db.OpenRecordset("Saídas - Produtos", , dbReadOnly)
  Set rsSaidas_Serv = db.OpenRecordset("Saídas - Serviços", , dbReadOnly)
  Set rsSaída_Cheques = db.OpenRecordset("Movimento - Cheques", , dbReadOnly)
  Set rsSaída_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsSaída_Cartoes = db.OpenRecordset("Movimento - Cartoes", , dbReadOnly)
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsOp_Saída = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsResumo_Diário = db.OpenRecordset("Resumo Diário")
  Set rsEtiquetas = db.OpenRecordset("Etiquetas")
  Set rsRes_Financeiro = db.OpenRecordset("Resumo Diário Financeiro")
  ' Set rsEstoque = db.OpenRecordset("Estoque")
  ' Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
  Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
  Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
  Set rsEmprestimos = db.OpenRecordset("Consignação Saída")
  Set rsCliFor = db.OpenRecordset("Cli_For")
  Set rsGrade = db.OpenRecordset("Códigos da Grade")
  'Set rsComissões = db.OpenRecordset("Comissão")
  Set rsComissões_Serv = db.OpenRecordset("Comissão Serviços")
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Preços", , dbReadOnly)
  Set rsConta_Cli = db.OpenRecordset("Conta Cliente")
  Set rsCartoes = db.OpenRecordset("Cartões", , dbReadOnly)
  Set rsBancos = db.OpenRecordset("Bancos", , dbReadOnly)
  Set rsServicos = db.OpenRecordset("Serviços", , dbReadOnly)

  Screen.MousePointer = vbDefault
   
  rsSaidas.Index = "Sequência"
  rsSaidas.Seek "=", Filial, Mov
  If rsSaidas.NoMatch Then
    Efetiva_Saída = 1
    Exit Function
  End If
  
  'Verifica se a saída já foi efetivada
  'Check realizado devido a problemas de a função estar sendo
  'chamada várias vezes em alguma circunstância ainda não encontrada
  'mpdea 17/08/2000
  If rsSaidas("Efetivada") Then
    Efetiva_Saída = 0
    Exit Function
  End If
  
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  '
  'lngWEB_ID > 0 = movimentação do tipo WEB
  'Inibido erro conforme marcações
  '---------------------------------------------------------------------------------
  lngWEB_ID = CLng("0" & rsSaidas.Fields("WebOrderFormID").Value)
  
  
  Rem Encontra a tabela de operações
  rsOp_Saída.Index = "Código"
  rsOp_Saída.Seek "=", rsSaidas("Operação")
  If rsOp_Saída.NoMatch Then
    Efetiva_Saída = 2
    Exit Function
  End If
  
  Rem Encontra Cliente
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", rsSaidas("Cliente")
  If rsCliFor.NoMatch And lngWEB_ID = 0 Then '-> Inibe erro
    Efetiva_Saída = 3
    Exit Function
  Else
    '29/10/2002 - mpdea
    'Código do cliente
    lngCodCliente = rsCliFor.Fields("Código").Value
    
    '28/09/2005 - mpdea
    'Nome do cliente
    strNomeCliente = rsCliFor.Fields("Nome").Value & ""
    
    '20/12/2005 - mpdea
    'Quick Fiscal
    'Variáveis para impressão de dados em cupom de parcelamento
    'Bematech
    'CASE: Margarete Parizoto ME (QS71277-474)
    g_str_nome_cliente = strNomeCliente
    g_lng_nr_sequencia = Mov
  End If

  Rem Encontra Funcionário
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", rsSaidas("Digitador")
  If rsFuncionarios.NoMatch And lngWEB_ID = 0 Then '-> Inibe erro
    Efetiva_Saída = 4
    Exit Function
  End If

  '---------------------------------------------------------------------------------
  '07/05/2002 - mpdea
  '
  'Alterado verificação da existência da tabela de preços para operações do tipo
  'WEB (tabela de preços dinâmica [DB:Preços])
  'Somente verifica se o campo WebOrderFormID = 0 (venda não WEB)
  '---------------------------------------------------------------------------------
  If lngWEB_ID = 0 Then
    'Encontra tabela de preços
    rsTabelas.Index = "Tabela"
    rsTabelas.Seek "=", rsSaidas("Tabela")
    If rsTabelas.NoMatch Then
      Efetiva_Saída = 5
      Exit Function
    End If
  End If
  '---------------------------------------------------------------------------------
 
 
  
  Screen.MousePointer = vbHourglass
 
' Call ws.BeginTrans

  '---[ Atualiza a data da movimentação para a data atual, devido a problemas com o estoque, financeiro, etc. ]---'
    With rsSaidas
      If (.Fields("Data") <> Data_Atual) Then
        MsgBox "Atenção !" & vbCrLf & vbCrLf & "Essa movimentação foi gerada no dia " & .Fields("Data") & ". A data da movimentação está sendo ajustada para " & Data_Atual, vbInformation, "Quick Store"
        
        .LockEdits = True
        .Edit
        .Fields("Data") = Data_Atual
        .Update
        .LockEdits = False
        
        .Index = "Sequência"
        .Seek "=", Filial, Mov
        If .NoMatch Then
          Efetiva_Saída = 1
          Exit Function
        End If
      End If
    End With
  '---[ Atualiza a data da movimentação para a data atual, devido a problemas com o estoque, financeiro, etc. ]---'

  '12/07/2002 - mpdea
  'Obtém códigos de operação de saída WEB
  Call GetWEBCod_Op(intCodOpReserva, intCodOpVenda, 0)
  
  '---------------------------------------------------------------------------------
  '07/05/2002 - mpdea
  '
  'Implementado exclusão de tabelas temporárias para a venda do tipo WEB
  'e atualização do pedido e histórico para Pagamento Recebido
  '
  '12/07/2002
  '
  'Somente executa quando a operação for de venda, conforme configuração
  '
  '28/06/2005 - Daniel
  '
  'Existiram casos em que o usuário fazia o recebimento (efetivação da venda)
  'com outra operação distinta da original da venda virtual e ao tentar prosseguir,
  'no gerenciador de pedidos, o Quick dava a seguinte mensagem:
  '"Efetue o recebimento na tela de Saídas para confirmar o Pagamento"
  'Ocorrências: Osório (SEBO)
  'Antiga condição..: If lngWEB_ID <> 0 And rsSaidas.Fields("Operação").Value = intCodOpVenda Then
  'Nova condição....: If lngWEB_ID <> 0 And (rsSaidas.Fields("Operação").Value = intCodOpVenda Or rsOp_Saída.Fields("AlteraStatusPedidoWeb").Value) Then
  '---------------------------------------------------------------------------------
  If lngWEB_ID <> 0 And (rsSaidas.Fields("Operação").Value = intCodOpVenda Or rsOp_Saída.Fields("AlteraStatusPedidoWeb").Value) Then
    'Exclui tabelas temporárias
    strListPrice = Replace(LIST_PRICE_WEB, REPLACE_TQW, _
                           Format(lngWEB_ID, String(Len(REPLACE_TQW), "0")))
    Call db.Execute("DELETE * FROM Preços WHERE Tabela = '" & strListPrice & _
      "'", dbFailOnError)
    
    'Obtém descrição para o status de pagamento recebido
    Call GetDataDescPasso(ofsConfirmedPayment, strStatusShopper, strStatusAdmin)
    
    'Atualiza o Pedido
    Call db.Execute("UPDATE WEB_OrderForms SET " & _
      "StatusShopper = '" & strStatusShopper & _
      "', StatusAdmin = '" & strStatusAdmin & _
      "', Passo = " & ofsConfirmedPayment & " WHERE ID = " & lngWEB_ID, dbFailOnError)
    
    'Atualiza o Histórico do Pedido
    Call db.Execute("INSERT INTO WEB_OrderStatusHistoric " & _
      "(OrderFormID, Passo, StatusShopper, StatusAdmin, Data, WebSynchronize) " & _
      "VALUES (" & lngWEB_ID & ", " & ofsConfirmedPayment & ", '" & strStatusShopper & _
      "', '" & strStatusAdmin & "', #" & Format(Now, "MM/DD/YYYY HH:MM:SS") & _
      "#, True)", dbFailOnError)
  End If
  '---------------------------------------------------------------------------------

  '12/07/2002 - mpdea
  'Desvia em operação de reserva WEB
  If rsSaidas.Fields("Operação").Value <> intCodOpReserva Then
    If rsOp_Saída("Tipo") = "V" Then
      With rsCliFor
        .LockEdits = True
        .Edit
        .Fields("Última Compra").Value = rsSaidas.Fields("Data").Value
        .Fields("Data Alteração").Value = Format(Date, "dd/mm/yyyy")
        .Update
        .LockEdits = False
      End With
    End If
  End If
  
 Rem Atualiza Caixa, se for o caso
 'frmEntradas.Percent.Value = 4
 If rsOp_Saída.Fields("Dinheiro").Value Then
 
' ' If rsSaidas("Recebe - Dinheiro") <> 0 Or rsSaidas("Recebe - Cartão") <> 0 Or rsSaidas("Recebe - Vale") <> 0 Then
'    Erro = False
'    Caixa_Novo = False
'    Ordem = 0
'
'    rsCaixa.Index = "Data"
'    rsCaixa.Seek "<", Filial, rsSaidas("Caixa"), rsSaidas("Data"), 9999
'    If rsCaixa.NoMatch Then Caixa_Novo = True
'    If Caixa_Novo = False Then If rsCaixa("Filial") <> Filial Then Caixa_Novo = True
'    If Caixa_Novo = False Then If rsCaixa("Data") <> rsSaidas("Data") Then Caixa_Novo = True
'    If Caixa_Novo = False Then If rsCaixa("Caixa") <> rsSaidas("Caixa") Then Caixa_Novo = True
'
'    If Caixa_Novo = True Then 'Começa o Caixa do dia
'       Erro = False
'       rsCaixa.Index = "Data"
'       rsCaixa.Seek "<", Filial, rsSaidas("Caixa"), rsSaidas("Data"), 0
'       If rsCaixa.NoMatch Then Erro = True
'       If Not Erro Then If rsCaixa("Filial") <> Filial Then Erro = True
'       If Not Erro Then If rsCaixa("Caixa") <> rsSaidas("Caixa") Then Erro = True
'       If Erro = True Then  'Não existe dia anterior
'          rsCaixa.AddNew
'           rsCaixa("Filial") = Filial
'           rsCaixa("Caixa") = rsSaidas("Caixa")
'           rsCaixa("Data") = rsSaidas("Data")
'           rsCaixa("Hora") = Format(Time, "hh:mm:ss")
'           Ordem = 1
'           rsCaixa("Ordem") = Ordem
'           rsCaixa("Saldo Anterior") = 0
'           rsCaixa("Final") = 0
'           rsCaixa("Descrição") = "Início do dia"
'          rsCaixa.Update
'       Else
'          Ordem = 1
'          Saldo_Ant = rsCaixa("Final")
'          Tot_Dinheiro = gsHandleNull(rsCaixa("Total Dinheiro"))
'          Tot_Cheques = gsHandleNull(rsCaixa("Total Cheques"))
'          Tot_Cheques_Pre = gsHandleNull(rsCaixa("Total Cheques Pré"))
'          Tot_Cartões = gsHandleNull(rsCaixa("Total Cartões"))
'          Tot_Vales = gsHandleNull(rsCaixa("Total Vales"))
'          Tot_Parcelas = gsHandleNull(rsCaixa("Total Parcelamento"))
'
'          rsCaixa.AddNew
'            rsCaixa("Filial") = Filial
'            rsCaixa("Data") = rsSaidas("Data")
'            rsCaixa("Hora") = Format(Time, "hh:mm:ss")
'            rsCaixa("Caixa") = rsSaidas("Caixa")
'            rsCaixa("Ordem") = Ordem
'            rsCaixa("Funcionário") = rsSaidas("Operador")
'            rsCaixa("Descrição") = "Início do dia"
'            rsCaixa("Saldo Anterior") = Saldo_Ant
'            rsCaixa("Dinheiro") = Tot_Dinheiro
'            rsCaixa("Cheques") = Tot_Cheques
'            rsCaixa("Cheques Pré") = Tot_Cheques_Pre
'            rsCaixa("Cartões") = Tot_Cartões
'            rsCaixa("Vales") = Tot_Vales
'            rsCaixa("Total Dinheiro") = Tot_Dinheiro
'            rsCaixa("Total Cheques") = Tot_Cheques
'            rsCaixa("Total Cheques Pré") = Tot_Cheques_Pre
'            rsCaixa("Total Cartões") = Tot_Cartões
'            rsCaixa("Total Vales") = Tot_Vales
'            rsCaixa("Total Parcelamento") = Tot_Parcelas
'            rsCaixa("Final") = Saldo_Ant
'          rsCaixa.Update
'      End If
'
'      rsCaixa.Index = "Caixa"
'      rsCaixa.Seek "<", Filial, rsSaidas("Data"), rsSaidas("Caixa"), 9999
'    End If

        
'    Rem Neste ponto tem o último caixa no buffer
'    Ordem = rsCaixa("Ordem")
'    Ordem = Ordem + 1
'    Saldo_Ant = rsCaixa("Final")
'    Tot_Dinheiro = rsCaixa("Total Dinheiro")
'    Tot_Cheques = rsCaixa("Total Cheques")
'    Tot_Cheques_Pre = rsCaixa("Total Cheques Pré")
'    Tot_Cartões = rsCaixa("Total Cartões")
'    Tot_Vales = rsCaixa("Total Vales")
'    Tot_Parcelas = rsCaixa("Total Parcelamento")
    
     Rem Acha cheques
    Val_Cheques = 0
    rsSaída_Cheques.Index = "Ordem"
    Ordem = 0
    Erro = False
    Do
     rsSaída_Cheques.Seek ">", Filial, Mov, Ordem
     If rsSaída_Cheques.NoMatch Then Erro = True
     If Erro = False Then If rsSaída_Cheques("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSaída_Cheques("Sequência") <> Mov Then Erro = True

     If Erro = False Then
       If rsSaída_Cheques("Bom") = rsSaidas("Data") Then
         Val_Cheques = Val_Cheques + rsSaída_Cheques("Valor")
       End If
       If rsSaída_Cheques("Bom") <> rsSaidas("Data") Then
         Val_Cheques_Pré = Val_Cheques_Pré + rsSaída_Cheques("Valor")
       End If
       Ordem = rsSaída_Cheques("Ordem")
     End If
    Loop Until Erro = True


    'Acha pagamentos em cartões feitos na tela de saidas
    Val_Cartoes = 0
    rsSaída_Cartoes.Index = "Ordem"
    Ordem = 0
    Erro = False
    Do
     rsSaída_Cartoes.Seek ">", Filial, Mov, Ordem
     If rsSaída_Cartoes.NoMatch Then Erro = True
     If Erro = False Then If rsSaída_Cartoes("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSaída_Cartoes("Sequência") <> Mov Then Erro = True

     If Erro = False Then
       Val_Cartoes = rsSaída_Cartoes("Valor")
       Ordem = rsSaída_Cartoes("Ordem")
     End If
    Loop Until Erro = True
    
    
    '-----------------------------------------------------------------------------
    '06/05/2004 - mpdea
    '
    'Caixa
    '-----------------------------------------------------------------------------
    Set rsCaixa = db.OpenRecordset("Caixa")
    '
TryCaixa3022:
'
    blnCaseCaixa = True
    'Verifica o início do caixa, abertura do dia e retorna os últimos valores
    If Not gbCheckOpenCaixa(rsSaidas.Fields("Caixa").Value, _
      rsSaidas.Fields("Operador").Value, Saldo_Ant, Ordem, _
      typTotalizadores, False, True) Then
      'Ocorreu erro e a mensagem é exibida pela função
      Efetiva_Saída = 9
      Exit Function
    Else
      Ordem = Ordem + 1
      Tot_Dinheiro = typTotalizadores.dblDinheiro
      Tot_Cheques = typTotalizadores.dblCheque
      Tot_Cheques_Pre = typTotalizadores.dblChequePre
      Tot_Cartões = typTotalizadores.dblCartao
      Tot_Vales = typTotalizadores.dblVale
      Tot_Parcelas = typTotalizadores.dblParcelamento
    End If
    '
    With rsCaixa
      
      If .EditMode = dbEditAdd Then
        .CancelUpdate
      End If
      
      .AddNew
      .Fields("Filial").Value = Filial
      .Fields("Data").Value = rsSaidas("Data")
      .Fields("Caixa").Value = rsSaidas("Caixa")
      .Fields("Ordem").Value = Ordem
      .Fields("Funcionário").Value = rsSaidas("Operador")
      
      .Fields("Descrição").Value = "Saída nr. " & Mov
      
      '28/01/2005 - Daniel
      'Cliente Taupys - ES solicitou que a Referência saísse na descrição caso
      'estiver preenchida
      If Len(rsSaidas("Referência").Value) > 0 Then
        .Fields("Descrição").Value = .Fields("Descrição").Value & _
                                     " Ref. " & rsSaidas("Referência").Value
      End If
      
      '28/09/2005 - mpdea
      'Incluído o nome do cliente
      '
      '29/10/2002 - mpdea
      'Adicionado o código do cliente na descrição do registro de caixa
      If lngCodCliente > 0 Then
        .Fields("Descrição").Value = Left(.Fields("Descrição").Value & _
                                     " Cliente " & lngCodCliente & " - " & _
                                     strNomeCliente, .Fields("Descrição").Size)
      End If
      
      .Fields("Saldo Anterior").Value = Saldo_Ant
      .Fields("Cartões").Value = rsSaidas("Recebe - Cartão")
      .Fields("Total Cartões").Value = Tot_Cartões + rsSaidas("Recebe - Cartão")
      .Fields("Vales").Value = rsSaidas("Recebe - Vale")
      .Fields("Total Vales").Value = Tot_Vales + rsSaidas("Recebe - Vale")
      .Fields("Cheques").Value = Val_Cheques
      .Fields("Total Cheques").Value = Tot_Cheques + Val_Cheques
      .Fields("Cheques Pré").Value = Val_Cheques_Pré
      .Fields("Total Cheques Pré").Value = Tot_Cheques_Pre + Val_Cheques_Pré
      .Fields("Dinheiro").Value = rsSaidas("Recebe - Dinheiro")
      .Fields("Total Dinheiro").Value = Tot_Dinheiro + rsSaidas("Recebe - Dinheiro")
      .Fields("Parcelamento").Value = CDbl("0" & rsSaidas("Total Prazo"))
      .Fields("Total Parcelamento").Value = Tot_Parcelas + CDbl("0" & rsSaidas("Total Prazo"))
      .Fields("Final").Value = Tot_Dinheiro + rsSaidas("Recebe - Cartão") + rsSaidas("Recebe - Vale") + rsSaidas("Recebe - Dinheiro") + Val_Cheques + Val_Cheques_Pré + Tot_Cartões + Tot_Vales + Tot_Cheques + Tot_Cheques_Pre
      .Fields("Hora").Value = Format(Time, "hh:mm:ss")
      .Update
      .Close
    End With
    Set rsCaixa = Nothing
    blnCaseCaixa = False
    '-----------------------------------------------------------------------------
    
 ' End If
 End If
 
 '23/02/2005 - Daniel
 '
 'Solicitante: MRPR Automação
 '
 'Para a base deste cliente os campos abaixo em
 'algum momento estavam vindo com valores nulos.
 'Para isso foi criado o tratamento abaixo para não
 'ocorrer o erro 94 Invalid of null
 If (Not IsNull(rsSaidas("Recebe - Dinheiro"))) And (Not IsNull(rsSaidas("Recebe - Vale"))) Then
    Val_Vista = rsSaidas("Recebe - Dinheiro") + rsSaidas("Recebe - Vale")
 Else
    Val_Vista = 0
 End If
 
 Rem Fazer Lançamentos em Controle de Cheques
 rsSaída_Cheques.Index = "Ordem"
 rsBancos.Index = "Código"
 Ordem = 0
 
 Do
  Erro = False
  rsSaída_Cheques.Seek ">", Filial, Mov, Ordem
  If rsSaída_Cheques.NoMatch Then Erro = True
  If Erro = False Then If rsSaída_Cheques("Filial") <> Filial Then Erro = True
  If Erro = False Then If rsSaída_Cheques("Sequência") <> Mov Then Erro = True
  
  If Erro = False Then
    Ordem = rsSaída_Cheques("Ordem")
    'If rsSaída_Cheques("Bom") <> rsSaidas("Data") Then
      rsContas_Receber.AddNew
        
        rsContas_Receber("Tipo") = "C"
        rsContas_Receber("Filial") = Filial
        rsContas_Receber("Sequência") = Mov
        rsContas_Receber("Cliente") = rsSaidas("Cliente")
        rsContas_Receber("Banco") = rsSaída_Cheques("Banco")
        rsBancos.Seek "=", rsSaída_Cheques("Banco")
        If rsBancos.NoMatch Then rsContas_Receber("Banco") = 999
        rsContas_Receber("Cheque") = rsSaída_Cheques("Cheque")
        rsContas_Receber("Vencimento") = rsSaída_Cheques("Bom")
        rsContas_Receber("Valor") = rsSaída_Cheques("Valor")
        rsContas_Receber("Vendedor") = rsSaidas("Digitador")
        rsContas_Receber("Data Emissão") = rsSaidas("Data")
        rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")
        If rsSaída_Cheques("Bom") = rsSaidas("Data") Then
          rsContas_Receber("Processado") = True
          rsContas_Receber("Valor Recebido") = rsSaída_Cheques("Valor")
          rsContas_Receber("Data Recebimento") = rsSaída_Cheques("Bom")
        End If
        
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
          "modEfetivaSaida_Efetiva_Saída", _
          "Contas a Receber", g_strArquivoSystemLog
        End If
      rsContas_Receber.Update
     
'     24/02/2003 - Maikel
'     Descrição : Na tabela de contas a receber era gravado o cheque duas vezes, uma linha era do que foi informado
'                 na tela de recebimento, outra era de pagamento a vista segundo as linhas comentadas abaixo. Consequencia:
'                 no relatório de fluxo de caixa aparecia o cheque duas vezes.
'---------------------------------------------------------------------------------
'     Rem Grava conta recebida
'     If rsParametros("Gerar Conta Paga") = True Then
'      If rsSaída_Cheques("Bom") = rsSaidas("Data") Then
'       rsContas_Receber.AddNew
'         rsContas_Receber("Tipo") = "R"
'         rsContas_Receber("Filial") = Filial
'         rsContas_Receber("Cliente") = rsSaidas("Cliente")
'         rsContas_Receber("Data Emissão") = rsSaidas("Data")
'         rsContas_Receber("Descrição") = "Pagamento à vista"
'         rsContas_Receber("Vencimento") = rsSaidas("Data")
'         rsContas_Receber("Valor") = rsSaída_Cheques("Valor")
'         rsContas_Receber("Sequência") = Mov
'         rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
'         rsContas_Receber("Vendedor") = rsSaidas("Digitador")
'         rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")
'         rsContas_Receber("Data Recebimento") = rsSaidas("Data")
'         rsContas_Receber("Valor Recebido") = rsSaída_Cheques("Valor")
'       rsContas_Receber.Update
'      End If
'     End If
'---------------------------------------------------------------------------------

   ' End If
  End If
 Loop Until Erro = True

 
 Rem Faz Lançamentos em controle de cartões, se for o caso
 If rsSaidas("Recebe - Cartão") <> 0 Then
  
   rsContas_Receber.Index = "Contas"
   rsContas_Receber.Seek ">", "O", gnCodFilial, rsSaidas("Sequência"), 0
   If Not rsContas_Receber.NoMatch Then
     If rsContas_Receber("Tipo") = "O" Then
       If rsContas_Receber("Filial") = gnCodFilial Then
         If rsContas_Receber("Sequência") = rsSaidas("Sequência") Then
            '10/09/2007 - Anderson
            'Gera arquivo log do sistema
            If g_bolSystemLog Then
              SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
              "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
              "modEfetivaSaida_Efetiva_Saída", _
              "Contas a Receber", g_strArquivoSystemLog
            End If
           rsContas_Receber.Delete
         End If
       End If
    End If
   End If
   
   '10/12/2009 - Andrea
   'Verifica se a variável Val_Cartoes = 0 (esta variável acumula o valor recebido em cartões na tela de Recebimento)
   'Se tiver zerada, o valor que foi recebido em cartões veio do Quick Fiscal
   If Val_Cartoes = 0 Then
          
     For nI = 0 To rsSaidas("Qtde Parcelas") - 1
       rsCartoes.Index = "Código"
       rsCartoes.Seek "=", rsSaidas("Recebe - Emp Cartão")
       If Not rsCartoes.NoMatch Then
         rsContas_Receber.AddNew
         rsContas_Receber("Tipo") = "O"
         rsContas_Receber("Filial") = gnCodFilial
         rsContas_Receber("Sequência") = rsSaidas("Sequência")
         rsContas_Receber("Cliente") = rsSaidas("Cliente")
         rsContas_Receber("Administradora") = rsSaidas("Recebe - Emp Cartão")
         rsContas_Receber("Cartão") = rsSaidas("Recebe - Num Cartão")
         
         '08/10/2007 - Anderson
         'Alteração para evitar que o vencimento dos cartões seja nos finais de semana
         'Solitante: Agrotama
         rsContas_Receber("Vencimento") = (rsSaidas("Data") + rsCartoes("Dias Pagar") + (nI * 30))
         If (Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 1 Or Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 2) Then
           rsContas_Receber("Vencimento") = DateAdd("d", 3 - Weekday(rsContas_Receber("Vencimento"), vbSaturday), rsContas_Receber("Vencimento"))
         End If
         
         rsContas_Receber("Data Emissão") = rsSaidas("Data")
         If rsSaidas("Qtde Parcelas") = 1 Then
           rsContas_Receber("Valor Cartão") = rsSaidas("Recebe - Cartão")
           rsContas_Receber("Valor") = Round(CDbl(rsSaidas("Recebe - Cartão") * ((1 - rsCartoes("Taxa") / 100))), 2)
         Else
           rsContas_Receber("Valor Cartão") = rsSaidas("Valor Parcela")
           rsContas_Receber("Valor") = Round(CDbl(rsSaidas("Valor Parcela") * ((1 - rsCartoes("Taxa") / 100))), 2)
         End If
         rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")
         
         '10/09/2007 - Anderson
         'Gera arquivo log do sistema
         If g_bolSystemLog Then
           SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
           "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
           "modEfetivaSaida_Efetiva_Saída", _
           "Contas a Receber", g_strArquivoSystemLog
         End If
         rsContas_Receber.Update
       End If
     Next nI
  '10/12/2009 - Andrea
  'Senão (se foi preenchido em outras telas) gravado os cartões no contas a receber a partir da tabela Movimento - Cartões
  Else
    nI = 0
    Ordem = 0
    Erro = False
    Do
     rsSaída_Cartoes.Seek ">", Filial, Mov, Ordem
     If rsSaída_Cartoes.NoMatch Then Erro = True
     If Erro = False Then If rsSaída_Cartoes("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSaída_Cartoes("Sequência") <> Mov Then Erro = True

     If Erro = False Then
       
       'Val_Cartoes = Val_Cartoes + rsSaída_Cartoes("Valor")
       rsCartoes.Index = "Nome"
       rsCartoes.Seek "=", rsSaída_Cartoes("Administradora")
       If Not rsCartoes.NoMatch Then
         For nI = 0 To rsSaída_Cartoes("Parcelas") - 1
          rsContas_Receber.AddNew
          rsContas_Receber("Tipo") = "O"
          rsContas_Receber("Filial") = gnCodFilial
          rsContas_Receber("Sequência") = rsSaidas("Sequência")
          rsContas_Receber("Cliente") = rsSaidas("Cliente")
          rsContas_Receber("Administradora") = rsCartoes("Código")
          rsContas_Receber("Cartão") = rsSaída_Cartoes("NumeroCartao")
          
          'Evita que o vencimento dos cartões seja nos finais de semana
          rsContas_Receber("Vencimento") = (rsSaidas("Data") + rsCartoes("Dias Pagar") + (nI * 30))
          If (Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 1 Or Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 2) Then
            rsContas_Receber("Vencimento") = DateAdd("d", 3 - Weekday(rsContas_Receber("Vencimento"), vbSaturday), rsContas_Receber("Vencimento"))
          End If
          
          rsContas_Receber("Data Emissão") = rsSaidas("Data")
          If rsSaída_Cartoes("Parcelas") = 1 Then 'Cartão em 1 parcela
            rsContas_Receber("Valor Cartão") = rsSaída_Cartoes("Valor")
            rsContas_Receber("Valor") = Round(CDbl(rsSaída_Cartoes("Valor") * ((1 - rsCartoes("Taxa") / 100))), 2)
          Else 'Cartão Parcelado
            If nI = 0 Then ' É a primeira parcela
              Dim dbl_valor_parcela As Double
              Dim sht_numero_parcelas As Integer
              Dim dbl_valor_parcelar As Double
              
              sht_numero_parcelas = rsSaída_Cartoes("Parcelas")
              dbl_valor_parcelar = rsSaída_Cartoes("Valor")
              dbl_valor_parcela = dbl_valor_parcelar / sht_numero_parcelas
                
              'usada para arredondamento das parcelas para valores inteiros
              'Dim dbl_adicional_primera_parcela As Double
              Dim dbl_primeira_parcela As Double
              
              'dbl_adicional_primera_parcela = 0
              dbl_valor_parcela = Round(dbl_valor_parcela, 2)
              dbl_primeira_parcela = 0
              
              '========================================================================
              ' Tratamento para dizima periódica
              '========================================================================
              If (dbl_valor_parcela * sht_numero_parcelas) < dbl_valor_parcelar Then
                dbl_primeira_parcela = dbl_valor_parcelar - (dbl_valor_parcela * sht_numero_parcelas)
                dbl_primeira_parcela = dbl_primeira_parcela + dbl_valor_parcela
                dbl_primeira_parcela = Round(dbl_primeira_parcela, 2)
                rsContas_Receber("Valor Cartão") = dbl_primeira_parcela
                rsContas_Receber("Valor") = Round(CDbl(dbl_primeira_parcela) * ((1 - rsCartoes("Taxa") / 100)), 2)
              Else
                Dim J As Double
                J = 0
                dbl_primeira_parcela = dbl_valor_parcela
                J = (dbl_valor_parcelar - (dbl_valor_parcela * sht_numero_parcelas))
                dbl_primeira_parcela = dbl_primeira_parcela + J
                dbl_primeira_parcela = Round(dbl_primeira_parcela, 2)
                rsContas_Receber("Valor Cartão") = dbl_primeira_parcela
                rsContas_Receber("Valor") = Round(CDbl(dbl_primeira_parcela) * ((1 - rsCartoes("Taxa") / 100)), 2)
              End If
            Else
              'Cartão parcelado - segunda parcela em diante
              rsContas_Receber("Valor Cartão") = rsSaída_Cartoes("Valor") / rsSaída_Cartoes("Parcelas")
              rsContas_Receber("Valor") = Round(CDbl(rsContas_Receber("Valor Cartão") * ((1 - rsCartoes("Taxa") / 100))), 2)
            End If
          End If
          rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")
          
          'Gera arquivo log do sistema
          If g_bolSystemLog Then
            SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
            "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
            "modEfetivaSaida_Efetiva_Saída", _
            "Contas a Receber", g_strArquivoSystemLog
          End If
          rsContas_Receber.Update
         Next nI
       End If

       Ordem = rsSaída_Cartoes("Ordem")
     End If
    Loop Until Erro = True
      
  End If
 End If
 
  Rem Faz contas a receber, se for o caso
 Erro = False
 Ordem = 0
 Aux_Int = 1
 rsSaída_Parcelas.Index = "Ordem"
 
 Do
   rsSaída_Parcelas.Seek ">", Filial, Mov, Ordem
   If rsSaída_Parcelas.NoMatch Then Erro = True
   If Erro = False Then If rsSaída_Parcelas("Filial") <> Filial Then Erro = True
   If Erro = False Then If rsSaída_Parcelas("Sequência") <> Mov Then Erro = True

   If Erro = False Then
     Ordem = rsSaída_Parcelas("Ordem")
       rsContas_Receber.AddNew
         rsContas_Receber("Tipo") = "R"
         rsContas_Receber("Filial") = Filial
         rsContas_Receber("Cliente") = rsSaidas("Cliente")
         rsContas_Receber("Data Emissão") = rsSaidas("Data")
         rsContas_Receber("Parcela") = Trim(str(Aux_Int))
         rsContas_Receber("Descrição") = "Parcela " & str(Aux_Int) & "/" & str(rsSaída_Parcelas("Parcelas"))
         rsContas_Receber("Vencimento") = rsSaída_Parcelas("Bom")
         rsContas_Receber("Valor") = rsSaída_Parcelas("Valor")
         rsContas_Receber("Sequência") = Mov
         '21/02/2005 - Daniel
         'Tratamento para não ocorrer o erro 94 (Invalid use of null)
         'Solicitante: MRPR Automação - Curitiba - PR
         If Not IsNull(rsSaidas("Nota Impressa")) Then
          rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
         Else
          rsContas_Receber("Nota") = 0
         End If
         'O erro 94 estava ocorrendo exatamente na linha abaixo que ficou
         'comentada e tratada logo em seguida
         'If Val(rsSaidas("Nota Impressa")) <> 0 Then
         '  rsContas_Receber("Fatura") = str(rsSaidas("Nota Impressa")) + "\" + str(Aux_Int)
         'End If
         If IsNumeric(rsSaidas("Nota Impressa")) Then
          If rsSaidas("Nota Impressa") <> 0 Then rsContas_Receber("Fatura") = str(rsSaidas("Nota Impressa")) + "\" + str(Aux_Int)
         End If
         '--------------------------------------------------------------------------------------------------
         rsContas_Receber("Vendedor") = rsSaidas("Digitador")
         rsContas_Receber("Tipo Parcelamento") = rsSaidas("Tipo Parcela")
         rsContas_Receber("Conta Boleto") = rsSaidas("Conta")
         rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")

        '25/09/2007 - Anderson
        'Implementação do campo código de Barras para impressão em Carnês
        If rsContas_Receber("Tipo Parcelamento") = "T" Then
          rsContas_Receber("CarneCodigoBarras") = "*" & Format(Filial, "00") & Format(rsSaidas("Cliente"), "000000") & Format(Mov, "000000") & Format(Trim(str(Aux_Int)), "00") & "*"
        End If
         
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
          "modEfetivaSaida_Efetiva_Saída", _
          "Contas a Receber", g_strArquivoSystemLog
        End If
       rsContas_Receber.Update
       Aux_Int = Aux_Int + 1
   End If
 Loop Until Erro = True
 
 
 Rem atualiza contas a receber para vendas a vista, se for o caso
 If Val_Vista <> 0 Then
   If rsParametros("Gerar Conta Paga") = True Then
     rsContas_Receber.AddNew
       rsContas_Receber("Tipo") = "R"
       rsContas_Receber("Filial") = Filial
       rsContas_Receber("Cliente") = rsSaidas("Cliente")
       rsContas_Receber("Data Emissão") = rsSaidas("Data")
       rsContas_Receber("Descrição") = "Pagamento à vista"
       rsContas_Receber("Vencimento") = rsSaidas("Data")
       rsContas_Receber("Valor") = Val_Vista
       rsContas_Receber("Sequência") = Mov
       rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
       rsContas_Receber("Vendedor") = rsSaidas("Digitador")
       rsContas_Receber("Data Alteração") = Format(Date, "dd/mm/yyyy")
       rsContas_Receber("Data Recebimento") = rsSaidas("Data")
       rsContas_Receber("Valor Recebido") = Val_Vista
      
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
        "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequência") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
        "modEfetivaSaida_Efetiva_Saída", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
     rsContas_Receber.Update
   End If
 End If
 
 Rem Atualiza Resumo Diário
 If rsOp_Saída("Tipo") <> "O" Then
   rsResumo_Diário.Index = "Data"
   rsResumo_Diário.Seek "=", Filial, rsSaidas("Data")
   If rsResumo_Diário.NoMatch Then
     rsResumo_Diário.AddNew
     rsResumo_Diário("Filial") = Filial
     rsResumo_Diário("Data") = rsSaidas("Data")
   Else
     rsResumo_Diário.LockEdits = True
     rsResumo_Diário.Edit
   End If
   If rsOp_Saída("Tipo") = "V" Then
        rsResumo_Diário("Valor Vendas") = CDbl(rsResumo_Diário("Valor Vendas")) + CDbl(rsSaidas("Total")) - CDbl(rsSaidas("Serviços"))
        rsResumo_Diário("Valor Serviços") = rsResumo_Diário("Valor Serviços") + CDbl(rsSaidas("Serviços"))
   End If
   If rsOp_Saída("Tipo") = "T" Then rsResumo_Diário("Valor T Saída") = CDbl(rsResumo_Diário("Valor T Saída")) + CDbl(rsSaidas("Total"))
   If rsOp_Saída("Tipo") = "A" Then rsResumo_Diário("Valor A Saída") = CDbl(rsResumo_Diário("Valor A Saída")) + CDbl(rsSaidas("Total"))
   If rsOp_Saída("Tipo") = "G" Then rsResumo_Diário("Valor G Saída") = CDbl(rsResumo_Diário("Valor G Saída")) + CDbl(rsSaidas("Total"))
   If rsOp_Saída("Tipo") = "E" Then rsResumo_Diário("Valor E Saída") = CDbl(rsResumo_Diário("Valor E Saída")) + CDbl(rsSaidas("Total"))

  rsResumo_Diário.Update
End If
 
 
 
 
  Rem Atualiza Resumo Diário Financeiro
  If rsOp_Saída("Dinheiro") = True Then
    rsRes_Financeiro.Index = "Data"
    rsRes_Financeiro.Seek "=", Filial, rsSaidas("Data")
    If rsRes_Financeiro.NoMatch Then
       rsRes_Financeiro.AddNew
       rsRes_Financeiro("Filial") = Filial
       rsRes_Financeiro("Data") = rsSaidas("Data")
    Else
       rsRes_Financeiro.LockEdits = True
       rsRes_Financeiro.Edit
    End If
    
    If rsOp_Saída("Tipo") = "V" Then
        rsRes_Financeiro("Valor Vendas") = CDbl(rsRes_Financeiro("Valor Vendas")) + CDbl(rsSaidas("Total")) - CDbl(rsSaidas("Serviços"))
        rsRes_Financeiro("Valor Serviços") = CDbl(rsRes_Financeiro("Valor Serviços")) + CDbl(rsSaidas("Serviços"))
    End If
    If rsOp_Saída("Tipo") = "T" Then rsRes_Financeiro("Valor T Saída") = CDbl(rsRes_Financeiro("Valor T Saída")) + CDbl(rsSaidas("Total"))
    If rsOp_Saída("Tipo") = "A" Then rsRes_Financeiro("Valor A Saída") = CDbl(rsRes_Financeiro("Valor A Saída")) + CDbl(rsSaidas("Total"))
    If rsOp_Saída("Tipo") = "G" Then rsRes_Financeiro("Valor G Saída") = CDbl(rsRes_Financeiro("Valor G Saída")) + CDbl(rsSaidas("Total"))
    If rsOp_Saída("Tipo") = "E" Then rsRes_Financeiro("Valor E Saída") = CDbl(rsRes_Financeiro("Valor E Saída")) + CDbl(rsSaidas("Total"))


    rsRes_Financeiro.Update
  End If


  rsSaidas_Prod.Index = "Sequência"
  Ordem = 0
Prox_Prod:
  rsSaidas_Prod.Seek ">", Filial, Mov, Ordem
  
  If rsSaidas_Prod.NoMatch Then GoTo Ve_Serv
  If rsSaidas_Prod("Filial") <> Filial Then GoTo Ve_Serv
  If rsSaidas_Prod("sequência") <> Mov Then GoTo Ve_Serv
  
  Ordem = rsSaidas_Prod("Linha")
  'Verifica se tem grade
  Cód = rsSaidas_Prod("Código")
  Tamanho = 0
  Cor = 0
  Edição = 0
  
  If Not IsNull(rsSaidas_Prod("Descricao Adicional")) Then
     sDescrAdicional = rsSaidas_Prod("Descricao Adicional")
  Else
     sDescrAdicional = ""
  End If
  
   Aux_Código = Trim(Cód)
   Call Acha_Produto(Aux_Código, Cód, Tamanho, Cor, Edição, Aux_Tipo, Aux_Erro)
   If Aux_Erro <> 0 Then GoTo Prox_Prod
   Cód = Trim(UCase(Cód))
   rsProdutos.Index = "Código"
   rsProdutos.Seek "=", Cód
   
  'Neste ponto CÓD tem o código do produto
  'Tamanho, Cor e Edição contém os respectivos dados
  'Agora grava arquivo do estoque
  
  Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
  
  '---[ Gera acerto de consignação de entrada ]---'
  '13/08/2004 - Maikel
  '
  '27/08/2004 - Daniel
  'Validação realizada em cima do campo [Operações Saída].AcertaEmprestimoEntrada
  If rsOp_Saída("AcertaEmprestimoEntrada").Value Then
    dblPrecoVenda = Format(rsSaidas_Prod("Preço Final").Value, FORMAT_VALUE)
  
    Call GeraAcertoConsignacaoEntrada(gnCodFilial, rsSaidas_Prod("Sequência").Value, rsSaidas_Prod("Código Sem Grade").Value, rsSaidas_Prod("Qtde").Value, dblPrecoVenda)
  End If
  '---[ Gera acerto de consignação de entrada ]---'
  'testar aqui
  'Ajusta Estoque
  If rsOp_Saída("Estoque") = True And Not rsSaidas_Prod("InGeradoViaConsig") Then
  
'-------------------------------------------------------------------------------------
    '14/11/2003 - mpdea
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
             " ORDER BY Data"
             
'    'LOG ESPECIFICO PARA MARE MANSA
'    Dim sSQL_Aux As String
'    sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'    sSQL_Aux = sSQL_Aux & Left("1 - " & Filial & " : " & Cód & " : " & Tamanho & " : " & Cor & " : " & Edição, 80) & "', 'VENDENDO MARE')"
'    db.Execute sSQL_Aux, dbFailOnError
'    'fim
    
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
'        .MoveFirst
        .MoveLast
        Estoque_Final = .Fields("Estoque Final")
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
'      sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'      sSQL_Aux = sSQL_Aux & Left("2 - " & Cód & " : " & Estoque_Final & " : " & Format(sData_mare, "mm/dd/yyyy"), 80) & "', 'VENDENDO MARE')"
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
             " AND Data = #" & Format(rsSaidas("Data"), "mm/dd/yyyy") & "#"
            
'    'LOG ESPECIFICO PARA MARE MANSA
'    sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'    sSQL_Aux = sSQL_Aux & Left("3 - " & Filial & " : " & Cód & " : " & Tamanho & " : " & Cor & " : " & Edição & " : " & Format(rsSaidas("Data"), "mm/dd/yyyy"), 80) & "', 'VENDENDO MARE')"
'    db.Execute sSQL_Aux, dbFailOnError
'    'fim
            
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .LockEdits = True
        .Edit
    
'        'LOG ESPECIFICO PARA MARE MANSA
'        sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'        sSQL_Aux = sSQL_Aux & Left("4B - " & Cód & " : UPDATE NA LINHA NA TABELA DE ESTOQUE", 80) & "', 'VENDENDO MARE')"
'        db.Execute sSQL_Aux, dbFailOnError
'        'fim
      
      Else
        .AddNew
        .Fields("Filial").Value = Filial
        .Fields("Data").Value = rsSaidas("Data").Value
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
'        sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'        sSQL_Aux = sSQL_Aux & Left("4A - " & Cód & " : " & rsSaidas("Data").Value & " : " & Estoque_Final & " : " & "NOVA LINHA NA TABELA DE ESTOQUE", 80) & "', 'VENDENDO MARE')"
'        db.Execute sSQL_Aux, dbFailOnError
'        'fim
      
      End If
    End With
'-------------------------------------------------------------------------------------
    
'    Rem Encontra a posição do estoque
'    Criar_Registro = False
'    Estoque_Final = 0
'    rsEstoque.Index = "Produto"
'    rsEstoque.Seek "=", Filial, rsSaidas("Data"), Cód, Tamanho, Cor, Edição
'
'    If Not rsEstoque.NoMatch Then Estoque_Final = rsEstoque("Estoque Final")
'    If rsEstoque.NoMatch Then
'       rsEstoque.Index = "Data"
'       rsEstoque.Seek "<", Filial, Cód, Tamanho, Cor, Edição, rsSaidas("Data")
'       If rsEstoque.NoMatch Then Criar_Registro = True
'       If Not rsEstoque.NoMatch Then
'          If rsEstoque("Filial") = Filial And rsEstoque("Produto") = Cód And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edição") = Edição Then
'             Criar_Registro = True
'             Estoque_Final = rsEstoque("Estoque Final")
'           End If
'       End If
'
'       rsEstoque.AddNew
'       rsEstoque("Filial") = Filial
'       rsEstoque("Data") = rsSaidas("Data")
'       rsEstoque("Produto") = Cód
'       rsEstoque("Tamanho") = Tamanho
'       rsEstoque("Cor") = Cor
'       rsEstoque("Edição") = Edição
'       rsEstoque("Classe") = rsProdutos("Classe")
'       rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
'       rsEstoque("Estoque Anterior") = Estoque_Final
'       rsEstoque.Update
'
'       rsEstoque.Index = "Produto"
'       rsEstoque.Seek "=", Filial, rsSaidas("Data"), Cód, Tamanho, Cor, Edição
'    End If

'-------------------------------------------------------------------------------------

      Rem neste ponto esta com o registro de estoque
      Rem no buffer, agora soma com os valores da movimentação
'      rsEstoque.Edit
      If rsOp_Saída("Tipo") = "V" Then
         rsEstoque("Vendas") = rsEstoque("Vendas") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Vendas") = Format(rsEstoque("Valor Vendas") + rsSaidas_Prod("Preço Final"), "############0.00")
      End If
      If rsOp_Saída("Tipo") = "T" Then
         rsEstoque("Transf Saída") = rsEstoque("Transf Saída") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor T Saída") = Format(rsEstoque("Valor T Saída") + rsSaidas_Prod("Preço Final"), "############0.00")
      End If
      If rsOp_Saída("Tipo") = "A" Then
         rsEstoque("Ajuste Saída") = rsEstoque("Ajuste Saída") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Ajuste Saída") = Format(rsEstoque("Valor Ajuste Saída") + rsSaidas_Prod("Preço Final"), "############0.00")
      End If
      If rsOp_Saída("Tipo") = "G" Then
         rsEstoque("Grátis Saída") = rsEstoque("Grátis Saída") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Grátis Saída") = Format(rsEstoque("Valor Grátis Saída") + rsSaidas_Prod("Preço Final"), "############0.00")
      End If
      If rsOp_Saída("Tipo") = "E" Then
         rsEstoque("Empre Saída") = rsEstoque("Empre Saída") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Empre Saída") = Format(rsEstoque("Valor Empre Saída") + rsSaidas_Prod("Preço Final"), "############0.00")
      End If

      Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
      Estoque_Final = Estoque_Final - rsEstoque("Transf Saída") + rsEstoque("Transf Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Ajuste Saída") + rsEstoque("Ajuste Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Grátis Saída") + rsEstoque("Grátis Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Empre Saída") + rsEstoque("Empre Entra")
      
      '08/08/2003 - maikel
      '             Descomentada a soma da coluna Devolução para resolver o problema de estoque
      Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolução")

      If rsProdutos("Estoque") = False Then
        Estoque_Final = 0
      End If

      rsEstoque("Estoque Final") = Estoque_Final
      rsEstoque.Update
      
      rsEstoque.LockEdits = False
      
      rsEstoque.Close
      
      Rem Arruma Estoque Final
      Grava_Estoque_Final gnCodFilial, Cód, Tamanho, Cor, Edição, CSng(Estoque_Final), rsSaidas("Data")
      
      
  End If

  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Cód

  Rem Grava Conta do Cliente, se for o caso
  If rsSaidas("Recebe - Conta") = True Then
     rsConta_Cli.AddNew
     rsConta_Cli("Filial") = Filial
     rsConta_Cli("Cliente") = rsSaidas("Cliente")
     rsConta_Cli("Data") = rsSaidas("Data")
     rsConta_Cli("Produto") = Cód
     
     strAuxiliar = rsProdutos("Nome") & ""
     strAuxiliar = Replace(strAuxiliar, "        ", " ")
     strAuxiliar = Replace(strAuxiliar, "       ", " ")
     strAuxiliar = Replace(strAuxiliar, "      ", " ")
     strAuxiliar = Replace(strAuxiliar, "     ", " ")
     strAuxiliar = Replace(strAuxiliar, "    ", " ")
     
     If Len(Trim(strAuxiliar)) > 70 Then

        rsConta_Cli("Descrição") = Left(Trim(strAuxiliar & ""), 70)
     Else
        rsConta_Cli("Descrição") = strAuxiliar
     End If
     
     rsConta_Cli("Qtde") = rsSaidas_Prod("Qtde")
     rsConta_Cli("Valor") = rsSaidas_Prod("Preço Final")
     rsConta_Cli("TabPrecos") = rsSaidas("Tabela") & ""
     rsConta_Cli("Sequência") = rsSaidas("Sequência")
     rsConta_Cli("Data Alteração") = Format(Date, "dd/mm/yyyy")
     rsConta_Cli.Update
  End If

  '18/07/2003 - mpdea
  'Formatado valores com FORMAT_VALUE e
  'comissão foi truncada com 6 casas decimais
  '
  'Grava comissões
  'If rsOp_Saída("Comissão") = True Then
  '
  '  Dim iComissionado As Integer
  '  Dim iCont As Integer
  '  Dim Colecao As New Collection
  '  Colecao.Add rsSaidas("Digitador").Value
  '  If Not IsNull(rsSaidas("PrestadorServico")) Then
  '    If Val(rsSaidas("PrestadorServico")) > 0 Then Colecao.Add rsSaidas("PrestadorServico").Value
  '  End If
  '
  '  For iCont = 1 To Colecao.Count'

'      iComissionado = CInt(Colecao.Item(iCont))
'
'      '22/10/2003 - Maikel
'      '---[ Adicionada a verificação para diminuir a comissão do funcionário caso haja desconto na movimentação ]---'
'        If (rsSaidas.Fields("Desconto")) > 0 Or (rsSaidas.Fields("DescontoSubTotal") > 0) Then
'          blnDiminuiComissao = True
'        End If
'      '---[ Adicionada a verificação para diminuir a comissão do funcionário caso haja desconto na movimentação ]---'
'
'      Comissão = rsFuncionarios("Comissão")
'      If rsProdutos("Comissão Sobrepõe") = True Then
'        Comissão = rsProdutos("Comissão")
'      End If
'      Comissão = Comissão * rsTabelas("Multiplicador Comissão")
'      Comissão = Format(Comissão, FORMAT_VALUE)
'
'      Set objComissao = New clsComissao
'      objComissao.Filial = Filial
'      objComissao.Data = rsSaidas("Data")
'      objComissao.Vendedor = iComissionado
'      objComissao.Produto = Cód
'      objComissao.Tamanho = Tamanho
'      objComissao.Cor = Cor
'      objComissao.Edição = Edição
'      objComissao.Qtde = rsSaidas_Prod("Qtde")
'
'      If Not IsNull(rsSaidas("Recebe - Vale")) And rsSaidas("Recebe - Vale") <> 0 Then
'          objComissao.Valor = Format(rsSaidas_Prod("Preço Final") - rsSaidas("Recebe - Vale"), FORMAT_VALUE)
'      Else
'          objComissao.Valor = Format(rsSaidas_Prod("Preço Final"), FORMAT_VALUE)
'      End If
'
'      objComissao.Sequência = rsSaidas("Sequência")
'
'      '-----------------------------------------------------------------------------------------------------
'      '03/07/2006 - mpdea
'      'Movido códigos da tela de Recebimento para o módulo Efetiva Saída, pois
'      'não estava sendo executado no recebimento simplificado
'      '
'      '25/04/2005 - Daniel
'      '22/03/2005 - Daniel
'      'Case.....: Bem Me Quer
'      'Projeto..: Retenção sobre comissões
'      If rsSaidas.Fields("Recebe - Cartão").Value > 0 Then
'        '-----------------------------------------------------------------------------
'        '17/07/2006 - Andrea
'        'Alterado para acertar o valor pago em cartao por item para gravar no
'''        'arquivo de comissao com retencao.
        'essa linha era o que tinha 'dblVlPagoEmCartao = rsSaidas.Fields("Recebe - Cartão").Value
'        '-----------------------------------------------------------------------------
'        dblTotalSaidas = rsSaidas.Fields("Total").Value
'        dblRecebe_Cartao = rsSaidas.Fields("Recebe - Cartão").Value
''        dblPercentualPagoEmCartao = ((dblRecebe_Cartao * 100) / dblTotalSaidas)
'        dblVlPagoEmCartao = ((rsSaidas_Prod("Preço Final").Value * dblPercentualPagoEmCartao) / 100)
'        '-----------------------------------------------------------------------------
'        '
'
'        dblVlPagoEmCartaoComRetencao = 0
'
'        Set rstCartoes = db.OpenRecordset("SELECT Taxa FROM Cartões WHERE Código = " & rsSaidas.Fields("Recebe - Emp Cartão").Value, dbOpenDynaset, dbReadOnly)
'        With rstCartoes
'          If Not (.BOF And .EOF) Then
'            .MoveFirst
'            dblAuxi = Format(((dblVlPagoEmCartao * .Fields("Taxa").Value) / 100), FORMAT_VALUE)
'            dblVlPagoEmCartaoComRetencao = Format(dblVlPagoEmCartao - dblAuxi, FORMAT_VALUE)
'            sngTaxaRetencao = Format(.Fields("Taxa").Value, "##,###,##0.0000")
'          End If
'          .Close
'        End With
'        Set rstCartoes = Nothing
'      End If
'      '
'      objComissao.VlPagoEmCartao = dblVlPagoEmCartao
'      objComissao.VlPagoEmCartaoComRetencao = dblVlPagoEmCartaoComRetencao
'      objComissao.TaxaRetencao = Format(sngTaxaRetencao, "##,###,##0.0000")
'      '-----------------------------------------------------------------------------------------------------
'
'
'      If Not IsNull(rsSaidas("Recebe - Vale")) And rsSaidas("Recebe - Vale") <> 0 Then
'          dblValorComissao = (Comissão * (rsSaidas_Prod("Preço Final") - rsSaidas("Recebe - Vale")) / 100)
'      Else
'          dblValorComissao = (Comissão * rsSaidas_Prod("Preço Final") / 100)
'      End If
'
'
'      If blnDiminuiComissao Then
'        Set rstTabelaPrecos = db.OpenRecordset(" SELECT * FROM [Tabela de Preços] " & _
'                                               " WHERE Tabela = '" & rsSaidas.Fields("Tabela") & "'", dbOpenSnapshot)
'
'        With rstTabelaPrecos
'          If Not (.BOF And .EOF) Then
'            If IsNumeric(rstTabelaPrecos.Fields("PercentualComissaoDesconto")) Then
'              '11/02/2005 - Daniel
'              'Problema levantado pela Daring
'              'Se for dado desconto só em um produto em uma nota com X produtos, a comissão
'              'do vendedor em todos os ítens da nota estava caindo pela metade sendo que o
'              'correto é reduzir só a do produto que teve alteração de preço ou desconto
'              If (rsSaidas.Fields("Desconto") > 0) Then 'Houve desconto para algum ítem. Nota: Para o desconto no subtotal continuamos abater de todos os ítens
'                Dim dblValorDoCadastroProduto As Double
'
'                Call ReduzirComissao(rsSaidas.Fields("Tabela") & "", rsSaidas_Prod("Código sem Grade") & "", dblValorDoCadastroProduto)
'
'                'Se for diferente ocorre o abatimento
'                If dblValorDoCadastroProduto <> Format((rsSaidas_Prod("Preço Final").Value / rsSaidas_Prod("Qtde")), FORMAT_VALUE) Then dblValorComissao = dblValorComissao * ((100 - rstTabelaPrecos.Fields("PercentualComissaoDesconto")) / 100)
'
'              Else
'                dblValorComissao = dblValorComissao * ((100 - rstTabelaPrecos.Fields("PercentualComissaoDesconto")) / 100)
'              End If
'            End If
'          End If
'
'          .Close
'          Set rstTabelaPrecos = Nothing
'        End With
'      End If
'
'      objComissao.Comissão = Truncate(dblValorComissao, 6)
'      objComissao.Cliente = rsSaidas("Cliente")
'      objComissao.Tabela = rsSaidas("Tabela")
'      objComissao.Insert
'      Set objComissao = Nothing
'    Next
'    Set Colecao = Nothing
'
'  End If

  Rem Grava etiquetas
  If rsSaidas_Prod("Etiqueta") = True Then
    rsEtiquetas.Index = "Funcionário"
    rsEtiquetas.Seek "=", rsSaidas("Digitador"), Cód, Tamanho, Cor
    If rsEtiquetas.NoMatch Then
       rsEtiquetas.AddNew
    Else
       rsEtiquetas.LockEdits = True
       rsEtiquetas.Edit
    End If
    rsEtiquetas("Funcionário") = rsSaidas("Digitador")
    rsEtiquetas("Produto") = Cód
    rsEtiquetas("Tamanho") = Tamanho
    rsEtiquetas("Cor") = Cor
    rsEtiquetas("Qtde") = rsEtiquetas("Qtde") + rsSaidas_Prod("Qtde")
    rsEtiquetas("Sequência") = Mov
    rsEtiquetas.Update
  End If

 
  '12/07/2002 - mpdea
  'Desvia em operação de reserva WEB
  If rsSaidas.Fields("Operação").Value <> intCodOpReserva Then
    Rem Atualiza arquivo de Resumo de Clientes
    Rem se for Comrpa
    If rsOp_Saída("Tipo") = "V" Then
       rsResumo_Clientes.Index = "Cliente"
       rsResumo_Clientes.Seek "=", rsSaidas("Cliente"), rsSaidas("Data"), Cód, Tamanho, Cor, Edição, Mov
       If rsResumo_Clientes.NoMatch Then
          rsResumo_Clientes.AddNew
            rsResumo_Clientes("Dia") = rsSaidas("Data")
            rsResumo_Clientes("Cliente") = rsSaidas("Cliente")
            rsResumo_Clientes("Produto") = Cód
            rsResumo_Clientes("Tamanho") = Tamanho
            rsResumo_Clientes("Cor") = Cor
            rsResumo_Clientes("Edição") = Edição
            rsResumo_Clientes("Qtde") = 0
            rsResumo_Clientes("Valor Total") = 0
            rsResumo_Clientes("Sequência") = Mov
            rsResumo_Clientes("Descricao Adicional") = ""
  '          rsResumo_Clientes("Descricao Adicional") = sDescrAdicional
       Else
          rsResumo_Clientes.LockEdits = True
          rsResumo_Clientes.Edit
       End If
  
        rsResumo_Clientes("Qtde") = rsResumo_Clientes("Qtde") + rsSaidas_Prod("Qtde")
        rsResumo_Clientes("Valor Total") = Format((rsResumo_Clientes("Valor Total") + rsSaidas_Prod("Preço Final")), "############0.00")
        rsResumo_Clientes("Filial") = Filial
        rsResumo_Clientes("Tipo") = "C"
        rsResumo_Clientes("Descricao Adicional") = rsResumo_Clientes("Descricao Adicional") & sDescrAdicional & "-"
          
       rsResumo_Clientes.Update
    End If
  End If


  Rem Atualiza arquivo de Empréstimos
  If rsOp_Saída("Tipo") = "E" And Not rsSaidas_Prod("InGeradoViaConsig") Then
     rsEmprestimos.Index = "Cliente"
     
     
     Rem Saldo Emprestado = 0 para este empréstimo
     Saldo_Emp = 0
     Ordem_Emp = Ordem
     Emp_Existe = False
             
     rsEmprestimos.Seek "<", gnCodFilial, Mov, rsSaidas("Cliente"), Cód, Tamanho, Cor, Edição, 999999
       If Not rsEmprestimos.NoMatch Then
         If rsEmprestimos("Filial") = gnCodFilial Then
           If rsEmprestimos("Sequência") = Mov Then
             If rsEmprestimos("Cliente") = rsSaidas("Cliente") Then
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
         rsEmprestimos.LockEdits = True
         rsEmprestimos.Edit
      Else
        rsEmprestimos.AddNew
          rsEmprestimos("Filial") = gnCodFilial
          rsEmprestimos("Sequência") = Mov
          rsEmprestimos("Cliente") = rsSaidas("Cliente")
          rsEmprestimos("Produto") = Cód
          rsEmprestimos("Tamanho") = Tamanho
          rsEmprestimos("Cor") = Cor
          rsEmprestimos("Edição") = Edição
          rsEmprestimos("Ordem") = Ordem_Emp
      End If
      
      rsEmprestimos("Saldo Anterior") = Saldo_Emp
      rsEmprestimos("Novo Empréstimo") = rsSaidas_Prod("Qtde")
      rsEmprestimos("Saldo Atual") = Saldo_Emp + rsSaidas_Prod("Qtde")
      rsEmprestimos("Preço Unitário") = (rsSaidas_Prod("Preço Final") / rsSaidas_Prod("Qtde"))
      rsEmprestimos("Data Operação") = rsSaidas("Data")
      rsEmprestimos("Data Alteração") = Format(Date, "dd/mm/yyyy")
      rsEmprestimos("Data Cobrança") = rsSaidas("Data Acerto Empréstimo")

      rsEmprestimos.Update
     
  End If
  
  GoTo Prox_Prod
  
  
  
  
  
  
  
Ve_Serv:
  rsSaidas_Serv.Index = "Sequência"
  rsServicos.Index = "Código"
  rsFuncionarios.Index = "Código"
  Ordem = 0
Prox_Serv:
  rsSaidas_Serv.Seek ">", Filial, Mov, Ordem
  
  If rsSaidas_Serv.NoMatch Then GoTo Fim_Efetiva
  If rsSaidas_Serv("Filial") <> Filial Then GoTo Fim_Efetiva
  If rsSaidas_Serv("Sequência") <> Mov Then GoTo Fim_Efetiva
  
  Ordem = rsSaidas_Serv("Linha")
  
  Cód_Serv = rsSaidas_Serv("Código")
  rsServicos.Seek "=", Cód_Serv
  If rsServicos.NoMatch Then GoTo Prox_Serv
  
  Comiss_Técnico = 0
  rsFuncionarios.Seek "=", rsSaidas("Técnico")
  If Not rsFuncionarios.NoMatch Then
    Comiss_Técnico = rsFuncionarios("Comissão Serviço")
  End If
  
  If rsServicos("Comissão Sobrepõe") = True Then
    Comiss_Técnico = rsServicos("Comissão")
  End If
  
  
  '--------------------------------------------------------------------------------
  '02/09/2003 - mpdea
  'Corrigido cálculo do valor do serviço
  '
  '18/07/2003 - mpdea
  'Comissão foi truncada com 6 casas decimais
  With rsComissões_Serv
    .AddNew
    .Fields("Data").Value = rsSaidas.Fields("Data").Value
    .Fields("Vendedor").Value = rsSaidas.Fields("Técnico").Value
    .Fields("Serviço").Value = Cód_Serv
    .Fields("Descrição").Value = rsSaidas_Serv.Fields("Descrição").Value & ""
    .Fields("Tempo").Value = rsSaidas_Serv.Fields("Tempo").Value
    .Fields("Valor").Value = CDbl(Format(CDbl(rsSaidas_Serv.Fields("Tempo").Value) * rsSaidas_Serv.Fields("Preço").Value, FORMAT_VALUE))
    .Fields("Comissão").Value = Comiss_Técnico
    .Fields("Valor Comissão").Value = Truncate((.Fields("Valor").Value * Comiss_Técnico / 100), 6)
    .Fields("Sequência").Value = rsSaidas.Fields("Sequência").Value
    .Fields("Cliente").Value = rsSaidas.Fields("Cliente").Value
    .Fields("Filial").Value = rsSaidas.Fields("Filial").Value
    .Update
  End With
  
  'Grava Conta do Cliente, se for o caso
  If rsSaidas("Recebe - Conta") Then
    With rsConta_Cli
      .AddNew
     .Fields("Filial").Value = Filial
     .Fields("Cliente").Value = rsSaidas.Fields("Cliente").Value
     .Fields("Data").Value = rsSaidas.Fields("Data").Value
     .Fields("Produto").Value = Cód_Serv
     
     If Len(Trim(rsSaidas_Serv.Fields("Descrição").Value)) > 70 Then
        .Fields("Descrição").Value = Left(Trim(rsSaidas_Serv.Fields("Descrição").Value & ""), 70)
     Else
        .Fields("Descrição").Value = rsSaidas_Serv.Fields("Descrição").Value & ""
     End If
     
     .Fields("Qtde").Value = CSng(rsSaidas_Serv.Fields("Tempo").Value)
     .Fields("Valor").Value = CDbl(Format(CDbl(rsSaidas_Serv.Fields("Tempo").Value) * rsSaidas_Serv.Fields("Preço").Value, FORMAT_VALUE))
     .Fields("TabPrecos").Value = rsSaidas.Fields("Tabela").Value & "" 'Jose
     .Fields("Sequência").Value = rsSaidas.Fields("Sequência").Value
     .Fields("Data Alteração").Value = Format(Date, "dd/mm/yyyy")
     .Update
    End With
  End If
  '--------------------------------------------------------------------------------
  
  
  GoTo Prox_Serv
  
  
Fim_Efetiva:

  'Verifica programa de fidelidade
  If gParticipaProgramaFidelidade = 1 Then 'Lançar registro no programa
      ProgramaFidelidadeCriarLancamento rsSaidas.Fields("Operação").Value, rsSaidas.Fields("Total").Value, rsSaidas.Fields("Cliente").Value, rsCliFor.Fields("CGC").Value, rsSaidas.Fields("Operador").Value, Mov, rsCliFor.Fields("Nome").Value
  
      If gClienteEntregouResgatePontos = True Then
          gClienteEntregouResgatePontos = False
          gSaldoCdGuidResgate = 0
          gCdGuidResgate = ""
          gCdClienteCdGuidResgate = 0
          gNmClienteCdGuidResgate = ""
      End If
  End If
Segue_Adiante:
    
  '08/10/2003 - mpdea
  'Ajusta os valores da conta do cliente caso exista Desconto no SubTotal
  If rsSaidas.Fields("Recebe - Conta").Value Then
    If rsSaidas.Fields("DescontoSubTotal").Value > 0 Then
      Call AdjustContaCliente(CByte(Filial), Mov, rsSaidas.Fields("Total").Value, CDbl(rsSaidas.Fields("DescontoSubTotal").Value))
    End If
  End If
  
  'Efetiva a Saída
  'mpdea 17/08/2000
  With rsSaidas
    .LockEdits = True
    .Edit
    .Fields("Efetivada").Value = True
    .Fields("NSU_Hora").Value = Format(Now, "hh:mm:ss")
    .Update
    .LockEdits = False
  End With

'  Call ws.CommitTrans
  
 rsSaidas.Close
 rsContas_Receber.Close
 rsProdutos.Close
 rsParametros.Close
 rsOp_Saída.Close
 rsResumo_Diário.Close
 rsEtiquetas.Close
 rsRes_Financeiro.Close
'  If Not rsEstoque Is Nothing Then rsEstoque.Close
' rsEstoque_Final.Close
 rsPreços.Close
 rsResumo_Clientes.Close
 rsEmprestimos.Close
 rsCliFor.Close
 rsGrade.Close
 rsSaidas_Prod.Close
 rsSaidas_Serv.Close
 rsSaída_Cheques.Close
 rsSaída_Parcelas.Close
 'rsComissões.Close
 rsComissões_Serv.Close
 rsFuncionarios.Close
 rsTabelas.Close
 rsConta_Cli.Close
 rsCartoes.Close
 rsBancos.Close
 rsServicos.Close
  
  
 Set rsSaidas = Nothing
 Set rsContas_Receber = Nothing
 Set rsProdutos = Nothing
 Set rsParametros = Nothing
 Set rsOp_Saída = Nothing
 Set rsResumo_Diário = Nothing
 Set rsEtiquetas = Nothing
 Set rsRes_Financeiro = Nothing
 Set rsEstoque = Nothing
' Set rsEstoque_Final = Nothing
 Set rsPreços = Nothing
 Set rsResumo_Clientes = Nothing
 Set rsEmprestimos = Nothing
 Set rsCliFor = Nothing
 Set rsGrade = Nothing
 Set rsSaidas_Prod = Nothing
 Set rsSaidas_Serv = Nothing
 Set rsSaída_Cheques = Nothing
 Set rsSaída_Parcelas = Nothing
 'Set rsComissões = Nothing
 Set rsComissões_Serv = Nothing
 Set rsFuncionarios = Nothing
 Set rsTabelas = Nothing
 Set rsConta_Cli = Nothing
 Set rsCartoes = Nothing
 Set rsBancos = Nothing
 Set rsServicos = Nothing

   Efetiva_Saída = 0
   
   Screen.MousePointer = vbDefault
   
   frmMain.Enabled = True
 
 
  '************************************************
  'Limpa variaveis do prog. fidelidade (se existir)
  gClienteEntregouResgatePontos = False
  gSaldoCdGuidResgate_clicou_ok_telaDesconto = False
  '************************************************
 
   Exit Function

ErrHandler:
  Screen.MousePointer = vbDefault
  Select Case Err.Number
    Case 3022 And blnCaseCaixa
      If intRepeatUpdate3022 < 1000 Then
        intRepeatUpdate3022 = intRepeatUpdate3022 + 1
        Call StatusMsg("Verificando registro...")
        Screen.MousePointer = vbHourglass
        Resume TryCaixa3022
      End If
      
    Case 3186, 3187, 3197, 3218, 3260 'Registro bloqueado
      If intRepeatUpdateLocked < 30 Then
        intRepeatUpdateLocked = intRepeatUpdateLocked + 1
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        Call frmAvisoBloqueio.ShowTentativas(30 - intRepeatUpdateLocked)
        Screen.MousePointer = vbHourglass
        Call WaitSeconds(1, False) 'Aguarda um segundo
        Resume
      Else
        
        If frmAvisoBloqueio.ShowRetryCancel = vbRetry Then
          intRepeatUpdateLocked = 0
          Screen.MousePointer = vbHourglass
          Resume
        Else
          Efetiva_Saída = -1 'Ação cancelada
          Exit Function
        End If
        
'        If MsgBox("Há no momento registros sendo atualizados no sistema por outra estação." & _
'          " É necessário aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Efetiva Saída") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        Else
'          Efetiva_Saída = -1 'Ação cancelada
'          Exit Function
'        End If
      End If
    Case Else
      'Repassa para a função de origem os outros erros
      Err.Raise Err.Number, Err.Source, Err.Description
      Exit Function
      
'      'Outros Erros
'      Select Case frmErro.gnShowErr(Err.Number, "Efetiva Saída")
'        Case 0 'Repetir
'          Resume
'        Case 1 'Prosseguir
'          Resume Next
'        Case 2 'Sair
'          Efetiva_Saída = -1 'Ação cancelada
'          Exit Function
'        Case 3 'Encerrar
'          End
'      End Select
  End Select

End Function


'mpdea
'Verifica a 1ª abertura de um caixa, o seu início de dia
'ou devolve os valores atuais caso esteja aberto
Public Function gbCheckOpenCaixa(ByVal nCaixa As Byte, ByVal nFuncionario As Integer, _
  ByRef dblSaldoAnterior As Double, ByRef nOrdem As Integer, _
  ByRef tValoresAtuais As tpPaymentType, _
  Optional ByVal blnTransaction As Boolean = True, _
  Optional ByVal blnByPassErr As Boolean = False) As Boolean
  
  Dim rsCaixa As Recordset
  Dim tFinalizadora As tpPaymentType
  Dim sSql As String
  Dim blnInTransaction As Boolean
  
  
  On Error GoTo ErrHandler
  
  'Inicia transação
  If blnTransaction Then ws.BeginTrans: blnInTransaction = True
  
  sSql = "SELECT * FROM Caixa WHERE Filial = " & gnCodFilial & " AND Caixa = " & _
    nCaixa & " ORDER BY Filial, Caixa, Data, Ordem"
  Set rsCaixa = db.OpenRecordset(sSql, dbOpenDynaset)
  
  With rsCaixa
    'Verifica se há informações sobre o Caixa solicitado
    If .RecordCount = 0 Then
      'É criado seu registro inicial (1ª ocorrência do Caixa)
      .AddNew
      !Filial = gnCodFilial
      !Data = Data_Atual
      !Caixa = nCaixa
      !Ordem = 1
      !Funcionário = 0
      !Descrição = "Início do dia"
      !Hora = Format(CStr(Time), "hh:mm:ss")
      .Update
      'Posição inicial
      nOrdem = 1
      'Saldo Anterior
      dblSaldoAnterior = 0
    Else
      'Verifica se há informações sobre o Caixa no dia solicitado
      .FindLast "Data <= #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
      If !Data <> Data_Atual Then
        'Realiza o início de dia (1ª ocorrência do dia)
        With tFinalizadora
          '24/06/2005 - Daniel
          '
          'Uso da função nativa do VB IIf para Tratamento evitando assim o erro 94 (Invalid use of Null)
          'Esta ocorrência foi registrada na empresa Barro Queimado
          .dblDinheiro = IIf(IsNumeric(rsCaixa![Total Dinheiro]), rsCaixa![Total Dinheiro], 0)
          .dblCheque = IIf(IsNumeric(rsCaixa![Total Cheques]), rsCaixa![Total Cheques], 0)
          .dblChequePre = IIf(IsNumeric(rsCaixa![Total Cheques Pré]), rsCaixa![Total Cheques Pré], 0)
          .dblCartao = IIf(IsNumeric(rsCaixa![Total Cartões]), rsCaixa![Total Cartões], 0)
          .dblVale = IIf(IsNumeric(rsCaixa![Total Vales]), rsCaixa![Total Vales], 0)
          'Parcelamento inicia com o valor igual a zero
          .dblParcelamento = 0
        End With
        'Saldo Anterior
        dblSaldoAnterior = IIf(IsNumeric(!Final), !Final, 0)
        .AddNew
        !Filial = gnCodFilial
        !Data = Data_Atual
        !Ordem = 1
        !Caixa = nCaixa
        !Hora = Format(CStr(Time), "hh:mm:ss")
        !Funcionário = nFuncionario
        !Descrição = "Início do dia"
        !Dinheiro = tFinalizadora.dblDinheiro
        ![Total Dinheiro] = tFinalizadora.dblDinheiro
        !Cheques = tFinalizadora.dblCheque
        ![Total Cheques] = tFinalizadora.dblCheque
        ![Cheques Pré] = tFinalizadora.dblChequePre
        ![Total Cheques Pré] = tFinalizadora.dblChequePre
        !Cartões = tFinalizadora.dblCartao
        ![Total Cartões] = tFinalizadora.dblCartao
        !Vales = tFinalizadora.dblVale
        ![Total Vales] = tFinalizadora.dblVale
        ![Saldo Anterior] = 0
        !Final = dblSaldoAnterior
        .Update
        'Posição inicial
        nOrdem = 1
      Else
        'Caixa com dia já iniciado, somente informa os valores atuais
        '28/10/2004 - Daniel
        'BUG: Tratamento para valores nulos
        With tFinalizadora
          .dblDinheiro = IIf(IsNumeric(rsCaixa.Fields("Total Dinheiro").Value), Format(rsCaixa![Total Dinheiro], FORMAT_VALUE), 0)               'rsCaixa![Total Dinheiro]
          .dblCheque = IIf(IsNumeric(rsCaixa.Fields("Total Cheques").Value), Format(rsCaixa![Total Cheques], FORMAT_VALUE), 0)                   'rsCaixa![Total Cheques]
          .dblChequePre = IIf(IsNumeric(rsCaixa.Fields("Total Cheques Pré").Value), Format(rsCaixa![Total Cheques Pré], FORMAT_VALUE), 0)        'rsCaixa![Total Cheques Pré]
          .dblCartao = IIf(IsNumeric(rsCaixa.Fields("Total Cartões").Value), Format(rsCaixa![Total Cartões], FORMAT_VALUE), 0)                   'rsCaixa![Total Cartões]
          .dblVale = IIf(IsNumeric(rsCaixa.Fields("Total Vales").Value), Format(rsCaixa![Total Vales], FORMAT_VALUE), 0)                         'rsCaixa![Total Vales]
          .dblParcelamento = IIf(IsNumeric(rsCaixa.Fields("Total Parcelamento").Value), Format(rsCaixa![Total Parcelamento], FORMAT_VALUE), 0)   'rsCaixa![Total Parcelamento]
        End With
        'Posição atual
        nOrdem = !Ordem
        'Saldo Anterior
        dblSaldoAnterior = !Final
      End If
      tValoresAtuais = tFinalizadora
    End If
    .Close
  End With
  Set rsCaixa = Nothing
  
  'Finaliza transação
  If blnTransaction Then ws.CommitTrans: blnInTransaction = False
  
  gbCheckOpenCaixa = True
  Exit Function

ErrHandler:
  If blnInTransaction Then ws.Rollback
  If blnByPassErr Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    MsgBox "Ocorreu o erro " & Err.Number & " - " & Err.Description & _
      vbCrLf & "Ao verificar a inicialização do Caixa [Início de dia].", vbCritical, "Erro"
  End If
  
End Function

'08/10/2003 - mpdea
'Ajusta os valores da conta de cliente caso exista Desconto no SubTotal
Private Sub AdjustContaCliente(ByVal bytFilial As Byte, ByVal lngSequencia As Long, ByVal dblTotal As Double, ByVal dblDescontoSubTotal As Double)
  Dim rstContaCliente As Recordset
  Dim strSQL As String
  Dim sngDescPerc As Single
  Dim sngNovoTotal As Single
  Dim sngDiferenca As Single
  Dim sngMaiorValor As Single
  
  
  If dblTotal = 0 Or dblDescontoSubTotal = 0 Then Exit Sub
  
  'Desconto percentual
  sngDescPerc = CSng(dblDescontoSubTotal / (dblTotal + dblDescontoSubTotal))
  
  'Tabela Conta Cliente
  strSQL = "SELECT * FROM [Conta Cliente] "
  strSQL = strSQL & "WHERE Filial = " & bytFilial
  strSQL = strSQL & " AND Sequência = " & lngSequencia
  
  'Total após desconto
  sngNovoTotal = 0
  
  Set rstContaCliente = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rstContaCliente
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        .Edit
        .Fields("Valor").Value = Format(.Fields("Valor").Value * (1 - sngDescPerc), FORMAT_VALUE)
        
        'Novo total
        sngNovoTotal = sngNovoTotal + .Fields("Valor").Value
        
        'Maior valor
        If .Fields("Valor").Value > sngMaiorValor Then
          sngMaiorValor = .Fields("Valor").Value
        End If
        
        .Update
        .MoveNext
      Loop
      
      'Verifica possivel diferença e aplica no item com maior valor
      sngDiferenca = CSng(dblTotal) - sngNovoTotal
      If sngDiferenca <> 0 Then
        .MoveFirst
        Do Until .EOF
          If .Fields("Valor").Value = sngMaiorValor Then
            .Edit
            .Fields("Valor").Value = Format(.Fields("Valor").Value + sngDiferenca, FORMAT_VALUE)
            .Update
            Exit Do
          End If
          .MoveNext
        Loop
      End If
      
    End If
    .Close
  End With
  Set rstContaCliente = Nothing
  
End Sub
Private Sub GeraAcertoConsignacaoEntrada(ByVal bytFilial As Byte, lngSequencia As Long, strCodigoProduto As String, dblQtde As Double, ByVal PrecoVenda As Double)
  Dim strSQL         As String
  Dim bytAuxi        As Byte
  Dim dblQuantidade  As Double
  Dim dblQtdeBaixar  As Double
  
  Dim rstEntraProd   As Recordset
  
  dblQuantidade = dblQtde
  
  'Verificar em qual [Entradas - Produtos] há este Produto
  strSQL = "SELECT * FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  strSQL = strSQL & " AND Código = '" & strCodigoProduto & "'"
  strSQL = strSQL & " AND EntradaConsignada " 'Realmente é uma consignação
  strSQL = strSQL & " AND NOT ConsignacaoFechada "
  'strSQL = strSQL & " AND NOT Selecionado "  o fato de ser selecionado ou não, não poderá implicar na criação de um novo registro na table de Acerto
  strSQL = strSQL & " AND NOT Acertado "      'Não foi ainda acertado 100%
  strSQL = strSQL & " ORDER BY Sequência "

  Set rstEntraProd = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstEntraProd
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do Until .EOF
        'Primeira Situação: Há de ficar Qtde Disponível
        If .Fields("QtdeAtual").Value >= dblQuantidade Then
          .Edit
          .Fields("QtdeAtual").Value = .Fields("QtdeAtual").Value - dblQuantidade
          .Update
          
          Call CriarAcerto(.Fields("Filial").Value, .Fields("Sequência").Value, Data_Atual, .Fields("Linha").Value, strCodigoProduto, dblQuantidade, bytFilial, lngSequencia, .Fields("Preço").Value, PrecoVenda)
          
          Exit Do
        End If
        'Segunda Situação: Há de zerar uma e baixar da outra
        If .Fields("QtdeAtual").Value < dblQuantidade Then
          
          dblQuantidade = dblQuantidade - .Fields("QtdeAtual").Value
          dblQtdeBaixar = .Fields("QtdeAtual").Value
          
          .Edit
          .Fields("QtdeAtual").Value = 0
          .Update
          
          Call CriarAcerto(.Fields("Filial").Value, .Fields("Sequência").Value, Data_Atual, .Fields("Linha").Value, strCodigoProduto, dblQuantidade, bytFilial, lngSequencia, .Fields("Preço").Value, PrecoVenda)
        
        End If
      
      .MoveNext
      Loop

    End If
    .Close
  End With

  Set rstEntraProd = Nothing

End Sub

Private Sub CriarAcerto(ByVal Filial As Byte, ByVal sequencia As Long, ByVal DataAcerto As Date, ByVal LinhaProd As Byte, ByVal CodigoProduto As String, ByVal QtdeVendida As Double, ByVal FilialVenda As Byte, ByVal SequenciaVenda As Long, ByVal PrecoCusto As Double, ByVal PrecoVenda As Double)
  '14/09/2004 - Daniel
  'Case: Resultado
  'Criação de registros na tabela AcertoConsignacaoEntrada
  '14/10/2004 - Daniel
  'Adicionado os campos: Fornecedor, PrecoVenda, PrecoCusto
  Dim rstAcerto    As Recordset
  Dim rstEntraProd As Recordset
  Dim rstEntradas  As Recordset
  Dim strSQL       As String
  Dim blnFlag      As Boolean

  Set rstAcerto = db.OpenRecordset("AcertoConsignacaoEntrada", dbOpenDynaset)

  With rstAcerto
    .AddNew
    .Fields("Filial").Value = Filial
    .Fields("Sequencia").Value = sequencia
    .Fields("DataAcerto").Value = DataAcerto
    .Fields("LinhaProduto").Value = LinhaProd
    .Fields("CodigoProduto").Value = CodigoProduto
    .Fields("QtdeVendida").Value = QtdeVendida
    .Fields("FilialVenda").Value = FilialVenda
    .Fields("SequenciaVenda").Value = SequenciaVenda
    .Fields("PrecoCusto").Value = Format(PrecoCusto, FORMAT_VALUE)
    .Fields("PrecoVenda").Value = Format(PrecoVenda, FORMAT_VALUE)
    .Update
    .Close
  End With

  Set rstAcerto = Nothing

  '-----------------------------------------------------------------------
  ' Verificar [Entradas - Produtos] se QtdeAtual está zerada para podermos
  ' atualizar o campo [Entradas - Produtos].ConsignacaoFechada
  '-----------------------------------------------------------------------
  strSQL = "SELECT * FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequência = " & sequencia
  strSQL = strSQL & " AND Linha = " & LinhaProd

  Set rstEntraProd = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntraProd
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("QtdeAtual").Value = 0 Then
        .Edit
        .Fields("ConsignacaoFechada").Value = True
        .Update
      End If
      
    End If
    .Close
  End With
  
  Set rstEntraProd = Nothing

  '-----------------------------------------------------------------------
  ' Verificar todas as [Entradas - Produtos] da Entrada se o campo
  ' [Entradas - Produtos].ConsignacaoFechada está True em todas as
  ' [Entradas - Produtos] atualizaremos Entradas.ConsignacaoFechada
  '-----------------------------------------------------------------------
  strSQL = ""
  strSQL = "SELECT [Entradas - Produtos].ConsignacaoFechada "
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Entradas.Filial = " & Filial
  strSQL = strSQL & " AND Entradas.Sequência = " & sequencia
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = Entradas.Sequência "

  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        blnFlag = .Fields("ConsignacaoFechada").Value
        
        If Not blnFlag Then Exit Do
        
      .MoveNext
      Loop
    
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing

  If blnFlag Then Call AtualizarEntradas(Filial, sequencia)

End Sub

Private Sub AtualizarEntradas(ByVal Filial As Byte, ByVal sequencia As Long)
  Dim rstEntradas As Recordset
  Dim strSQL      As String
  
  strSQL = "SELECT Entradas.ConsignacaoFechada "
  strSQL = strSQL & " FROM Entradas "
  strSQL = strSQL & " WHERE Entradas.Filial = " & Filial
  strSQL = strSQL & " AND Entradas.Sequência = " & sequencia
  
  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      .Edit
      .Fields("ConsignacaoFechada").Value = True
      .Update
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing
  
End Sub
Private Sub ReduzirComissao(ByVal Tabela As String, ByVal Produto As String, ByRef dblValorDoCadastroProduto As Double)
  '11/02/2005 - Daniel
  'Problema levantado pela Daring
  'Se for dado desconto só em um produto em uma nota com X produtos, a comissão
  'do vendedor em todos os ítens da nota estava caindo pela metade sendo que o
  'correto é reduzir só a do produto que teve alteração de preço ou desconto
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

Private Sub BuscarRetencao(ByVal intCodigo As Integer, ByRef dblRetencao As Double)
  '22/03/2005 - Daniel
  '
  'Case.....: Bem Me Quer
  'Projeto..: Retenção sobre comissões
  Dim rstRetencao As Recordset

  dblRetencao = 0
  
  Set rstRetencao = db.OpenRecordset("SELECT ValorRetencao FROM Retencao WHERE Código = " & intCodigo, dbOpenSnapshot)
  
  With rstRetencao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      If IsNumeric(.Fields("ValorRetencao").Value) Then
        dblRetencao = Format(.Fields("ValorRetencao").Value, FORMAT_VALUE)
      Else
        dblRetencao = 0
      End If
    End If
    .Close
  End With

  Set rstRetencao = Nothing

End Sub
