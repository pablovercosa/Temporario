Attribute VB_Name = "modEfetivaSaida"
Option Explicit

'20/12/2005 - mpdea
'Quick Fiscal
'Vari�veis para impress�o de dados em cupom de parcelamento
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

Public Function Desefetiva_Sa�da(ByVal Filial As Integer, ByVal Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '07/08/2002 - mpdea
  'Inserido os recordsets que estavam a n�vel modular sem necessidade,
  'ocupando mais mem�ria
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
'  Dim rsParametros As Recordset
  Dim rsOp_Sa�da As Recordset
'  Dim rsContas_Receber As Recordset
  Dim rsResumo_Di�rio As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
'  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
'  Dim rsEstoque_Final As Recordset
'  Dim rsPre�os As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
'  Dim rsGrade As Recordset
  Dim rsSaidas As Recordset
  Dim rsSaidas_Prod As Recordset
'  Dim rsSaidas_Serv As Recordset
  Dim rsSa�da_Cheques As Recordset
  Dim rsSa�da_Parcelas As Recordset
  Dim rsComiss�es As Recordset
  Dim rsComiss�es_Serv As Recordset
'  Dim rsFuncionarios As Recordset
'  Dim rsTabelas As Recordset
  Dim rsConta_Cli As Recordset
'  Dim rsCartoes As Recordset
'  Dim rsBancos As Recordset
'  Dim rsEdicoes As Recordset
'  Dim rsServicos As Recordset
    
  '11/12/2009 - Andrea
  Dim rsSa�da_Cartoes As Recordset
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
 Dim Aux_Prod As String
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim Edi��o As Long
 Dim Aux_Tipo As Integer
 Dim Aux_Erro As Integer
 Dim Estoque_Final As Double
 Dim Criar_Registro As Integer
 Dim Mensagem As String
 Dim Saldo_Conta As Double
 
 Dim Aux_Val_Produto As Double
 Dim Aux_Val_Servi�o As Double
 
 Dim Val_Cheques As Double
 Dim Val_Cheques_Pr� As Double
 Dim Comiss�o As Double
 
  'Vari�vel de Tratamento de Erro
  Dim intRepeatUpdateLocked As Integer
 
  Dim strSQL As String
 
 On Error GoTo Processa_Erro

  Screen.MousePointer = vbHourglass
  
 Set rsSaidas = db.OpenRecordset("Sa�das")
' Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
' Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
 Set rsOp_Sa�da = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
 Set rsResumo_Di�rio = db.OpenRecordset("Resumo Di�rio")
 Set rsEtiquetas = db.OpenRecordset("Etiquetas")
 Set rsCaixa = db.OpenRecordset("Caixa")
 Set rsRes_Financeiro = db.OpenRecordset("Resumo Di�rio Financeiro")
' Set rsEstoque = db.OpenRecordset("Estoque")
' Set rsPre�os = db.OpenRecordset("Pre�os", , dbReadOnly)
 Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
 Set rsEmprestimos = db.OpenRecordset("Consigna��o Sa�da")
' Set rsGrade = db.OpenRecordset("C�digos da Grade")
 Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos", , dbReadOnly)
 Set rsSa�da_Cheques = db.OpenRecordset("Movimento - Cheques", , dbReadOnly)
 Set rsSa�da_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
 Set rsComiss�es = db.OpenRecordset("Comiss�o")
 Set rsComiss�es_Serv = db.OpenRecordset("Comiss�o Servi�os")
 Set rsConta_Cli = db.OpenRecordset("Conta Cliente")
' Set rsCartoes = db.OpenRecordset("Cart�es", , dbReadOnly)
' Set rsBancos = db.OpenRecordset("Bancos", , dbReadOnly)

 '11/12/2009 - Andrea
 Set rsSa�da_Cartoes = db.OpenRecordset("Movimento - Cartoes", , dbReadOnly)
  
  Screen.MousePointer = vbDefault
  
 rsSaidas.Index = "Sequ�ncia"
 rsSaidas.Seek "=", Filial, Mov
 If rsSaidas.NoMatch Then
   Desefetiva_Sa�da = 1
   Exit Function
 End If
 
 
 Rem Encontra a tabela de opera��es
 rsOp_Sa�da.Index = "C�digo"
 rsOp_Sa�da.Seek "=", rsSaidas("Opera��o")
 If rsOp_Sa�da.NoMatch Then
    Desefetiva_Sa�da = 2
    Exit Function
 End If


  Screen.MousePointer = vbHourglass
  
 Rem Atualiza Caixa, se for o caso
 'frmEntradas.Percent.Value = 4
 If rsOp_Sa�da("Dinheiro") = True Then
 ' If rsSaidas("Recebe - Dinheiro") <> 0 Or rsSaidas("Recebe - Cart�o") <> 0 Or rsSaidas("Recebe - Vale") <> 0 Then
    Erro = False
    Caixa_Novo = False
    Ordem = 0
      
    rsCaixa.Index = "Data"
    rsCaixa.Seek "<", Filial, rsSaidas("Caixa"), CDate(rsSaidas("Data")), 9999
    If rsCaixa.NoMatch Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Filial") <> Filial Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Data") <> rsSaidas("Data") Then Caixa_Novo = True
    If Caixa_Novo = False Then If rsCaixa("Caixa") <> rsSaidas("Caixa") Then Caixa_Novo = True
   
    If Caixa_Novo = True Then 'Come�a o Caixa do dia
       Desefetiva_Sa�da = 55
       Exit Function
    End If

    
    Rem Neste ponto tem o �ltimo caixa no buffer
    Rem Acha cheques
    Val_Cheques = 0
    rsSa�da_Cheques.Index = "Ordem"
    Ordem = 0
    Erro = False
    Do
     rsSa�da_Cheques.Seek ">", Filial, Mov, Ordem
     If rsSa�da_Cheques.NoMatch Then Erro = True
     If Erro = False Then If rsSa�da_Cheques("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSa�da_Cheques("Sequ�ncia") <> Mov Then Erro = True

     If Erro = False Then
       If rsSa�da_Cheques("Bom") = rsSaidas("Data") Then
         Val_Cheques = Val_Cheques + rsSa�da_Cheques("Valor")
       End If
       If rsSa�da_Cheques("Bom") <> rsSaidas("Data") Then
         Val_Cheques_Pr� = Val_Cheques_Pr� + rsSa�da_Cheques("Valor")
       End If
       Ordem = rsSa�da_Cheques("Ordem")
     End If
    Loop Until Erro = True

        
    Ordem = rsCaixa("Ordem")
    Ordem = Ordem + 1
    Saldo_Ant = rsCaixa("Final")
    Tot_Dinheiro = rsCaixa("Total Dinheiro")
    Tot_Cheques = rsCaixa("Total Cheques")
    Tot_Cheques_Pre = rsCaixa("Total Cheques Pr�")
    Tot_Cart�es = rsCaixa("Total Cart�es")
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
       rsCaixa("Descri��o") = "Cancelada Sa�da n�mero " & str(Mov)
       rsCaixa("Saldo Anterior") = Saldo_Ant
       'rsCaixa("Total Cheques Pr�") = Tot_Cheques_Pre
       rsCaixa("Cart�es") = -rsSaidas("Recebe - Cart�o")
       rsCaixa("Total Cart�es") = Tot_Cart�es - rsSaidas("Recebe - Cart�o")
       rsCaixa("Vales") = -rsSaidas("Recebe - Vale")
       rsCaixa("Total Vales") = Tot_Vales - rsSaidas("Recebe - Vale")
       rsCaixa("Cheques") = -Val_Cheques
       rsCaixa("Total Cheques") = Tot_Cheques - Val_Cheques
       rsCaixa("Cheques Pr�") = -Val_Cheques_Pr�
       rsCaixa("Total Cheques Pr�") = Tot_Cheques_Pre - Val_Cheques_Pr�
       rsCaixa("Parcelamento") = -rsSaidas("Total Prazo")
       rsCaixa("Total Parcelamento") = Tot_Parcela - rsSaidas("Total Prazo")
       rsCaixa("Dinheiro") = -rsSaidas("Recebe - Dinheiro")
       rsCaixa("Total Dinheiro") = Tot_Dinheiro - rsSaidas("Recebe - Dinheiro")
       rsCaixa("Final") = Tot_Dinheiro - rsSaidas("Recebe - Cart�o") - rsSaidas("Recebe - Vale") - rsSaidas("Recebe - Dinheiro") - Val_Cheques - Val_Cheques_Pr� + Tot_Cart�es + Tot_Vales + Tot_Cheques + Tot_Cheques_Pre
     rsCaixa.Update
 ' End If
 End If
 
 
  '---------------------------------------------------------------------------------
  '20/05/2002 - mpdea
  '
  'Otimizado a exclus�o dos registros da tabela de Contas a Receber
  '---------------------------------------------------------------------------------
 
 
' Rem Apagar Lan�amentos em Controle de Cheques
' rsContas_Receber.Index = "Contas"
' Erro = False
'Lp1_Cheque1:
' rsContas_Receber.Seek ">", "C", Filial, Mov, 0
' If rsContas_Receber.NoMatch Then Erro = True
' If Erro = False Then If rsContas_Receber("Sequ�ncia") <> Mov Then Erro = True
' If Erro = False Then If rsContas_Receber("Filial") <> Filial Then Erro = True
' If Erro = False Then If rsContas_Receber("Tipo") <> "C" Then Erro = True
' If Erro = False Then
'   rsContas_Receber.Delete
'   GoTo Lp1_Cheque1
' End If
'
'
'
' Rem Apagar Lan�amentos em controle de cart�es, se for o caso
' rsContas_Receber.Index = "Contas"
' Erro = False
' Do While True
'    rsContas_Receber.Seek ">", "O", Filial, Mov, 0
'    If rsContas_Receber.NoMatch Then Erro = True
'    If Erro = False Then If rsContas_Receber("Sequ�ncia") <> Mov Then Erro = True
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
' If Erro = False Then If rsContas_Receber("Sequ�ncia") <> Mov Then Erro = True
' If Erro = False Then If rsContas_Receber("Filial") <> Filial Then Erro = True
' If Erro = False Then If rsContas_Receber("Tipo") <> "R" Then Erro = True
' If Erro = False Then
'   rsContas_Receber.Delete
'   GoTo Lp1_Receber1
' End If
  
  
  db.Execute "DELETE * FROM [Contas a Receber] WHERE Filial = " & Filial & _
    " AND Sequ�ncia = " & Mov, dbFailOnError
  '10/09/2007 - Anderson
  'Gera arquivo log do sistema
  If g_bolSystemLog Then
    SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
    "DELETE * FROM [Contas a Receber] WHERE Filial = " & Filial & " AND Sequ�ncia = " & Mov, _
    "modEfetivaSaida_Desefetiva_Sa�da", _
    "Contas a Receber", g_strArquivoSystemLog
  End If
  
  '---------------------------------------------------------------------------------
 
 
 Aux_Val_Produto = CDbl(rsSaidas("Total")) - CDbl(rsSaidas("Servi�os"))
 Aux_Val_Servi�o = CDbl(rsSaidas("Servi�os"))

 Rem Atualiza Resumo Di�rio
 If rsOp_Sa�da("Tipo") <> "O" Then
   rsResumo_Di�rio.Index = "Data"
   rsResumo_Di�rio.Seek "=", Filial, rsSaidas("Data")
   If rsResumo_Di�rio.NoMatch Then
     rsResumo_Di�rio.AddNew
     rsResumo_Di�rio("Filial") = Filial
     rsResumo_Di�rio("Data") = rsSaidas("Data")
   Else
     rsResumo_Di�rio.Edit
   End If
   
   If rsOp_Sa�da("Tipo") = "V" Then
       rsResumo_Di�rio("Valor Vendas") = CDbl(rsResumo_Di�rio("Valor Vendas")) - Aux_Val_Produto
       rsResumo_Di�rio("Valor Servi�os") = CDbl(rsResumo_Di�rio("Valor Servi�os")) - Aux_Val_Servi�o
   End If
   
   If rsOp_Sa�da("Tipo") = "T" Then rsResumo_Di�rio("Valor T Sa�da") = CDbl(rsResumo_Di�rio("Valor T Sa�da")) - CDbl(rsSaidas("Total"))
   If rsOp_Sa�da("Tipo") = "A" Then rsResumo_Di�rio("Valor A Sa�da") = CDbl(rsResumo_Di�rio("Valor A Sa�da")) - CDbl(rsSaidas("Total"))
   If rsOp_Sa�da("Tipo") = "G" Then rsResumo_Di�rio("Valor G Sa�da") = CDbl(rsResumo_Di�rio("Valor G Sa�da")) - CDbl(rsSaidas("Total"))
   If rsOp_Sa�da("Tipo") = "E" Then rsResumo_Di�rio("Valor E Sa�da") = CDbl(rsResumo_Di�rio("Valor E Sa�da")) - CDbl(rsSaidas("Total"))

  rsResumo_Di�rio.Update
End If

 
  Rem Atualiza Resumo Di�rio Financeiro
  If rsOp_Sa�da("Dinheiro") = True Then
    rsRes_Financeiro.Index = "Data"
    rsRes_Financeiro.Seek "=", Filial, rsSaidas("Data")
    If rsRes_Financeiro.NoMatch Then
       rsRes_Financeiro.AddNew
       rsRes_Financeiro("Filial") = Filial
       rsRes_Financeiro("Data") = rsSaidas("Data")
    Else
       rsRes_Financeiro.Edit
    End If
    
    If rsOp_Sa�da("Tipo") = "V" Then
       rsRes_Financeiro("Valor Vendas") = CDbl(rsRes_Financeiro("Valor Vendas")) - Aux_Val_Produto
       rsRes_Financeiro("Valor Servi�os") = CDbl(rsRes_Financeiro("Valor Servi�os")) - Aux_Val_Servi�o
    End If
    If rsOp_Sa�da("Tipo") = "T" Then rsRes_Financeiro("Valor T Sa�da") = CDbl(rsRes_Financeiro("Valor T Sa�da")) - CDbl(rsSaidas("Total"))
    If rsOp_Sa�da("Tipo") = "A" Then rsRes_Financeiro("Valor A Sa�da") = CDbl(rsRes_Financeiro("Valor A Sa�da")) - CDbl(rsSaidas("Total"))
    If rsOp_Sa�da("Tipo") = "G" Then rsRes_Financeiro("Valor G Sa�da") = CDbl(rsRes_Financeiro("Valor G Sa�da")) - CDbl(rsSaidas("Total"))
    If rsOp_Sa�da("Tipo") = "E" Then rsRes_Financeiro("Valor E Sa�da") = CDbl(rsRes_Financeiro("Valor E Sa�da")) - CDbl(rsSaidas("Total"))


    rsRes_Financeiro.Update
  End If


  Rem Apaga conta do cliente
  rsConta_Cli.Index = "Sequ�ncia"
  Erro = False
Lp1_Conta_Cli1:
  rsConta_Cli.Seek ">", Filial, Mov, 0
  If rsConta_Cli.NoMatch Then Erro = True
  If Erro = False Then If rsConta_Cli("Sequ�ncia") <> Mov Then Erro = True
  If Erro = False Then If rsConta_Cli("Filial") <> Filial Then Erro = True
  If Erro = False Then
    rsConta_Cli.Delete
    GoTo Lp1_Conta_Cli1
  End If

  Rem Apaga comiss�es
  rsComiss�es.Index = "Sequ�ncia"
  Erro = False
Lp1_Comiss�o1:
  rsComiss�es.Seek ">", Filial, Mov, 0
  If rsComiss�es.NoMatch Then Erro = True
  If Erro = False Then If rsComiss�es("Sequ�ncia") <> Mov Then Erro = True
  If Erro = False Then If rsComiss�es("Filial") <> Filial Then Erro = True
  If Erro = False Then
    rsComiss�es.Delete
    GoTo Lp1_Comiss�o1
  End If


  Rem Apaga Resumo Clientes
  rsResumo_Clientes.Index = "Sequ�ncia"
  Erro = False
Lp1_ResumoCli1:
  rsResumo_Clientes.Seek ">=", Filial, Mov
  If rsResumo_Clientes.NoMatch Then Erro = True
  If Erro = False Then If rsResumo_Clientes("Sequ�ncia") <> Mov Then Erro = True
  If Erro = False Then If rsResumo_Clientes("Filial") <> Filial Then Erro = True
  If Erro = False Then
    rsResumo_Clientes.Delete
    GoTo Lp1_ResumoCli1
  End If






  Rem Loop dos Produtos
  rsSaidas_Prod.Index = "Sequ�ncia"
  Ordem = 0
  rsProdutos.Index = "C�digo"
Prox_Prod:
  rsSaidas_Prod.Seek ">", Filial, Mov, Ordem
  
  If rsSaidas_Prod.NoMatch Then GoTo Ve_Serv
  If rsSaidas_Prod("Filial") <> Filial Then GoTo Ve_Serv
  If rsSaidas_Prod("Sequ�ncia") <> Mov Then GoTo Ve_Serv
  
  Ordem = rsSaidas_Prod("Linha")
  'Verifica se tem grade
  C�d = rsSaidas_Prod("C�digo")
  Tamanho = 0
  Cor = 0
  Aux_Prod = C�d
  
  Acha_Produto Aux_Prod, C�d, Tamanho, Cor, Edi��o, Aux_Tipo, Aux_Erro
  If Aux_Erro <> 0 Then
    GoTo Prox_Prod
  End If
   
  C�d = UCase(C�d)
   
  'Neste ponto C�D tem o c�digo do produto
  'Tamanho e Cor cont�m os respectivos dados
  'Agora grava arquivo do estoque
  rsProdutos.Seek "=", C�d
  
  Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))

  Rem  Ajusta Estoque
  If rsOp_Sa�da("Estoque") = True Then
  
'-------------------------------------------------------------------------------------
    '16/11/2003 - mpdea
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
             " AND Data = #" & Format(rsSaidas("Data"), "mm/dd/yyyy") & "#"
            
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .Edit
      Else
        .AddNew
        .Fields("Filial").Value = Filial
        .Fields("Data").Value = rsSaidas("Data").Value
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


'    Rem Encontra a posi��o do estoque
'    Criar_Registro = False
'    Estoque_Final = 0
'    rsEstoque.Index = "Produto"
'    rsEstoque.Seek "=", Filial, rsSaidas("Data"), C�d, Tamanho, Cor, Edi��o
'
'    If Not rsEstoque.NoMatch Then Estoque_Final = rsEstoque("Estoque Final")
'    If rsEstoque.NoMatch Then
'       rsEstoque.Index = "Data"
'       rsEstoque.Seek "<", Filial, C�d, Tamanho, Cor, Edi��o, rsSaidas("Data")
'       If rsEstoque.NoMatch Then Criar_Registro = True
'       If Not rsEstoque.NoMatch Then
'          If rsEstoque("Filial") = Filial And rsEstoque("Produto") = C�d And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edi��o") = Edi��o Then
'             Criar_Registro = True
'             Estoque_Final = rsEstoque("Estoque Final")
'           End If
'       End If
'
'       rsEstoque.AddNew
'       rsEstoque("Filial") = Filial
'       rsEstoque("Data") = rsSaidas("Data")
'       rsEstoque("Produto") = C�d
'       rsEstoque("Tamanho") = Tamanho
'       rsEstoque("Cor") = Cor
'       rsEstoque("Edi��o") = Edi��o
'       rsEstoque("Classe") = rsProdutos("Classe")
'       rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
'       rsEstoque("Estoque Anterior") = Estoque_Final
'       rsEstoque.Update
'
'       rsEstoque.Index = "Produto"
'       rsEstoque.Seek "=", Filial, rsSaidas("Data"), C�d, Tamanho, Cor, Edi��o
'      End If

'-------------------------------------------------------------------------------------

      Rem neste ponto esta com o registro de estoque
      Rem no buffer, agora soma com os valores da movimenta��o
'      rsEstoque.Edit
      If rsOp_Sa�da("Tipo") = "V" Then
          rsEstoque("Vendas") = rsEstoque("Vendas") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Vendas") = Format(rsEstoque("Valor Vendas") - rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
            
      If rsOp_Sa�da("Tipo") = "T" Then
          rsEstoque("Transf Sa�da") = rsEstoque("Transf Sa�da") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor T Sa�da") = Format(rsEstoque("Valor T Sa�da") - rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
      
      If rsOp_Sa�da("Tipo") = "A" Then
          rsEstoque("Ajuste Sa�da") = rsEstoque("Ajuste Sa�da") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Ajuste Sa�da") = Format(rsEstoque("Valor Ajuste Sa�da") - rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
          
      If rsOp_Sa�da("Tipo") = "G" Then
          rsEstoque("Gr�tis Sa�da") = rsEstoque("Gr�tis Sa�da") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Gr�tis Sa�da") = Format(rsEstoque("Valor Gr�tis Sa�da") - rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
      
      If rsOp_Sa�da("Tipo") = "E" Then
          rsEstoque("Empre Sa�da") = rsEstoque("Empre Sa�da") - rsSaidas_Prod("Qtde")
          rsEstoque("Valor Empre Sa�da") = Format(rsEstoque("Valor Empre Sa�da") - rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If

      Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
      Estoque_Final = Estoque_Final + rsEstoque("Transf Sa�da") + rsEstoque("Transf Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Ajuste Sa�da") + rsEstoque("Ajuste Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Gr�tis Sa�da") + rsEstoque("Gr�tis Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Empre Sa�da") + rsEstoque("Empre Entra")
      
      '08/08/2003 - maikel
      '             Descomentada a soma da coluna Devolu��o para resolver o problema de estoque
      Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolu��o")

      If rsProdutos("Estoque") = False Then
        Estoque_Final = 0
      End If

      rsEstoque("Estoque Final") = Estoque_Final
      rsEstoque.Update
      rsEstoque.Close
      
      Grava_Estoque_Final Filial, C�d, Tamanho, Cor, Edi��o, CSng(Estoque_Final), Date
      
  End If

  

  Rem Apaga etiquetas
  If rsSaidas_Prod("Etiqueta") = True Then
    rsEtiquetas.Index = "Funcion�rio"
    rsEtiquetas.Seek "=", rsSaidas("Digitador"), C�d, Tamanho, Cor
    If rsEtiquetas.NoMatch Then
       rsEtiquetas.AddNew
    Else
       rsEtiquetas.Edit
    End If
    rsEtiquetas("Funcion�rio") = rsSaidas("Digitador")
    rsEtiquetas("Produto") = C�d
    rsEtiquetas("Tamanho") = Tamanho
    rsEtiquetas("Cor") = Cor
    rsEtiquetas("Qtde") = rsEtiquetas("Qtde") - rsSaidas_Prod("Qtde")
    rsEtiquetas("Sequ�ncia") = Mov
    rsEtiquetas.Update
  End If





  Rem Atualiza arquivo de Empr�stimos
  If rsOp_Sa�da("Tipo") = "E" Then
     rsEmprestimos.Index = "Cliente"
Lp_Emp1:
     rsEmprestimos.Seek ">", rsSaidas("Filial"), rsSaidas("Sequ�ncia"), rsSaidas("Cliente"), 0, 0, 0, 0, 0
     If Not rsEmprestimos.NoMatch Then
       If rsEmprestimos("Filial") = rsSaidas("Filial") Then
         If rsEmprestimos("Sequ�ncia") = rsSaidas("Sequ�ncia") Then
           rsEmprestimos.Delete
           GoTo Lp_Emp1
         End If
       End If
     End If
  End If
     

  GoTo Prox_Prod
  
  
Ve_Serv:
  rsComiss�es_Serv.Index = "Sequ�ncia"
  Ordem = 0
Prox_Serv:
  rsComiss�es_Serv.Seek ">", Filial, Mov, Ordem
    
  If rsComiss�es_Serv.NoMatch Then GoTo Fim_Desefetiva
  If rsComiss�es_Serv("Filial") <> Filial Then GoTo Fim_Desefetiva
  If rsComiss�es_Serv("Sequ�ncia") <> Mov Then GoTo Fim_Desefetiva
  
  rsComiss�es_Serv.Delete
  
  GoTo Prox_Serv
  
  
Fim_Desefetiva:
   Desefetiva_Sa�da = 0
  
  '---------------------------------------------------------------------------------
  '20/05/2002 - mpdea
  '
  'Inclu�do o fechamento dos recordsets abertos e suas desassocia��es
  '---------------------------------------------------------------------------------
  
  rsSaidas.Close
  rsProdutos.Close
  rsOp_Sa�da.Close
  rsResumo_Di�rio.Close
  rsEtiquetas.Close
  rsCaixa.Close
  rsRes_Financeiro.Close
'  If Not rsEstoque Is Nothing Then rsEstoque.Close
  rsResumo_Clientes.Close
  rsEmprestimos.Close
  rsSaidas_Prod.Close
  rsSa�da_Cheques.Close
  rsSa�da_Parcelas.Close
  rsComiss�es.Close
  rsComiss�es_Serv.Close
  rsConta_Cli.Close
  
  '11/12/2009 - Andrea
  rsSa�da_Cartoes.Close
   
  Set rsSaidas = Nothing
  Set rsProdutos = Nothing
  Set rsOp_Sa�da = Nothing
  Set rsResumo_Di�rio = Nothing
  Set rsEtiquetas = Nothing
  Set rsCaixa = Nothing
  Set rsRes_Financeiro = Nothing
  Set rsEstoque = Nothing
  Set rsResumo_Clientes = Nothing
  Set rsEmprestimos = Nothing
  Set rsSaidas_Prod = Nothing
  Set rsSa�da_Cheques = Nothing
  Set rsSa�da_Parcelas = Nothing
  Set rsComiss�es = Nothing
  Set rsComiss�es_Serv = Nothing
  Set rsConta_Cli = Nothing
  Set rsSa�da_Cartoes = Nothing
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
          Desefetiva_Sa�da = -1 'A��o cancelada
          Exit Function
        End If
        
'        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
'          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Desefetiva Sa�da") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        Else
'          Desefetiva_Sa�da = -1 'A��o cancelada
'          Exit Function
'        End If
      End If
    Case Else
      'Outros Erros
      Select Case frmErro.gnShowErr(Err.Number, "Desefetiva Sa�da")
        Case 0 'Repetir
          Resume
        Case 1 'Prosseguir
          Resume Next
        Case 2 'Sair
          Desefetiva_Sa�da = -1 'A��o cancelada
          Exit Function
        Case 3 'Encerrar
          End
      End Select
  End Select

End Function

Public Function Efetiva_Sa�da(ByVal Filial As Integer, ByVal Mov As Long) As Integer
  '---------------------------------------------------------------------------------
  '07/08/2002 - mpdea
  'Inserido os recordsets que estavam a n�vel modular sem necessidade,
  'ocupando mais mem�ria
  '---------------------------------------------------------------------------------
  Dim rsProdutos As Recordset
  Dim rsParametros As Recordset
  Dim rsOp_Sa�da As Recordset
  Dim rsContas_Receber As Recordset
  Dim rsResumo_Di�rio As Recordset
  Dim rsEtiquetas As Recordset
  Dim rsCaixa As Recordset
  Dim rsRes_Financeiro As Recordset
  Dim rsResumo As Recordset
  Dim rsEstoque As Recordset
'  Dim rsEstoque_Final As Recordset
  Dim rsPre�os As Recordset
  Dim rsResumo_Clientes As Recordset
  Dim rsEmprestimos As Recordset
  Dim rsCliFor As Recordset
  Dim rsGrade As Recordset
  Dim rsSaidas As Recordset
  Dim rsSaidas_Prod As Recordset
  Dim rsSaidas_Serv As Recordset
  Dim rsSa�da_Cheques As Recordset
  Dim rsSa�da_Parcelas As Recordset
  Dim rsComiss�es As Recordset
  Dim rsComiss�es_Serv As Recordset
  Dim rsFuncionarios As Recordset
  Dim rsTabelas As Recordset
  Dim rsConta_Cli As Recordset
  Dim rsCartoes As Recordset
  Dim rsBancos As Recordset
  Dim rsEdicoes As Recordset
  Dim rsServicos As Recordset
  '10/12/2009 - Andrea
  Dim rsSa�da_Cartoes As Recordset
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
 Dim Tot_Cart�es As Double
 Dim Tot_Vales As Double
 Dim Tot_Parcelas As Double
 Dim C�d As String
 Dim C�d_Serv As Integer
 Dim Tamanho As Integer
 Dim Cor As Integer
 Dim Edi��o As Long
 Dim Estoque_Final As Double
 Dim Criar_Registro As Integer
 Dim Mensagem As String
 Dim Saldo_Conta As Double
 Dim Val_Cheques As Double
 Dim Val_Cheques_Pr� As Double
 Dim Comiss�o As Double
 Dim Aux_C�digo As String
 Dim Aux_Tipo As Integer
 Dim Aux_Erro As Integer
 Dim Saldo_Emp As Long
 Dim Ordem_Emp As Long
 Dim Comiss_T�cnico As Single
 Dim Val_Vista As Double
 Dim Emp_Existe As Boolean
 Dim sDescrAdicional As String
 
 Dim strAuxiliar As String
 
  
 '10/12/2009 - Andrea
 Dim Val_Cartoes As Double
  
  
  'Vari�vel de Tratamento de Erro
  Dim blnCaseCaixa As Boolean
  Dim intRepeatUpdateLocked As Integer
  Dim intRepeatUpdate3022 As Integer
  
  'Estrutura de formas de pagamentos
  Dim typTotalizadores As tpPaymentType
  
  'Vari�veis WEB
  Dim lngWEB_ID As Long
  Dim strStatusShopper As String
  Dim strStatusAdmin As String
  Dim strListPrice As String
  Dim intCodOpVenda As Integer
  Dim intCodOpReserva As Integer
  
  Dim strSQL As String
 
  '29/10/2002 - mpdea
  'C�digo do cliente
  Dim lngCodCliente As Long
  
  
  '28/09/2005 - mpdea
  'Nome do cliente
  Dim strNomeCliente As String
  
 
  Dim blnDiminuiComissao  As Boolean
  Dim rstTabelaPrecos     As Recordset
  Dim dblValorComissao    As Double
 
  '16/10/2004 - Daniel
  'Adicionada a var PrecoVenda que ser� tratada em GeraAcertoConsignacao
  'Case: Resultado
  Dim dblPrecoVenda As Double
  
  '03/07/2006 - mpdea
  'Comiss�o com reten��o
  'Case.....: Bem Me Quer
  'Projeto..: Reten��o sobre comiss�es
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
  
  Set rsSaidas = db.OpenRecordset("Sa�das")
  Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos", , dbReadOnly)
  Set rsSaidas_Serv = db.OpenRecordset("Sa�das - Servi�os", , dbReadOnly)
  Set rsSa�da_Cheques = db.OpenRecordset("Movimento - Cheques", , dbReadOnly)
  Set rsSa�da_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsSa�da_Cartoes = db.OpenRecordset("Movimento - Cartoes", , dbReadOnly)
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber")
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  Set rsOp_Sa�da = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
  Set rsResumo_Di�rio = db.OpenRecordset("Resumo Di�rio")
  Set rsEtiquetas = db.OpenRecordset("Etiquetas")
  Set rsRes_Financeiro = db.OpenRecordset("Resumo Di�rio Financeiro")
  ' Set rsEstoque = db.OpenRecordset("Estoque")
  ' Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
  Set rsPre�os = db.OpenRecordset("Pre�os", , dbReadOnly)
  Set rsResumo_Clientes = db.OpenRecordset("Resumo Clientes")
  Set rsEmprestimos = db.OpenRecordset("Consigna��o Sa�da")
  Set rsCliFor = db.OpenRecordset("Cli_For")
  Set rsGrade = db.OpenRecordset("C�digos da Grade")
  'Set rsComiss�es = db.OpenRecordset("Comiss�o")
  Set rsComiss�es_Serv = db.OpenRecordset("Comiss�o Servi�os")
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Pre�os", , dbReadOnly)
  Set rsConta_Cli = db.OpenRecordset("Conta Cliente")
  Set rsCartoes = db.OpenRecordset("Cart�es", , dbReadOnly)
  Set rsBancos = db.OpenRecordset("Bancos", , dbReadOnly)
  Set rsServicos = db.OpenRecordset("Servi�os", , dbReadOnly)

  Screen.MousePointer = vbDefault
   
  rsSaidas.Index = "Sequ�ncia"
  rsSaidas.Seek "=", Filial, Mov
  If rsSaidas.NoMatch Then
    Efetiva_Sa�da = 1
    Exit Function
  End If
  
  'Verifica se a sa�da j� foi efetivada
  'Check realizado devido a problemas de a fun��o estar sendo
  'chamada v�rias vezes em alguma circunst�ncia ainda n�o encontrada
  'mpdea 17/08/2000
  If rsSaidas("Efetivada") Then
    Efetiva_Sa�da = 0
    Exit Function
  End If
  
  '---------------------------------------------------------------------------------
  '12/07/2002 - mpdea
  '
  'lngWEB_ID > 0 = movimenta��o do tipo WEB
  'Inibido erro conforme marca��es
  '---------------------------------------------------------------------------------
  lngWEB_ID = CLng("0" & rsSaidas.Fields("WebOrderFormID").Value)
  
  
  Rem Encontra a tabela de opera��es
  rsOp_Sa�da.Index = "C�digo"
  rsOp_Sa�da.Seek "=", rsSaidas("Opera��o")
  If rsOp_Sa�da.NoMatch Then
    Efetiva_Sa�da = 2
    Exit Function
  End If
  
  Rem Encontra Cliente
  rsCliFor.Index = "C�digo"
  rsCliFor.Seek "=", rsSaidas("Cliente")
  If rsCliFor.NoMatch And lngWEB_ID = 0 Then '-> Inibe erro
    Efetiva_Sa�da = 3
    Exit Function
  Else
    '29/10/2002 - mpdea
    'C�digo do cliente
    lngCodCliente = rsCliFor.Fields("C�digo").Value
    
    '28/09/2005 - mpdea
    'Nome do cliente
    strNomeCliente = rsCliFor.Fields("Nome").Value & ""
    
    '20/12/2005 - mpdea
    'Quick Fiscal
    'Vari�veis para impress�o de dados em cupom de parcelamento
    'Bematech
    'CASE: Margarete Parizoto ME (QS71277-474)
    g_str_nome_cliente = strNomeCliente
    g_lng_nr_sequencia = Mov
  End If

  Rem Encontra Funcion�rio
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", rsSaidas("Digitador")
  If rsFuncionarios.NoMatch And lngWEB_ID = 0 Then '-> Inibe erro
    Efetiva_Sa�da = 4
    Exit Function
  End If

  '---------------------------------------------------------------------------------
  '07/05/2002 - mpdea
  '
  'Alterado verifica��o da exist�ncia da tabela de pre�os para opera��es do tipo
  'WEB (tabela de pre�os din�mica [DB:Pre�os])
  'Somente verifica se o campo WebOrderFormID = 0 (venda n�o WEB)
  '---------------------------------------------------------------------------------
  If lngWEB_ID = 0 Then
    'Encontra tabela de pre�os
    rsTabelas.Index = "Tabela"
    rsTabelas.Seek "=", rsSaidas("Tabela")
    If rsTabelas.NoMatch Then
      Efetiva_Sa�da = 5
      Exit Function
    End If
  End If
  '---------------------------------------------------------------------------------
 
 
  
  Screen.MousePointer = vbHourglass
 
' Call ws.BeginTrans

  '---[ Atualiza a data da movimenta��o para a data atual, devido a problemas com o estoque, financeiro, etc. ]---'
    With rsSaidas
      If (.Fields("Data") <> Data_Atual) Then
        MsgBox "Aten��o !" & vbCrLf & vbCrLf & "Essa movimenta��o foi gerada no dia " & .Fields("Data") & ". A data da movimenta��o est� sendo ajustada para " & Data_Atual, vbInformation, "Quick Store"
        
        .LockEdits = True
        .Edit
        .Fields("Data") = Data_Atual
        .Update
        .LockEdits = False
        
        .Index = "Sequ�ncia"
        .Seek "=", Filial, Mov
        If .NoMatch Then
          Efetiva_Sa�da = 1
          Exit Function
        End If
      End If
    End With
  '---[ Atualiza a data da movimenta��o para a data atual, devido a problemas com o estoque, financeiro, etc. ]---'

  '12/07/2002 - mpdea
  'Obt�m c�digos de opera��o de sa�da WEB
  Call GetWEBCod_Op(intCodOpReserva, intCodOpVenda, 0)
  
  '---------------------------------------------------------------------------------
  '07/05/2002 - mpdea
  '
  'Implementado exclus�o de tabelas tempor�rias para a venda do tipo WEB
  'e atualiza��o do pedido e hist�rico para Pagamento Recebido
  '
  '12/07/2002
  '
  'Somente executa quando a opera��o for de venda, conforme configura��o
  '
  '28/06/2005 - Daniel
  '
  'Existiram casos em que o usu�rio fazia o recebimento (efetiva��o da venda)
  'com outra opera��o distinta da original da venda virtual e ao tentar prosseguir,
  'no gerenciador de pedidos, o Quick dava a seguinte mensagem:
  '"Efetue o recebimento na tela de Sa�das para confirmar o Pagamento"
  'Ocorr�ncias: Os�rio (SEBO)
  'Antiga condi��o..: If lngWEB_ID <> 0 And rsSaidas.Fields("Opera��o").Value = intCodOpVenda Then
  'Nova condi��o....: If lngWEB_ID <> 0 And (rsSaidas.Fields("Opera��o").Value = intCodOpVenda Or rsOp_Sa�da.Fields("AlteraStatusPedidoWeb").Value) Then
  '---------------------------------------------------------------------------------
  If lngWEB_ID <> 0 And (rsSaidas.Fields("Opera��o").Value = intCodOpVenda Or rsOp_Sa�da.Fields("AlteraStatusPedidoWeb").Value) Then
    'Exclui tabelas tempor�rias
    strListPrice = Replace(LIST_PRICE_WEB, REPLACE_TQW, _
                           Format(lngWEB_ID, String(Len(REPLACE_TQW), "0")))
    Call db.Execute("DELETE * FROM Pre�os WHERE Tabela = '" & strListPrice & _
      "'", dbFailOnError)
    
    'Obt�m descri��o para o status de pagamento recebido
    Call GetDataDescPasso(ofsConfirmedPayment, strStatusShopper, strStatusAdmin)
    
    'Atualiza o Pedido
    Call db.Execute("UPDATE WEB_OrderForms SET " & _
      "StatusShopper = '" & strStatusShopper & _
      "', StatusAdmin = '" & strStatusAdmin & _
      "', Passo = " & ofsConfirmedPayment & " WHERE ID = " & lngWEB_ID, dbFailOnError)
    
    'Atualiza o Hist�rico do Pedido
    Call db.Execute("INSERT INTO WEB_OrderStatusHistoric " & _
      "(OrderFormID, Passo, StatusShopper, StatusAdmin, Data, WebSynchronize) " & _
      "VALUES (" & lngWEB_ID & ", " & ofsConfirmedPayment & ", '" & strStatusShopper & _
      "', '" & strStatusAdmin & "', #" & Format(Now, "MM/DD/YYYY HH:MM:SS") & _
      "#, True)", dbFailOnError)
  End If
  '---------------------------------------------------------------------------------

  '12/07/2002 - mpdea
  'Desvia em opera��o de reserva WEB
  If rsSaidas.Fields("Opera��o").Value <> intCodOpReserva Then
    If rsOp_Sa�da("Tipo") = "V" Then
      With rsCliFor
        .LockEdits = True
        .Edit
        .Fields("�ltima Compra").Value = rsSaidas.Fields("Data").Value
        .Fields("Data Altera��o").Value = Format(Date, "dd/mm/yyyy")
        .Update
        .LockEdits = False
      End With
    End If
  End If
  
 Rem Atualiza Caixa, se for o caso
 'frmEntradas.Percent.Value = 4
 If rsOp_Sa�da.Fields("Dinheiro").Value Then
 
' ' If rsSaidas("Recebe - Dinheiro") <> 0 Or rsSaidas("Recebe - Cart�o") <> 0 Or rsSaidas("Recebe - Vale") <> 0 Then
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
'    If Caixa_Novo = True Then 'Come�a o Caixa do dia
'       Erro = False
'       rsCaixa.Index = "Data"
'       rsCaixa.Seek "<", Filial, rsSaidas("Caixa"), rsSaidas("Data"), 0
'       If rsCaixa.NoMatch Then Erro = True
'       If Not Erro Then If rsCaixa("Filial") <> Filial Then Erro = True
'       If Not Erro Then If rsCaixa("Caixa") <> rsSaidas("Caixa") Then Erro = True
'       If Erro = True Then  'N�o existe dia anterior
'          rsCaixa.AddNew
'           rsCaixa("Filial") = Filial
'           rsCaixa("Caixa") = rsSaidas("Caixa")
'           rsCaixa("Data") = rsSaidas("Data")
'           rsCaixa("Hora") = Format(Time, "hh:mm:ss")
'           Ordem = 1
'           rsCaixa("Ordem") = Ordem
'           rsCaixa("Saldo Anterior") = 0
'           rsCaixa("Final") = 0
'           rsCaixa("Descri��o") = "In�cio do dia"
'          rsCaixa.Update
'       Else
'          Ordem = 1
'          Saldo_Ant = rsCaixa("Final")
'          Tot_Dinheiro = gsHandleNull(rsCaixa("Total Dinheiro"))
'          Tot_Cheques = gsHandleNull(rsCaixa("Total Cheques"))
'          Tot_Cheques_Pre = gsHandleNull(rsCaixa("Total Cheques Pr�"))
'          Tot_Cart�es = gsHandleNull(rsCaixa("Total Cart�es"))
'          Tot_Vales = gsHandleNull(rsCaixa("Total Vales"))
'          Tot_Parcelas = gsHandleNull(rsCaixa("Total Parcelamento"))
'
'          rsCaixa.AddNew
'            rsCaixa("Filial") = Filial
'            rsCaixa("Data") = rsSaidas("Data")
'            rsCaixa("Hora") = Format(Time, "hh:mm:ss")
'            rsCaixa("Caixa") = rsSaidas("Caixa")
'            rsCaixa("Ordem") = Ordem
'            rsCaixa("Funcion�rio") = rsSaidas("Operador")
'            rsCaixa("Descri��o") = "In�cio do dia"
'            rsCaixa("Saldo Anterior") = Saldo_Ant
'            rsCaixa("Dinheiro") = Tot_Dinheiro
'            rsCaixa("Cheques") = Tot_Cheques
'            rsCaixa("Cheques Pr�") = Tot_Cheques_Pre
'            rsCaixa("Cart�es") = Tot_Cart�es
'            rsCaixa("Vales") = Tot_Vales
'            rsCaixa("Total Dinheiro") = Tot_Dinheiro
'            rsCaixa("Total Cheques") = Tot_Cheques
'            rsCaixa("Total Cheques Pr�") = Tot_Cheques_Pre
'            rsCaixa("Total Cart�es") = Tot_Cart�es
'            rsCaixa("Total Vales") = Tot_Vales
'            rsCaixa("Total Parcelamento") = Tot_Parcelas
'            rsCaixa("Final") = Saldo_Ant
'          rsCaixa.Update
'      End If
'
'      rsCaixa.Index = "Caixa"
'      rsCaixa.Seek "<", Filial, rsSaidas("Data"), rsSaidas("Caixa"), 9999
'    End If

        
'    Rem Neste ponto tem o �ltimo caixa no buffer
'    Ordem = rsCaixa("Ordem")
'    Ordem = Ordem + 1
'    Saldo_Ant = rsCaixa("Final")
'    Tot_Dinheiro = rsCaixa("Total Dinheiro")
'    Tot_Cheques = rsCaixa("Total Cheques")
'    Tot_Cheques_Pre = rsCaixa("Total Cheques Pr�")
'    Tot_Cart�es = rsCaixa("Total Cart�es")
'    Tot_Vales = rsCaixa("Total Vales")
'    Tot_Parcelas = rsCaixa("Total Parcelamento")
    
     Rem Acha cheques
    Val_Cheques = 0
    rsSa�da_Cheques.Index = "Ordem"
    Ordem = 0
    Erro = False
    Do
     rsSa�da_Cheques.Seek ">", Filial, Mov, Ordem
     If rsSa�da_Cheques.NoMatch Then Erro = True
     If Erro = False Then If rsSa�da_Cheques("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSa�da_Cheques("Sequ�ncia") <> Mov Then Erro = True

     If Erro = False Then
       If rsSa�da_Cheques("Bom") = rsSaidas("Data") Then
         Val_Cheques = Val_Cheques + rsSa�da_Cheques("Valor")
       End If
       If rsSa�da_Cheques("Bom") <> rsSaidas("Data") Then
         Val_Cheques_Pr� = Val_Cheques_Pr� + rsSa�da_Cheques("Valor")
       End If
       Ordem = rsSa�da_Cheques("Ordem")
     End If
    Loop Until Erro = True


    'Acha pagamentos em cart�es feitos na tela de saidas
    Val_Cartoes = 0
    rsSa�da_Cartoes.Index = "Ordem"
    Ordem = 0
    Erro = False
    Do
     rsSa�da_Cartoes.Seek ">", Filial, Mov, Ordem
     If rsSa�da_Cartoes.NoMatch Then Erro = True
     If Erro = False Then If rsSa�da_Cartoes("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSa�da_Cartoes("Sequ�ncia") <> Mov Then Erro = True

     If Erro = False Then
       Val_Cartoes = rsSa�da_Cartoes("Valor")
       Ordem = rsSa�da_Cartoes("Ordem")
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
    'Verifica o in�cio do caixa, abertura do dia e retorna os �ltimos valores
    If Not gbCheckOpenCaixa(rsSaidas.Fields("Caixa").Value, _
      rsSaidas.Fields("Operador").Value, Saldo_Ant, Ordem, _
      typTotalizadores, False, True) Then
      'Ocorreu erro e a mensagem � exibida pela fun��o
      Efetiva_Sa�da = 9
      Exit Function
    Else
      Ordem = Ordem + 1
      Tot_Dinheiro = typTotalizadores.dblDinheiro
      Tot_Cheques = typTotalizadores.dblCheque
      Tot_Cheques_Pre = typTotalizadores.dblChequePre
      Tot_Cart�es = typTotalizadores.dblCartao
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
      .Fields("Funcion�rio").Value = rsSaidas("Operador")
      
      .Fields("Descri��o").Value = "Sa�da nr. " & Mov
      
      '28/01/2005 - Daniel
      'Cliente Taupys - ES solicitou que a Refer�ncia sa�sse na descri��o caso
      'estiver preenchida
      If Len(rsSaidas("Refer�ncia").Value) > 0 Then
        .Fields("Descri��o").Value = .Fields("Descri��o").Value & _
                                     " Ref. " & rsSaidas("Refer�ncia").Value
      End If
      
      '28/09/2005 - mpdea
      'Inclu�do o nome do cliente
      '
      '29/10/2002 - mpdea
      'Adicionado o c�digo do cliente na descri��o do registro de caixa
      If lngCodCliente > 0 Then
        .Fields("Descri��o").Value = Left(.Fields("Descri��o").Value & _
                                     " Cliente " & lngCodCliente & " - " & _
                                     strNomeCliente, .Fields("Descri��o").Size)
      End If
      
      .Fields("Saldo Anterior").Value = Saldo_Ant
      .Fields("Cart�es").Value = rsSaidas("Recebe - Cart�o")
      .Fields("Total Cart�es").Value = Tot_Cart�es + rsSaidas("Recebe - Cart�o")
      .Fields("Vales").Value = rsSaidas("Recebe - Vale")
      .Fields("Total Vales").Value = Tot_Vales + rsSaidas("Recebe - Vale")
      .Fields("Cheques").Value = Val_Cheques
      .Fields("Total Cheques").Value = Tot_Cheques + Val_Cheques
      .Fields("Cheques Pr�").Value = Val_Cheques_Pr�
      .Fields("Total Cheques Pr�").Value = Tot_Cheques_Pre + Val_Cheques_Pr�
      .Fields("Dinheiro").Value = rsSaidas("Recebe - Dinheiro")
      .Fields("Total Dinheiro").Value = Tot_Dinheiro + rsSaidas("Recebe - Dinheiro")
      .Fields("Parcelamento").Value = CDbl("0" & rsSaidas("Total Prazo"))
      .Fields("Total Parcelamento").Value = Tot_Parcelas + CDbl("0" & rsSaidas("Total Prazo"))
      .Fields("Final").Value = Tot_Dinheiro + rsSaidas("Recebe - Cart�o") + rsSaidas("Recebe - Vale") + rsSaidas("Recebe - Dinheiro") + Val_Cheques + Val_Cheques_Pr� + Tot_Cart�es + Tot_Vales + Tot_Cheques + Tot_Cheques_Pre
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
 'Solicitante: MRPR Automa��o
 '
 'Para a base deste cliente os campos abaixo em
 'algum momento estavam vindo com valores nulos.
 'Para isso foi criado o tratamento abaixo para n�o
 'ocorrer o erro 94 Invalid of null
 If (Not IsNull(rsSaidas("Recebe - Dinheiro"))) And (Not IsNull(rsSaidas("Recebe - Vale"))) Then
    Val_Vista = rsSaidas("Recebe - Dinheiro") + rsSaidas("Recebe - Vale")
 Else
    Val_Vista = 0
 End If
 
 Rem Fazer Lan�amentos em Controle de Cheques
 rsSa�da_Cheques.Index = "Ordem"
 rsBancos.Index = "C�digo"
 Ordem = 0
 
 Do
  Erro = False
  rsSa�da_Cheques.Seek ">", Filial, Mov, Ordem
  If rsSa�da_Cheques.NoMatch Then Erro = True
  If Erro = False Then If rsSa�da_Cheques("Filial") <> Filial Then Erro = True
  If Erro = False Then If rsSa�da_Cheques("Sequ�ncia") <> Mov Then Erro = True
  
  If Erro = False Then
    Ordem = rsSa�da_Cheques("Ordem")
    'If rsSa�da_Cheques("Bom") <> rsSaidas("Data") Then
      rsContas_Receber.AddNew
        
        rsContas_Receber("Tipo") = "C"
        rsContas_Receber("Filial") = Filial
        rsContas_Receber("Sequ�ncia") = Mov
        rsContas_Receber("Cliente") = rsSaidas("Cliente")
        rsContas_Receber("Banco") = rsSa�da_Cheques("Banco")
        rsBancos.Seek "=", rsSa�da_Cheques("Banco")
        If rsBancos.NoMatch Then rsContas_Receber("Banco") = 999
        rsContas_Receber("Cheque") = rsSa�da_Cheques("Cheque")
        rsContas_Receber("Vencimento") = rsSa�da_Cheques("Bom")
        rsContas_Receber("Valor") = rsSa�da_Cheques("Valor")
        rsContas_Receber("Vendedor") = rsSaidas("Digitador")
        rsContas_Receber("Data Emiss�o") = rsSaidas("Data")
        rsContas_Receber("Data Altera��o") = Format(Date, "dd/mm/yyyy")
        If rsSa�da_Cheques("Bom") = rsSaidas("Data") Then
          rsContas_Receber("Processado") = True
          rsContas_Receber("Valor Recebido") = rsSa�da_Cheques("Valor")
          rsContas_Receber("Data Recebimento") = rsSa�da_Cheques("Bom")
        End If
        
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequ�ncia") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
          "modEfetivaSaida_Efetiva_Sa�da", _
          "Contas a Receber", g_strArquivoSystemLog
        End If
      rsContas_Receber.Update
     
'     24/02/2003 - Maikel
'     Descri��o : Na tabela de contas a receber era gravado o cheque duas vezes, uma linha era do que foi informado
'                 na tela de recebimento, outra era de pagamento a vista segundo as linhas comentadas abaixo. Consequencia:
'                 no relat�rio de fluxo de caixa aparecia o cheque duas vezes.
'---------------------------------------------------------------------------------
'     Rem Grava conta recebida
'     If rsParametros("Gerar Conta Paga") = True Then
'      If rsSa�da_Cheques("Bom") = rsSaidas("Data") Then
'       rsContas_Receber.AddNew
'         rsContas_Receber("Tipo") = "R"
'         rsContas_Receber("Filial") = Filial
'         rsContas_Receber("Cliente") = rsSaidas("Cliente")
'         rsContas_Receber("Data Emiss�o") = rsSaidas("Data")
'         rsContas_Receber("Descri��o") = "Pagamento � vista"
'         rsContas_Receber("Vencimento") = rsSaidas("Data")
'         rsContas_Receber("Valor") = rsSa�da_Cheques("Valor")
'         rsContas_Receber("Sequ�ncia") = Mov
'         rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
'         rsContas_Receber("Vendedor") = rsSaidas("Digitador")
'         rsContas_Receber("Data Altera��o") = Format(Date, "dd/mm/yyyy")
'         rsContas_Receber("Data Recebimento") = rsSaidas("Data")
'         rsContas_Receber("Valor Recebido") = rsSa�da_Cheques("Valor")
'       rsContas_Receber.Update
'      End If
'     End If
'---------------------------------------------------------------------------------

   ' End If
  End If
 Loop Until Erro = True

 
 Rem Faz Lan�amentos em controle de cart�es, se for o caso
 If rsSaidas("Recebe - Cart�o") <> 0 Then
  
   rsContas_Receber.Index = "Contas"
   rsContas_Receber.Seek ">", "O", gnCodFilial, rsSaidas("Sequ�ncia"), 0
   If Not rsContas_Receber.NoMatch Then
     If rsContas_Receber("Tipo") = "O" Then
       If rsContas_Receber("Filial") = gnCodFilial Then
         If rsContas_Receber("Sequ�ncia") = rsSaidas("Sequ�ncia") Then
            '10/09/2007 - Anderson
            'Gera arquivo log do sistema
            If g_bolSystemLog Then
              SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Excluir, _
              "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequ�ncia") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
              "modEfetivaSaida_Efetiva_Sa�da", _
              "Contas a Receber", g_strArquivoSystemLog
            End If
           rsContas_Receber.Delete
         End If
       End If
    End If
   End If
   
   '10/12/2009 - Andrea
   'Verifica se a vari�vel Val_Cartoes = 0 (esta vari�vel acumula o valor recebido em cart�es na tela de Recebimento)
   'Se tiver zerada, o valor que foi recebido em cart�es veio do Quick Fiscal
   If Val_Cartoes = 0 Then
          
     For nI = 0 To rsSaidas("Qtde Parcelas") - 1
       rsCartoes.Index = "C�digo"
       rsCartoes.Seek "=", rsSaidas("Recebe - Emp Cart�o")
       If Not rsCartoes.NoMatch Then
         rsContas_Receber.AddNew
         rsContas_Receber("Tipo") = "O"
         rsContas_Receber("Filial") = gnCodFilial
         rsContas_Receber("Sequ�ncia") = rsSaidas("Sequ�ncia")
         rsContas_Receber("Cliente") = rsSaidas("Cliente")
         rsContas_Receber("Administradora") = rsSaidas("Recebe - Emp Cart�o")
         rsContas_Receber("Cart�o") = rsSaidas("Recebe - Num Cart�o")
         
         '08/10/2007 - Anderson
         'Altera��o para evitar que o vencimento dos cart�es seja nos finais de semana
         'Solitante: Agrotama
         rsContas_Receber("Vencimento") = (rsSaidas("Data") + rsCartoes("Dias Pagar") + (nI * 30))
         If (Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 1 Or Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 2) Then
           rsContas_Receber("Vencimento") = DateAdd("d", 3 - Weekday(rsContas_Receber("Vencimento"), vbSaturday), rsContas_Receber("Vencimento"))
         End If
         
         rsContas_Receber("Data Emiss�o") = rsSaidas("Data")
         If rsSaidas("Qtde Parcelas") = 1 Then
           rsContas_Receber("Valor Cart�o") = rsSaidas("Recebe - Cart�o")
           rsContas_Receber("Valor") = Round(CDbl(rsSaidas("Recebe - Cart�o") * ((1 - rsCartoes("Taxa") / 100))), 2)
         Else
           rsContas_Receber("Valor Cart�o") = rsSaidas("Valor Parcela")
           rsContas_Receber("Valor") = Round(CDbl(rsSaidas("Valor Parcela") * ((1 - rsCartoes("Taxa") / 100))), 2)
         End If
         rsContas_Receber("Data Altera��o") = Format(Date, "dd/mm/yyyy")
         
         '10/09/2007 - Anderson
         'Gera arquivo log do sistema
         If g_bolSystemLog Then
           SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
           "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequ�ncia") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
           "modEfetivaSaida_Efetiva_Sa�da", _
           "Contas a Receber", g_strArquivoSystemLog
         End If
         rsContas_Receber.Update
       End If
     Next nI
  '10/12/2009 - Andrea
  'Sen�o (se foi preenchido em outras telas) gravado os cart�es no contas a receber a partir da tabela Movimento - Cart�es
  Else
    nI = 0
    Ordem = 0
    Erro = False
    Do
     rsSa�da_Cartoes.Seek ">", Filial, Mov, Ordem
     If rsSa�da_Cartoes.NoMatch Then Erro = True
     If Erro = False Then If rsSa�da_Cartoes("Filial") <> Filial Then Erro = True
     If Erro = False Then If rsSa�da_Cartoes("Sequ�ncia") <> Mov Then Erro = True

     If Erro = False Then
       
       'Val_Cartoes = Val_Cartoes + rsSa�da_Cartoes("Valor")
       rsCartoes.Index = "Nome"
       rsCartoes.Seek "=", rsSa�da_Cartoes("Administradora")
       If Not rsCartoes.NoMatch Then
         For nI = 0 To rsSa�da_Cartoes("Parcelas") - 1
          rsContas_Receber.AddNew
          rsContas_Receber("Tipo") = "O"
          rsContas_Receber("Filial") = gnCodFilial
          rsContas_Receber("Sequ�ncia") = rsSaidas("Sequ�ncia")
          rsContas_Receber("Cliente") = rsSaidas("Cliente")
          rsContas_Receber("Administradora") = rsCartoes("C�digo")
          rsContas_Receber("Cart�o") = rsSa�da_Cartoes("NumeroCartao")
          
          'Evita que o vencimento dos cart�es seja nos finais de semana
          rsContas_Receber("Vencimento") = (rsSaidas("Data") + rsCartoes("Dias Pagar") + (nI * 30))
          If (Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 1 Or Weekday(rsContas_Receber("Vencimento"), vbSaturday) = 2) Then
            rsContas_Receber("Vencimento") = DateAdd("d", 3 - Weekday(rsContas_Receber("Vencimento"), vbSaturday), rsContas_Receber("Vencimento"))
          End If
          
          rsContas_Receber("Data Emiss�o") = rsSaidas("Data")
          If rsSa�da_Cartoes("Parcelas") = 1 Then 'Cart�o em 1 parcela
            rsContas_Receber("Valor Cart�o") = rsSa�da_Cartoes("Valor")
            rsContas_Receber("Valor") = Round(CDbl(rsSa�da_Cartoes("Valor") * ((1 - rsCartoes("Taxa") / 100))), 2)
          Else 'Cart�o Parcelado
            If nI = 0 Then ' � a primeira parcela
              Dim dbl_valor_parcela As Double
              Dim sht_numero_parcelas As Integer
              Dim dbl_valor_parcelar As Double
              
              sht_numero_parcelas = rsSa�da_Cartoes("Parcelas")
              dbl_valor_parcelar = rsSa�da_Cartoes("Valor")
              dbl_valor_parcela = dbl_valor_parcelar / sht_numero_parcelas
                
              'usada para arredondamento das parcelas para valores inteiros
              'Dim dbl_adicional_primera_parcela As Double
              Dim dbl_primeira_parcela As Double
              
              'dbl_adicional_primera_parcela = 0
              dbl_valor_parcela = Round(dbl_valor_parcela, 2)
              dbl_primeira_parcela = 0
              
              '========================================================================
              ' Tratamento para dizima peri�dica
              '========================================================================
              If (dbl_valor_parcela * sht_numero_parcelas) < dbl_valor_parcelar Then
                dbl_primeira_parcela = dbl_valor_parcelar - (dbl_valor_parcela * sht_numero_parcelas)
                dbl_primeira_parcela = dbl_primeira_parcela + dbl_valor_parcela
                dbl_primeira_parcela = Round(dbl_primeira_parcela, 2)
                rsContas_Receber("Valor Cart�o") = dbl_primeira_parcela
                rsContas_Receber("Valor") = Round(CDbl(dbl_primeira_parcela) * ((1 - rsCartoes("Taxa") / 100)), 2)
              Else
                Dim J As Double
                J = 0
                dbl_primeira_parcela = dbl_valor_parcela
                J = (dbl_valor_parcelar - (dbl_valor_parcela * sht_numero_parcelas))
                dbl_primeira_parcela = dbl_primeira_parcela + J
                dbl_primeira_parcela = Round(dbl_primeira_parcela, 2)
                rsContas_Receber("Valor Cart�o") = dbl_primeira_parcela
                rsContas_Receber("Valor") = Round(CDbl(dbl_primeira_parcela) * ((1 - rsCartoes("Taxa") / 100)), 2)
              End If
            Else
              'Cart�o parcelado - segunda parcela em diante
              rsContas_Receber("Valor Cart�o") = rsSa�da_Cartoes("Valor") / rsSa�da_Cartoes("Parcelas")
              rsContas_Receber("Valor") = Round(CDbl(rsContas_Receber("Valor Cart�o") * ((1 - rsCartoes("Taxa") / 100))), 2)
            End If
          End If
          rsContas_Receber("Data Altera��o") = Format(Date, "dd/mm/yyyy")
          
          'Gera arquivo log do sistema
          If g_bolSystemLog Then
            SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
            "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequ�ncia") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
            "modEfetivaSaida_Efetiva_Sa�da", _
            "Contas a Receber", g_strArquivoSystemLog
          End If
          rsContas_Receber.Update
         Next nI
       End If

       Ordem = rsSa�da_Cartoes("Ordem")
     End If
    Loop Until Erro = True
      
  End If
 End If
 
  Rem Faz contas a receber, se for o caso
 Erro = False
 Ordem = 0
 Aux_Int = 1
 rsSa�da_Parcelas.Index = "Ordem"
 
 Do
   rsSa�da_Parcelas.Seek ">", Filial, Mov, Ordem
   If rsSa�da_Parcelas.NoMatch Then Erro = True
   If Erro = False Then If rsSa�da_Parcelas("Filial") <> Filial Then Erro = True
   If Erro = False Then If rsSa�da_Parcelas("Sequ�ncia") <> Mov Then Erro = True

   If Erro = False Then
     Ordem = rsSa�da_Parcelas("Ordem")
       rsContas_Receber.AddNew
         rsContas_Receber("Tipo") = "R"
         rsContas_Receber("Filial") = Filial
         rsContas_Receber("Cliente") = rsSaidas("Cliente")
         rsContas_Receber("Data Emiss�o") = rsSaidas("Data")
         rsContas_Receber("Parcela") = Trim(str(Aux_Int))
         rsContas_Receber("Descri��o") = "Parcela " & str(Aux_Int) & "/" & str(rsSa�da_Parcelas("Parcelas"))
         rsContas_Receber("Vencimento") = rsSa�da_Parcelas("Bom")
         rsContas_Receber("Valor") = rsSa�da_Parcelas("Valor")
         rsContas_Receber("Sequ�ncia") = Mov
         '21/02/2005 - Daniel
         'Tratamento para n�o ocorrer o erro 94 (Invalid use of null)
         'Solicitante: MRPR Automa��o - Curitiba - PR
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
         rsContas_Receber("Data Altera��o") = Format(Date, "dd/mm/yyyy")

        '25/09/2007 - Anderson
        'Implementa��o do campo c�digo de Barras para impress�o em Carn�s
        If rsContas_Receber("Tipo Parcelamento") = "T" Then
          rsContas_Receber("CarneCodigoBarras") = "*" & Format(Filial, "00") & Format(rsSaidas("Cliente"), "000000") & Format(Mov, "000000") & Format(Trim(str(Aux_Int)), "00") & "*"
        End If
         
        '10/09/2007 - Anderson
        'Gera arquivo log do sistema
        If g_bolSystemLog Then
          SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
          "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequ�ncia") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
          "modEfetivaSaida_Efetiva_Sa�da", _
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
       rsContas_Receber("Data Emiss�o") = rsSaidas("Data")
       rsContas_Receber("Descri��o") = "Pagamento � vista"
       rsContas_Receber("Vencimento") = rsSaidas("Data")
       rsContas_Receber("Valor") = Val_Vista
       rsContas_Receber("Sequ�ncia") = Mov
       rsContas_Receber("Nota") = rsSaidas("Nota Impressa")
       rsContas_Receber("Vendedor") = rsSaidas("Digitador")
       rsContas_Receber("Data Altera��o") = Format(Date, "dd/mm/yyyy")
       rsContas_Receber("Data Recebimento") = rsSaidas("Data")
       rsContas_Receber("Valor Recebido") = Val_Vista
      
      '10/09/2007 - Anderson
      'Gera arquivo log do sistema
      If g_bolSystemLog Then
        SystemLog Format(Now, "dd/mm/yyyy"), Format(Now, "hh:mm"), gnUserCode & "-" & gsUserName, Inserir, _
        "Cli:" & rsContas_Receber("Cliente") & "- Seq:" & rsContas_Receber("Sequ�ncia") & "- NF:" & rsContas_Receber("Nota") & "- Venc:" & rsContas_Receber("Vencimento") & "- Valor:" & rsContas_Receber("Valor"), _
        "modEfetivaSaida_Efetiva_Sa�da", _
        "Contas a Receber", g_strArquivoSystemLog
      End If
     rsContas_Receber.Update
   End If
 End If
 
 Rem Atualiza Resumo Di�rio
 If rsOp_Sa�da("Tipo") <> "O" Then
   rsResumo_Di�rio.Index = "Data"
   rsResumo_Di�rio.Seek "=", Filial, rsSaidas("Data")
   If rsResumo_Di�rio.NoMatch Then
     rsResumo_Di�rio.AddNew
     rsResumo_Di�rio("Filial") = Filial
     rsResumo_Di�rio("Data") = rsSaidas("Data")
   Else
     rsResumo_Di�rio.LockEdits = True
     rsResumo_Di�rio.Edit
   End If
   If rsOp_Sa�da("Tipo") = "V" Then
        rsResumo_Di�rio("Valor Vendas") = CDbl(rsResumo_Di�rio("Valor Vendas")) + CDbl(rsSaidas("Total")) - CDbl(rsSaidas("Servi�os"))
        rsResumo_Di�rio("Valor Servi�os") = rsResumo_Di�rio("Valor Servi�os") + CDbl(rsSaidas("Servi�os"))
   End If
   If rsOp_Sa�da("Tipo") = "T" Then rsResumo_Di�rio("Valor T Sa�da") = CDbl(rsResumo_Di�rio("Valor T Sa�da")) + CDbl(rsSaidas("Total"))
   If rsOp_Sa�da("Tipo") = "A" Then rsResumo_Di�rio("Valor A Sa�da") = CDbl(rsResumo_Di�rio("Valor A Sa�da")) + CDbl(rsSaidas("Total"))
   If rsOp_Sa�da("Tipo") = "G" Then rsResumo_Di�rio("Valor G Sa�da") = CDbl(rsResumo_Di�rio("Valor G Sa�da")) + CDbl(rsSaidas("Total"))
   If rsOp_Sa�da("Tipo") = "E" Then rsResumo_Di�rio("Valor E Sa�da") = CDbl(rsResumo_Di�rio("Valor E Sa�da")) + CDbl(rsSaidas("Total"))

  rsResumo_Di�rio.Update
End If
 
 
 
 
  Rem Atualiza Resumo Di�rio Financeiro
  If rsOp_Sa�da("Dinheiro") = True Then
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
    
    If rsOp_Sa�da("Tipo") = "V" Then
        rsRes_Financeiro("Valor Vendas") = CDbl(rsRes_Financeiro("Valor Vendas")) + CDbl(rsSaidas("Total")) - CDbl(rsSaidas("Servi�os"))
        rsRes_Financeiro("Valor Servi�os") = CDbl(rsRes_Financeiro("Valor Servi�os")) + CDbl(rsSaidas("Servi�os"))
    End If
    If rsOp_Sa�da("Tipo") = "T" Then rsRes_Financeiro("Valor T Sa�da") = CDbl(rsRes_Financeiro("Valor T Sa�da")) + CDbl(rsSaidas("Total"))
    If rsOp_Sa�da("Tipo") = "A" Then rsRes_Financeiro("Valor A Sa�da") = CDbl(rsRes_Financeiro("Valor A Sa�da")) + CDbl(rsSaidas("Total"))
    If rsOp_Sa�da("Tipo") = "G" Then rsRes_Financeiro("Valor G Sa�da") = CDbl(rsRes_Financeiro("Valor G Sa�da")) + CDbl(rsSaidas("Total"))
    If rsOp_Sa�da("Tipo") = "E" Then rsRes_Financeiro("Valor E Sa�da") = CDbl(rsRes_Financeiro("Valor E Sa�da")) + CDbl(rsSaidas("Total"))


    rsRes_Financeiro.Update
  End If


  rsSaidas_Prod.Index = "Sequ�ncia"
  Ordem = 0
Prox_Prod:
  rsSaidas_Prod.Seek ">", Filial, Mov, Ordem
  
  If rsSaidas_Prod.NoMatch Then GoTo Ve_Serv
  If rsSaidas_Prod("Filial") <> Filial Then GoTo Ve_Serv
  If rsSaidas_Prod("sequ�ncia") <> Mov Then GoTo Ve_Serv
  
  Ordem = rsSaidas_Prod("Linha")
  'Verifica se tem grade
  C�d = rsSaidas_Prod("C�digo")
  Tamanho = 0
  Cor = 0
  Edi��o = 0
  
  If Not IsNull(rsSaidas_Prod("Descricao Adicional")) Then
     sDescrAdicional = rsSaidas_Prod("Descricao Adicional")
  Else
     sDescrAdicional = ""
  End If
  
   Aux_C�digo = Trim(C�d)
   Call Acha_Produto(Aux_C�digo, C�d, Tamanho, Cor, Edi��o, Aux_Tipo, Aux_Erro)
   If Aux_Erro <> 0 Then GoTo Prox_Prod
   C�d = Trim(UCase(C�d))
   rsProdutos.Index = "C�digo"
   rsProdutos.Seek "=", C�d
   
  'Neste ponto C�D tem o c�digo do produto
  'Tamanho, Cor e Edi��o cont�m os respectivos dados
  'Agora grava arquivo do estoque
  
  Call StatusMsg("Atualizando estoque de " & rsProdutos("Nome"))
  
  '---[ Gera acerto de consigna��o de entrada ]---'
  '13/08/2004 - Maikel
  '
  '27/08/2004 - Daniel
  'Valida��o realizada em cima do campo [Opera��es Sa�da].AcertaEmprestimoEntrada
  If rsOp_Sa�da("AcertaEmprestimoEntrada").Value Then
    dblPrecoVenda = Format(rsSaidas_Prod("Pre�o Final").Value, FORMAT_VALUE)
  
    Call GeraAcertoConsignacaoEntrada(gnCodFilial, rsSaidas_Prod("Sequ�ncia").Value, rsSaidas_Prod("C�digo Sem Grade").Value, rsSaidas_Prod("Qtde").Value, dblPrecoVenda)
  End If
  '---[ Gera acerto de consigna��o de entrada ]---'
  'testar aqui
  'Ajusta Estoque
  If rsOp_Sa�da("Estoque") = True And Not rsSaidas_Prod("InGeradoViaConsig") Then
  
'-------------------------------------------------------------------------------------
    '14/11/2003 - mpdea
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
             " ORDER BY Data"
             
'    'LOG ESPECIFICO PARA MARE MANSA
'    Dim sSQL_Aux As String
'    sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'    sSQL_Aux = sSQL_Aux & Left("1 - " & Filial & " : " & C�d & " : " & Tamanho & " : " & Cor & " : " & Edi��o, 80) & "', 'VENDENDO MARE')"
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
'      sSQL_Aux = sSQL_Aux & Left("2 - " & C�d & " : " & Estoque_Final & " : " & Format(sData_mare, "mm/dd/yyyy"), 80) & "', 'VENDENDO MARE')"
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
             " AND Data = #" & Format(rsSaidas("Data"), "mm/dd/yyyy") & "#"
            
'    'LOG ESPECIFICO PARA MARE MANSA
'    sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'    sSQL_Aux = sSQL_Aux & Left("3 - " & Filial & " : " & C�d & " : " & Tamanho & " : " & Cor & " : " & Edi��o & " : " & Format(rsSaidas("Data"), "mm/dd/yyyy"), 80) & "', 'VENDENDO MARE')"
'    db.Execute sSQL_Aux, dbFailOnError
'    'fim
            
    Set rsEstoque = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rsEstoque
      If Not (.BOF And .EOF) Then
        .LockEdits = True
        .Edit
    
'        'LOG ESPECIFICO PARA MARE MANSA
'        sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'        sSQL_Aux = sSQL_Aux & Left("4B - " & C�d & " : UPDATE NA LINHA NA TABELA DE ESTOQUE", 80) & "', 'VENDENDO MARE')"
'        db.Execute sSQL_Aux, dbFailOnError
'        'fim
      
      Else
        .AddNew
        .Fields("Filial").Value = Filial
        .Fields("Data").Value = rsSaidas("Data").Value
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
'        sSQL_Aux = "INSERT INTO ZZZLog (Data, Texto, Tipo) VALUES (#" & Now & "#, '"
'        sSQL_Aux = sSQL_Aux & Left("4A - " & C�d & " : " & rsSaidas("Data").Value & " : " & Estoque_Final & " : " & "NOVA LINHA NA TABELA DE ESTOQUE", 80) & "', 'VENDENDO MARE')"
'        db.Execute sSQL_Aux, dbFailOnError
'        'fim
      
      End If
    End With
'-------------------------------------------------------------------------------------
    
'    Rem Encontra a posi��o do estoque
'    Criar_Registro = False
'    Estoque_Final = 0
'    rsEstoque.Index = "Produto"
'    rsEstoque.Seek "=", Filial, rsSaidas("Data"), C�d, Tamanho, Cor, Edi��o
'
'    If Not rsEstoque.NoMatch Then Estoque_Final = rsEstoque("Estoque Final")
'    If rsEstoque.NoMatch Then
'       rsEstoque.Index = "Data"
'       rsEstoque.Seek "<", Filial, C�d, Tamanho, Cor, Edi��o, rsSaidas("Data")
'       If rsEstoque.NoMatch Then Criar_Registro = True
'       If Not rsEstoque.NoMatch Then
'          If rsEstoque("Filial") = Filial And rsEstoque("Produto") = C�d And rsEstoque("Tamanho") = Tamanho And rsEstoque("Cor") = Cor And rsEstoque("Edi��o") = Edi��o Then
'             Criar_Registro = True
'             Estoque_Final = rsEstoque("Estoque Final")
'           End If
'       End If
'
'       rsEstoque.AddNew
'       rsEstoque("Filial") = Filial
'       rsEstoque("Data") = rsSaidas("Data")
'       rsEstoque("Produto") = C�d
'       rsEstoque("Tamanho") = Tamanho
'       rsEstoque("Cor") = Cor
'       rsEstoque("Edi��o") = Edi��o
'       rsEstoque("Classe") = rsProdutos("Classe")
'       rsEstoque("Sub Classe") = rsProdutos("Sub Classe")
'       rsEstoque("Estoque Anterior") = Estoque_Final
'       rsEstoque.Update
'
'       rsEstoque.Index = "Produto"
'       rsEstoque.Seek "=", Filial, rsSaidas("Data"), C�d, Tamanho, Cor, Edi��o
'    End If

'-------------------------------------------------------------------------------------

      Rem neste ponto esta com o registro de estoque
      Rem no buffer, agora soma com os valores da movimenta��o
'      rsEstoque.Edit
      If rsOp_Sa�da("Tipo") = "V" Then
         rsEstoque("Vendas") = rsEstoque("Vendas") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Vendas") = Format(rsEstoque("Valor Vendas") + rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
      If rsOp_Sa�da("Tipo") = "T" Then
         rsEstoque("Transf Sa�da") = rsEstoque("Transf Sa�da") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor T Sa�da") = Format(rsEstoque("Valor T Sa�da") + rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
      If rsOp_Sa�da("Tipo") = "A" Then
         rsEstoque("Ajuste Sa�da") = rsEstoque("Ajuste Sa�da") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Ajuste Sa�da") = Format(rsEstoque("Valor Ajuste Sa�da") + rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
      If rsOp_Sa�da("Tipo") = "G" Then
         rsEstoque("Gr�tis Sa�da") = rsEstoque("Gr�tis Sa�da") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Gr�tis Sa�da") = Format(rsEstoque("Valor Gr�tis Sa�da") + rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If
      If rsOp_Sa�da("Tipo") = "E" Then
         rsEstoque("Empre Sa�da") = rsEstoque("Empre Sa�da") + rsSaidas_Prod("Qtde")
         rsEstoque("Valor Empre Sa�da") = Format(rsEstoque("Valor Empre Sa�da") + rsSaidas_Prod("Pre�o Final"), "############0.00")
      End If

      Estoque_Final = rsEstoque("Estoque Anterior") - rsEstoque("Vendas") + rsEstoque("Compras")
      Estoque_Final = Estoque_Final - rsEstoque("Transf Sa�da") + rsEstoque("Transf Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Ajuste Sa�da") + rsEstoque("Ajuste Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Gr�tis Sa�da") + rsEstoque("Gr�tis Entra")
      Estoque_Final = Estoque_Final - rsEstoque("Empre Sa�da") + rsEstoque("Empre Entra")
      
      '08/08/2003 - maikel
      '             Descomentada a soma da coluna Devolu��o para resolver o problema de estoque
      Estoque_Final = Estoque_Final - rsEstoque("Quebras") + rsEstoque("Devolu��o")

      If rsProdutos("Estoque") = False Then
        Estoque_Final = 0
      End If

      rsEstoque("Estoque Final") = Estoque_Final
      rsEstoque.Update
      
      rsEstoque.LockEdits = False
      
      rsEstoque.Close
      
      Rem Arruma Estoque Final
      Grava_Estoque_Final gnCodFilial, C�d, Tamanho, Cor, Edi��o, CSng(Estoque_Final), rsSaidas("Data")
      
      
  End If

  rsProdutos.Index = "C�digo"
  rsProdutos.Seek "=", C�d

  Rem Grava Conta do Cliente, se for o caso
  If rsSaidas("Recebe - Conta") = True Then
     rsConta_Cli.AddNew
     rsConta_Cli("Filial") = Filial
     rsConta_Cli("Cliente") = rsSaidas("Cliente")
     rsConta_Cli("Data") = rsSaidas("Data")
     rsConta_Cli("Produto") = C�d
     
     strAuxiliar = rsProdutos("Nome") & ""
     strAuxiliar = Replace(strAuxiliar, "        ", " ")
     strAuxiliar = Replace(strAuxiliar, "       ", " ")
     strAuxiliar = Replace(strAuxiliar, "      ", " ")
     strAuxiliar = Replace(strAuxiliar, "     ", " ")
     strAuxiliar = Replace(strAuxiliar, "    ", " ")
     
     If Len(Trim(strAuxiliar)) > 70 Then

        rsConta_Cli("Descri��o") = Left(Trim(strAuxiliar & ""), 70)
     Else
        rsConta_Cli("Descri��o") = strAuxiliar
     End If
     
     rsConta_Cli("Qtde") = rsSaidas_Prod("Qtde")
     rsConta_Cli("Valor") = rsSaidas_Prod("Pre�o Final")
     rsConta_Cli("TabPrecos") = rsSaidas("Tabela") & ""
     rsConta_Cli("Sequ�ncia") = rsSaidas("Sequ�ncia")
     rsConta_Cli("Data Altera��o") = Format(Date, "dd/mm/yyyy")
     rsConta_Cli.Update
  End If

  '18/07/2003 - mpdea
  'Formatado valores com FORMAT_VALUE e
  'comiss�o foi truncada com 6 casas decimais
  '
  'Grava comiss�es
  'If rsOp_Sa�da("Comiss�o") = True Then
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
'      '---[ Adicionada a verifica��o para diminuir a comiss�o do funcion�rio caso haja desconto na movimenta��o ]---'
'        If (rsSaidas.Fields("Desconto")) > 0 Or (rsSaidas.Fields("DescontoSubTotal") > 0) Then
'          blnDiminuiComissao = True
'        End If
'      '---[ Adicionada a verifica��o para diminuir a comiss�o do funcion�rio caso haja desconto na movimenta��o ]---'
'
'      Comiss�o = rsFuncionarios("Comiss�o")
'      If rsProdutos("Comiss�o Sobrep�e") = True Then
'        Comiss�o = rsProdutos("Comiss�o")
'      End If
'      Comiss�o = Comiss�o * rsTabelas("Multiplicador Comiss�o")
'      Comiss�o = Format(Comiss�o, FORMAT_VALUE)
'
'      Set objComissao = New clsComissao
'      objComissao.Filial = Filial
'      objComissao.Data = rsSaidas("Data")
'      objComissao.Vendedor = iComissionado
'      objComissao.Produto = C�d
'      objComissao.Tamanho = Tamanho
'      objComissao.Cor = Cor
'      objComissao.Edi��o = Edi��o
'      objComissao.Qtde = rsSaidas_Prod("Qtde")
'
'      If Not IsNull(rsSaidas("Recebe - Vale")) And rsSaidas("Recebe - Vale") <> 0 Then
'          objComissao.Valor = Format(rsSaidas_Prod("Pre�o Final") - rsSaidas("Recebe - Vale"), FORMAT_VALUE)
'      Else
'          objComissao.Valor = Format(rsSaidas_Prod("Pre�o Final"), FORMAT_VALUE)
'      End If
'
'      objComissao.Sequ�ncia = rsSaidas("Sequ�ncia")
'
'      '-----------------------------------------------------------------------------------------------------
'      '03/07/2006 - mpdea
'      'Movido c�digos da tela de Recebimento para o m�dulo Efetiva Sa�da, pois
'      'n�o estava sendo executado no recebimento simplificado
'      '
'      '25/04/2005 - Daniel
'      '22/03/2005 - Daniel
'      'Case.....: Bem Me Quer
'      'Projeto..: Reten��o sobre comiss�es
'      If rsSaidas.Fields("Recebe - Cart�o").Value > 0 Then
'        '-----------------------------------------------------------------------------
'        '17/07/2006 - Andrea
'        'Alterado para acertar o valor pago em cartao por item para gravar no
'''        'arquivo de comissao com retencao.
        'essa linha era o que tinha 'dblVlPagoEmCartao = rsSaidas.Fields("Recebe - Cart�o").Value
'        '-----------------------------------------------------------------------------
'        dblTotalSaidas = rsSaidas.Fields("Total").Value
'        dblRecebe_Cartao = rsSaidas.Fields("Recebe - Cart�o").Value
''        dblPercentualPagoEmCartao = ((dblRecebe_Cartao * 100) / dblTotalSaidas)
'        dblVlPagoEmCartao = ((rsSaidas_Prod("Pre�o Final").Value * dblPercentualPagoEmCartao) / 100)
'        '-----------------------------------------------------------------------------
'        '
'
'        dblVlPagoEmCartaoComRetencao = 0
'
'        Set rstCartoes = db.OpenRecordset("SELECT Taxa FROM Cart�es WHERE C�digo = " & rsSaidas.Fields("Recebe - Emp Cart�o").Value, dbOpenDynaset, dbReadOnly)
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
'          dblValorComissao = (Comiss�o * (rsSaidas_Prod("Pre�o Final") - rsSaidas("Recebe - Vale")) / 100)
'      Else
'          dblValorComissao = (Comiss�o * rsSaidas_Prod("Pre�o Final") / 100)
'      End If
'
'
'      If blnDiminuiComissao Then
'        Set rstTabelaPrecos = db.OpenRecordset(" SELECT * FROM [Tabela de Pre�os] " & _
'                                               " WHERE Tabela = '" & rsSaidas.Fields("Tabela") & "'", dbOpenSnapshot)
'
'        With rstTabelaPrecos
'          If Not (.BOF And .EOF) Then
'            If IsNumeric(rstTabelaPrecos.Fields("PercentualComissaoDesconto")) Then
'              '11/02/2005 - Daniel
'              'Problema levantado pela Daring
'              'Se for dado desconto s� em um produto em uma nota com X produtos, a comiss�o
'              'do vendedor em todos os �tens da nota estava caindo pela metade sendo que o
'              'correto � reduzir s� a do produto que teve altera��o de pre�o ou desconto
'              If (rsSaidas.Fields("Desconto") > 0) Then 'Houve desconto para algum �tem. Nota: Para o desconto no subtotal continuamos abater de todos os �tens
'                Dim dblValorDoCadastroProduto As Double
'
'                Call ReduzirComissao(rsSaidas.Fields("Tabela") & "", rsSaidas_Prod("C�digo sem Grade") & "", dblValorDoCadastroProduto)
'
'                'Se for diferente ocorre o abatimento
'                If dblValorDoCadastroProduto <> Format((rsSaidas_Prod("Pre�o Final").Value / rsSaidas_Prod("Qtde")), FORMAT_VALUE) Then dblValorComissao = dblValorComissao * ((100 - rstTabelaPrecos.Fields("PercentualComissaoDesconto")) / 100)
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
'      objComissao.Comiss�o = Truncate(dblValorComissao, 6)
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
    rsEtiquetas.Index = "Funcion�rio"
    rsEtiquetas.Seek "=", rsSaidas("Digitador"), C�d, Tamanho, Cor
    If rsEtiquetas.NoMatch Then
       rsEtiquetas.AddNew
    Else
       rsEtiquetas.LockEdits = True
       rsEtiquetas.Edit
    End If
    rsEtiquetas("Funcion�rio") = rsSaidas("Digitador")
    rsEtiquetas("Produto") = C�d
    rsEtiquetas("Tamanho") = Tamanho
    rsEtiquetas("Cor") = Cor
    rsEtiquetas("Qtde") = rsEtiquetas("Qtde") + rsSaidas_Prod("Qtde")
    rsEtiquetas("Sequ�ncia") = Mov
    rsEtiquetas.Update
  End If

 
  '12/07/2002 - mpdea
  'Desvia em opera��o de reserva WEB
  If rsSaidas.Fields("Opera��o").Value <> intCodOpReserva Then
    Rem Atualiza arquivo de Resumo de Clientes
    Rem se for Comrpa
    If rsOp_Sa�da("Tipo") = "V" Then
       rsResumo_Clientes.Index = "Cliente"
       rsResumo_Clientes.Seek "=", rsSaidas("Cliente"), rsSaidas("Data"), C�d, Tamanho, Cor, Edi��o, Mov
       If rsResumo_Clientes.NoMatch Then
          rsResumo_Clientes.AddNew
            rsResumo_Clientes("Dia") = rsSaidas("Data")
            rsResumo_Clientes("Cliente") = rsSaidas("Cliente")
            rsResumo_Clientes("Produto") = C�d
            rsResumo_Clientes("Tamanho") = Tamanho
            rsResumo_Clientes("Cor") = Cor
            rsResumo_Clientes("Edi��o") = Edi��o
            rsResumo_Clientes("Qtde") = 0
            rsResumo_Clientes("Valor Total") = 0
            rsResumo_Clientes("Sequ�ncia") = Mov
            rsResumo_Clientes("Descricao Adicional") = ""
  '          rsResumo_Clientes("Descricao Adicional") = sDescrAdicional
       Else
          rsResumo_Clientes.LockEdits = True
          rsResumo_Clientes.Edit
       End If
  
        rsResumo_Clientes("Qtde") = rsResumo_Clientes("Qtde") + rsSaidas_Prod("Qtde")
        rsResumo_Clientes("Valor Total") = Format((rsResumo_Clientes("Valor Total") + rsSaidas_Prod("Pre�o Final")), "############0.00")
        rsResumo_Clientes("Filial") = Filial
        rsResumo_Clientes("Tipo") = "C"
        rsResumo_Clientes("Descricao Adicional") = rsResumo_Clientes("Descricao Adicional") & sDescrAdicional & "-"
          
       rsResumo_Clientes.Update
    End If
  End If


  Rem Atualiza arquivo de Empr�stimos
  If rsOp_Sa�da("Tipo") = "E" And Not rsSaidas_Prod("InGeradoViaConsig") Then
     rsEmprestimos.Index = "Cliente"
     
     
     Rem Saldo Emprestado = 0 para este empr�stimo
     Saldo_Emp = 0
     Ordem_Emp = Ordem
     Emp_Existe = False
             
     rsEmprestimos.Seek "<", gnCodFilial, Mov, rsSaidas("Cliente"), C�d, Tamanho, Cor, Edi��o, 999999
       If Not rsEmprestimos.NoMatch Then
         If rsEmprestimos("Filial") = gnCodFilial Then
           If rsEmprestimos("Sequ�ncia") = Mov Then
             If rsEmprestimos("Cliente") = rsSaidas("Cliente") Then
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
         rsEmprestimos.LockEdits = True
         rsEmprestimos.Edit
      Else
        rsEmprestimos.AddNew
          rsEmprestimos("Filial") = gnCodFilial
          rsEmprestimos("Sequ�ncia") = Mov
          rsEmprestimos("Cliente") = rsSaidas("Cliente")
          rsEmprestimos("Produto") = C�d
          rsEmprestimos("Tamanho") = Tamanho
          rsEmprestimos("Cor") = Cor
          rsEmprestimos("Edi��o") = Edi��o
          rsEmprestimos("Ordem") = Ordem_Emp
      End If
      
      rsEmprestimos("Saldo Anterior") = Saldo_Emp
      rsEmprestimos("Novo Empr�stimo") = rsSaidas_Prod("Qtde")
      rsEmprestimos("Saldo Atual") = Saldo_Emp + rsSaidas_Prod("Qtde")
      rsEmprestimos("Pre�o Unit�rio") = (rsSaidas_Prod("Pre�o Final") / rsSaidas_Prod("Qtde"))
      rsEmprestimos("Data Opera��o") = rsSaidas("Data")
      rsEmprestimos("Data Altera��o") = Format(Date, "dd/mm/yyyy")
      rsEmprestimos("Data Cobran�a") = rsSaidas("Data Acerto Empr�stimo")

      rsEmprestimos.Update
     
  End If
  
  GoTo Prox_Prod
  
  
  
  
  
  
  
Ve_Serv:
  rsSaidas_Serv.Index = "Sequ�ncia"
  rsServicos.Index = "C�digo"
  rsFuncionarios.Index = "C�digo"
  Ordem = 0
Prox_Serv:
  rsSaidas_Serv.Seek ">", Filial, Mov, Ordem
  
  If rsSaidas_Serv.NoMatch Then GoTo Fim_Efetiva
  If rsSaidas_Serv("Filial") <> Filial Then GoTo Fim_Efetiva
  If rsSaidas_Serv("Sequ�ncia") <> Mov Then GoTo Fim_Efetiva
  
  Ordem = rsSaidas_Serv("Linha")
  
  C�d_Serv = rsSaidas_Serv("C�digo")
  rsServicos.Seek "=", C�d_Serv
  If rsServicos.NoMatch Then GoTo Prox_Serv
  
  Comiss_T�cnico = 0
  rsFuncionarios.Seek "=", rsSaidas("T�cnico")
  If Not rsFuncionarios.NoMatch Then
    Comiss_T�cnico = rsFuncionarios("Comiss�o Servi�o")
  End If
  
  If rsServicos("Comiss�o Sobrep�e") = True Then
    Comiss_T�cnico = rsServicos("Comiss�o")
  End If
  
  
  '--------------------------------------------------------------------------------
  '02/09/2003 - mpdea
  'Corrigido c�lculo do valor do servi�o
  '
  '18/07/2003 - mpdea
  'Comiss�o foi truncada com 6 casas decimais
  With rsComiss�es_Serv
    .AddNew
    .Fields("Data").Value = rsSaidas.Fields("Data").Value
    .Fields("Vendedor").Value = rsSaidas.Fields("T�cnico").Value
    .Fields("Servi�o").Value = C�d_Serv
    .Fields("Descri��o").Value = rsSaidas_Serv.Fields("Descri��o").Value & ""
    .Fields("Tempo").Value = rsSaidas_Serv.Fields("Tempo").Value
    .Fields("Valor").Value = CDbl(Format(CDbl(rsSaidas_Serv.Fields("Tempo").Value) * rsSaidas_Serv.Fields("Pre�o").Value, FORMAT_VALUE))
    .Fields("Comiss�o").Value = Comiss_T�cnico
    .Fields("Valor Comiss�o").Value = Truncate((.Fields("Valor").Value * Comiss_T�cnico / 100), 6)
    .Fields("Sequ�ncia").Value = rsSaidas.Fields("Sequ�ncia").Value
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
     .Fields("Produto").Value = C�d_Serv
     
     If Len(Trim(rsSaidas_Serv.Fields("Descri��o").Value)) > 70 Then
        .Fields("Descri��o").Value = Left(Trim(rsSaidas_Serv.Fields("Descri��o").Value & ""), 70)
     Else
        .Fields("Descri��o").Value = rsSaidas_Serv.Fields("Descri��o").Value & ""
     End If
     
     .Fields("Qtde").Value = CSng(rsSaidas_Serv.Fields("Tempo").Value)
     .Fields("Valor").Value = CDbl(Format(CDbl(rsSaidas_Serv.Fields("Tempo").Value) * rsSaidas_Serv.Fields("Pre�o").Value, FORMAT_VALUE))
     .Fields("TabPrecos").Value = rsSaidas.Fields("Tabela").Value & "" 'Jose
     .Fields("Sequ�ncia").Value = rsSaidas.Fields("Sequ�ncia").Value
     .Fields("Data Altera��o").Value = Format(Date, "dd/mm/yyyy")
     .Update
    End With
  End If
  '--------------------------------------------------------------------------------
  
  
  GoTo Prox_Serv
  
  
Fim_Efetiva:

  'Verifica programa de fidelidade
  If gParticipaProgramaFidelidade = 1 Then 'Lan�ar registro no programa
      ProgramaFidelidadeCriarLancamento rsSaidas.Fields("Opera��o").Value, rsSaidas.Fields("Total").Value, rsSaidas.Fields("Cliente").Value, rsCliFor.Fields("CGC").Value, rsSaidas.Fields("Operador").Value, Mov, rsCliFor.Fields("Nome").Value
  
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
  
  'Efetiva a Sa�da
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
 rsOp_Sa�da.Close
 rsResumo_Di�rio.Close
 rsEtiquetas.Close
 rsRes_Financeiro.Close
'  If Not rsEstoque Is Nothing Then rsEstoque.Close
' rsEstoque_Final.Close
 rsPre�os.Close
 rsResumo_Clientes.Close
 rsEmprestimos.Close
 rsCliFor.Close
 rsGrade.Close
 rsSaidas_Prod.Close
 rsSaidas_Serv.Close
 rsSa�da_Cheques.Close
 rsSa�da_Parcelas.Close
 'rsComiss�es.Close
 rsComiss�es_Serv.Close
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
 Set rsOp_Sa�da = Nothing
 Set rsResumo_Di�rio = Nothing
 Set rsEtiquetas = Nothing
 Set rsRes_Financeiro = Nothing
 Set rsEstoque = Nothing
' Set rsEstoque_Final = Nothing
 Set rsPre�os = Nothing
 Set rsResumo_Clientes = Nothing
 Set rsEmprestimos = Nothing
 Set rsCliFor = Nothing
 Set rsGrade = Nothing
 Set rsSaidas_Prod = Nothing
 Set rsSaidas_Serv = Nothing
 Set rsSa�da_Cheques = Nothing
 Set rsSa�da_Parcelas = Nothing
 'Set rsComiss�es = Nothing
 Set rsComiss�es_Serv = Nothing
 Set rsFuncionarios = Nothing
 Set rsTabelas = Nothing
 Set rsConta_Cli = Nothing
 Set rsCartoes = Nothing
 Set rsBancos = Nothing
 Set rsServicos = Nothing

   Efetiva_Sa�da = 0
   
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
          Efetiva_Sa�da = -1 'A��o cancelada
          Exit Function
        End If
        
'        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
'          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
'          "uma nova tentativa.", vbExclamation + vbOKCancel, "Efetiva Sa�da") = vbOK Then
'          intRepeatUpdateLocked = 0
'          Resume
'        Else
'          Efetiva_Sa�da = -1 'A��o cancelada
'          Exit Function
'        End If
      End If
    Case Else
      'Repassa para a fun��o de origem os outros erros
      Err.Raise Err.Number, Err.Source, Err.Description
      Exit Function
      
'      'Outros Erros
'      Select Case frmErro.gnShowErr(Err.Number, "Efetiva Sa�da")
'        Case 0 'Repetir
'          Resume
'        Case 1 'Prosseguir
'          Resume Next
'        Case 2 'Sair
'          Efetiva_Sa�da = -1 'A��o cancelada
'          Exit Function
'        Case 3 'Encerrar
'          End
'      End Select
  End Select

End Function


'mpdea
'Verifica a 1� abertura de um caixa, o seu in�cio de dia
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
  
  'Inicia transa��o
  If blnTransaction Then ws.BeginTrans: blnInTransaction = True
  
  sSql = "SELECT * FROM Caixa WHERE Filial = " & gnCodFilial & " AND Caixa = " & _
    nCaixa & " ORDER BY Filial, Caixa, Data, Ordem"
  Set rsCaixa = db.OpenRecordset(sSql, dbOpenDynaset)
  
  With rsCaixa
    'Verifica se h� informa��es sobre o Caixa solicitado
    If .RecordCount = 0 Then
      '� criado seu registro inicial (1� ocorr�ncia do Caixa)
      .AddNew
      !Filial = gnCodFilial
      !Data = Data_Atual
      !Caixa = nCaixa
      !Ordem = 1
      !Funcion�rio = 0
      !Descri��o = "In�cio do dia"
      !Hora = Format(CStr(Time), "hh:mm:ss")
      .Update
      'Posi��o inicial
      nOrdem = 1
      'Saldo Anterior
      dblSaldoAnterior = 0
    Else
      'Verifica se h� informa��es sobre o Caixa no dia solicitado
      .FindLast "Data <= #" & Format(Data_Atual, "mm/dd/yyyy") & "#"
      If !Data <> Data_Atual Then
        'Realiza o in�cio de dia (1� ocorr�ncia do dia)
        With tFinalizadora
          '24/06/2005 - Daniel
          '
          'Uso da fun��o nativa do VB IIf para Tratamento evitando assim o erro 94 (Invalid use of Null)
          'Esta ocorr�ncia foi registrada na empresa Barro Queimado
          .dblDinheiro = IIf(IsNumeric(rsCaixa![Total Dinheiro]), rsCaixa![Total Dinheiro], 0)
          .dblCheque = IIf(IsNumeric(rsCaixa![Total Cheques]), rsCaixa![Total Cheques], 0)
          .dblChequePre = IIf(IsNumeric(rsCaixa![Total Cheques Pr�]), rsCaixa![Total Cheques Pr�], 0)
          .dblCartao = IIf(IsNumeric(rsCaixa![Total Cart�es]), rsCaixa![Total Cart�es], 0)
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
        !Funcion�rio = nFuncionario
        !Descri��o = "In�cio do dia"
        !Dinheiro = tFinalizadora.dblDinheiro
        ![Total Dinheiro] = tFinalizadora.dblDinheiro
        !Cheques = tFinalizadora.dblCheque
        ![Total Cheques] = tFinalizadora.dblCheque
        ![Cheques Pr�] = tFinalizadora.dblChequePre
        ![Total Cheques Pr�] = tFinalizadora.dblChequePre
        !Cart�es = tFinalizadora.dblCartao
        ![Total Cart�es] = tFinalizadora.dblCartao
        !Vales = tFinalizadora.dblVale
        ![Total Vales] = tFinalizadora.dblVale
        ![Saldo Anterior] = 0
        !Final = dblSaldoAnterior
        .Update
        'Posi��o inicial
        nOrdem = 1
      Else
        'Caixa com dia j� iniciado, somente informa os valores atuais
        '28/10/2004 - Daniel
        'BUG: Tratamento para valores nulos
        With tFinalizadora
          .dblDinheiro = IIf(IsNumeric(rsCaixa.Fields("Total Dinheiro").Value), Format(rsCaixa![Total Dinheiro], FORMAT_VALUE), 0)               'rsCaixa![Total Dinheiro]
          .dblCheque = IIf(IsNumeric(rsCaixa.Fields("Total Cheques").Value), Format(rsCaixa![Total Cheques], FORMAT_VALUE), 0)                   'rsCaixa![Total Cheques]
          .dblChequePre = IIf(IsNumeric(rsCaixa.Fields("Total Cheques Pr�").Value), Format(rsCaixa![Total Cheques Pr�], FORMAT_VALUE), 0)        'rsCaixa![Total Cheques Pr�]
          .dblCartao = IIf(IsNumeric(rsCaixa.Fields("Total Cart�es").Value), Format(rsCaixa![Total Cart�es], FORMAT_VALUE), 0)                   'rsCaixa![Total Cart�es]
          .dblVale = IIf(IsNumeric(rsCaixa.Fields("Total Vales").Value), Format(rsCaixa![Total Vales], FORMAT_VALUE), 0)                         'rsCaixa![Total Vales]
          .dblParcelamento = IIf(IsNumeric(rsCaixa.Fields("Total Parcelamento").Value), Format(rsCaixa![Total Parcelamento], FORMAT_VALUE), 0)   'rsCaixa![Total Parcelamento]
        End With
        'Posi��o atual
        nOrdem = !Ordem
        'Saldo Anterior
        dblSaldoAnterior = !Final
      End If
      tValoresAtuais = tFinalizadora
    End If
    .Close
  End With
  Set rsCaixa = Nothing
  
  'Finaliza transa��o
  If blnTransaction Then ws.CommitTrans: blnInTransaction = False
  
  gbCheckOpenCaixa = True
  Exit Function

ErrHandler:
  If blnInTransaction Then ws.Rollback
  If blnByPassErr Then
    Err.Raise Err.Number, Err.Source, Err.Description
  Else
    MsgBox "Ocorreu o erro " & Err.Number & " - " & Err.Description & _
      vbCrLf & "Ao verificar a inicializa��o do Caixa [In�cio de dia].", vbCritical, "Erro"
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
  strSQL = strSQL & " AND Sequ�ncia = " & lngSequencia
  
  'Total ap�s desconto
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
      
      'Verifica possivel diferen�a e aplica no item com maior valor
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
  
  'Verificar em qual [Entradas - Produtos] h� este Produto
  strSQL = "SELECT * FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  strSQL = strSQL & " AND C�digo = '" & strCodigoProduto & "'"
  strSQL = strSQL & " AND EntradaConsignada " 'Realmente � uma consigna��o
  strSQL = strSQL & " AND NOT ConsignacaoFechada "
  'strSQL = strSQL & " AND NOT Selecionado "  o fato de ser selecionado ou n�o, n�o poder� implicar na cria��o de um novo registro na table de Acerto
  strSQL = strSQL & " AND NOT Acertado "      'N�o foi ainda acertado 100%
  strSQL = strSQL & " ORDER BY Sequ�ncia "

  Set rstEntraProd = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstEntraProd
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do Until .EOF
        'Primeira Situa��o: H� de ficar Qtde Dispon�vel
        If .Fields("QtdeAtual").Value >= dblQuantidade Then
          .Edit
          .Fields("QtdeAtual").Value = .Fields("QtdeAtual").Value - dblQuantidade
          .Update
          
          Call CriarAcerto(.Fields("Filial").Value, .Fields("Sequ�ncia").Value, Data_Atual, .Fields("Linha").Value, strCodigoProduto, dblQuantidade, bytFilial, lngSequencia, .Fields("Pre�o").Value, PrecoVenda)
          
          Exit Do
        End If
        'Segunda Situa��o: H� de zerar uma e baixar da outra
        If .Fields("QtdeAtual").Value < dblQuantidade Then
          
          dblQuantidade = dblQuantidade - .Fields("QtdeAtual").Value
          dblQtdeBaixar = .Fields("QtdeAtual").Value
          
          .Edit
          .Fields("QtdeAtual").Value = 0
          .Update
          
          Call CriarAcerto(.Fields("Filial").Value, .Fields("Sequ�ncia").Value, Data_Atual, .Fields("Linha").Value, strCodigoProduto, dblQuantidade, bytFilial, lngSequencia, .Fields("Pre�o").Value, PrecoVenda)
        
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
  'Cria��o de registros na tabela AcertoConsignacaoEntrada
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
  ' Verificar [Entradas - Produtos] se QtdeAtual est� zerada para podermos
  ' atualizar o campo [Entradas - Produtos].ConsignacaoFechada
  '-----------------------------------------------------------------------
  strSQL = "SELECT * FROM [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequ�ncia = " & sequencia
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
  ' [Entradas - Produtos].ConsignacaoFechada est� True em todas as
  ' [Entradas - Produtos] atualizaremos Entradas.ConsignacaoFechada
  '-----------------------------------------------------------------------
  strSQL = ""
  strSQL = "SELECT [Entradas - Produtos].ConsignacaoFechada "
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Entradas.Filial = " & Filial
  strSQL = strSQL & " AND Entradas.Sequ�ncia = " & sequencia
  strSQL = strSQL & " AND [Entradas - Produtos].Sequ�ncia = Entradas.Sequ�ncia "

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
  strSQL = strSQL & " AND Entradas.Sequ�ncia = " & sequencia
  
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
  'Se for dado desconto s� em um produto em uma nota com X produtos, a comiss�o
  'do vendedor em todos os �tens da nota estava caindo pela metade sendo que o
  'correto � reduzir s� a do produto que teve altera��o de pre�o ou desconto
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

Private Sub BuscarRetencao(ByVal intCodigo As Integer, ByRef dblRetencao As Double)
  '22/03/2005 - Daniel
  '
  'Case.....: Bem Me Quer
  'Projeto..: Reten��o sobre comiss�es
  Dim rstRetencao As Recordset

  dblRetencao = 0
  
  Set rstRetencao = db.OpenRecordset("SELECT ValorRetencao FROM Retencao WHERE C�digo = " & intCodigo, dbOpenSnapshot)
  
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
