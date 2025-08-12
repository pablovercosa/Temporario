Attribute VB_Name = "modFuncoes"
Option Explicit
Private rsParametros As Recordset
'10/09/2007 - Anderson
'Vari�vel utilizada para determinar o caminho do arquivo de log do sistema.
Public g_strArquivoSystemLog As String

'10/09/2007 - Anderson
'Utilizada para determinar o tipo de opera��o executada no log.
Public Enum enuSystemLog
  Inserir = 1
  Alterar = 2
  Excluir = 3
  Outros = 9
End Enum

'Anderson
'Case: Anitex
'Utilizado para realizar exporta��o de dados para o sistema da Brasil Inform�tica
Public Enum expBrasilInformaticaTipo
  Todos = 1
  Saidas = 2
  Entradas = 3
End Enum

'Anderson
'Case: Anitex
'Utilizado para realizar exporta��o de dados para o sistema da Brasil Inform�tica
Public Enum expBrasilInformaticaData
  DataEmissao = 1
  DataEntrada = 2
End Enum

'20/07/2007 - Anderson
'Case: Gurgel e Leite
'Utilizado para realizar exporta��o de dados para o sistema da Sadig Web
Public Enum expSadigWebTipo
  SaidasSadigWeb = 1
End Enum

'20/07/2007 - Anderson
'Case: Gurgel e Leite
'Utilizado para realizar exporta��o de dados para o sistema da Sadig Web
Public Enum expSadigWebData
  DataEmissaoSadigWeb = 1
  DataEntradaSadigWeb = 2
End Enum

Sub Calcula_Custo(Custo As Double, Fixa_Desc As String, Compra_Desc_V As Double, _
                  Compra_Desc_P As Double, Compra_Valor, Fixa_Frete As String, Compra_Frete_V As Double, _
                  Compra_Frete_P As Double, Fixa_ICM As String, Compra_ICM_V As Double, Compra_ICM_P As Double, _
                  Fixa_IPI As String, Compra_IPI_V As Double, Compra_IPI_P As Double, Fixa_Custo As String, _
                  Compra_Finan_V As Double, Compra_Finan_P As Double, Fixa_Outros As String, _
                  Compra_Outros_V As Double, Compra_Outros_P As Double)

  On Error GoTo ErrHandle

  Rem Desconto Fornecedor
  If Fixa_Desc = "P" Then
    Compra_Desc_V = CDbl(Compra_Desc_P) * CDbl(Compra_Valor) / 100
  Else
    If Compra_Desc_V = 0 Then
      Compra_Desc_P = 0
    Else
      Compra_Desc_P = Compra_Desc_V / Compra_Valor * 100
    End If
  End If
  
  '03/05/2013-Alexandre Afornali
  'Case Agropecuaria Colonia
  Dim rsParametros2 As Recordset
    Dim strSQL2 As String
    strSQL2 = ""
    strSQL2 = strSQL2 & "SELECT * "
    strSQL2 = strSQL2 & "FROM [Par�metros Filial] "
    strSQL2 = strSQL2 & "WHERE Filial = 1 "
    
    Set rsParametros2 = db.OpenRecordset(strSQL2, dbOpenSnapshot)
 Rem Frete Entrada
 If Fixa_Frete = "P" Then
    Compra_Frete_V = (Compra_Valor - CDbl(Compra_Desc_V)) * Compra_Frete_P / 100
 Else
    If (rsParametros2("Nome") <> "Agropecuaria Colonia") Then
        Compra_Frete_P = Compra_Frete_V / (Compra_Valor - CDbl(Compra_Desc_V)) * 100
    End If
 End If
 
 
 Rem ICMS Compra
 If Fixa_ICM = "P" Then
    Compra_ICM_V = Format(((Compra_Valor - CDbl(Compra_Desc_V)) * Compra_ICM_P / 100), "###,###,##0.00")
 Else
    Compra_ICM_P = Compra_ICM_V / (Compra_Valor - CDbl(Compra_Desc_V)) * 100
 End If
  
  
 Rem IPI Compra
 If Fixa_IPI = "P" Then
    Compra_IPI_V = Format(((Compra_Valor - CDbl(Compra_Desc_V)) * Compra_IPI_P / 100), "###,###,##0.00")
 Else
    Compra_IPI_P = Compra_IPI_V / (Compra_Valor - CDbl(Compra_Desc_V)) * 100
 End If
  
  
 Rem Custo Financeiro Compra
 If Fixa_Custo = "P" Then
    Compra_Finan_V = (Compra_Valor - CDbl(Compra_Desc_V) + CDbl(Compra_IPI_V)) * Compra_Finan_P / 100
 Else
    Compra_Finan_P = Compra_Frete_V / (Compra_Valor - CDbl(Compra_Desc_V) + CDbl(Compra_IPI_V)) * 100
 End If
 
 
 Rem Outros Compra
 If Fixa_Outros = "P" Then
    Compra_Outros_V = (Compra_Valor - CDbl(Compra_Desc_V) + CDbl(Compra_IPI_V)) * Compra_Outros_P / 100
 Else
    Compra_Outros_P = Compra_Outros_V / (Compra_Valor - CDbl(Compra_Desc_V) + CDbl(Compra_Outros_V)) * 100
 End If
 
 Custo = CDbl(Compra_Valor) - CDbl(Compra_Desc_V)
 Custo = Custo + Compra_Frete_V + Compra_IPI_V
 Custo = Custo + Compra_Finan_V + Compra_Outros_V
 
 'C_Compra_Custo.Caption = Format(Custo, "###,###,###,##0.00")
 Exit Sub
 
ErrHandle:
  gsTitle = LoadResString(201)
  gsMsg = "Erro de f�rmula. Verifique os valores informados."
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Resume Next

End Sub


Sub Verifica_Pend�ncias()
 Dim i As Integer
 Dim Data_Rec As Variant
 Dim Aux_Cliente As Long
 Dim Aux_Sequ�ncia As Long
 Dim Tem_Conta As Boolean
 Dim Tem_Cheque As Boolean
 Dim Tem_Cart�o As Boolean
 Dim Aux_Contador As Long
 Dim rsContas_Receber As Recordset
 Dim rsContas_Pagar As Recordset
 Dim rsEfetuados As Recordset
 
 
 Data_Rec = Data_Atual

 Set rsContas_Receber = db.OpenRecordset("Contas a Receber", , dbReadOnly)
 Set rsContas_Pagar = db.OpenRecordset("Contas a Pagar", , dbReadOnly)
 Set rsEfetuados = db.OpenRecordset("Contatos Efetuados", , dbReadOnly)

 Rem Zera Agenda e verifica Pend�ncias
 
 frmAgenda.lstPend.Clear
 
 i = 0
 
   Aux_Contador = 0
   rsContas_Receber.Index = "Agenda"
LP_Receber:
   rsContas_Receber.Seek ">", gnCodFilial, Data_Rec, 0, Aux_Contador
   If Not rsContas_Receber.NoMatch Then
     Aux_Contador = rsContas_Receber("Contador")
     If rsContas_Receber("Filial") = gnCodFilial Then
       If rsContas_Receber("Vencimento") = Data_Rec Then
        If rsContas_Receber("Valor Recebido") = 0 Then
           If rsContas_Receber("Tipo") = "O" Then Tem_Cart�o = True
           If rsContas_Receber("Tipo") = "C" Then Tem_Cheque = True
           If rsContas_Receber("Tipo") = "R" Then Tem_Conta = True
           GoTo LP_Receber
        End If
       End If
     End If
   End If
 
   If Tem_Cheque = True Then
      frmAgenda.lstPend.AddItem "Existem cheques pr� datados que devem ser depositados hoje."
       i = i + 1
   End If
   
   If Tem_Cart�o = True Then
       frmAgenda.lstPend.AddItem "Existem lan�amentos de cart�es de cr�dito sendo pagos hoje, n�o se esque�a de baix�-los."
       i = i + 1
   End If
   
   If Tem_Conta = True Then
       frmAgenda.lstPend.AddItem "Existem contas a receber vencendo hoje."
       i = i + 1
   End If
   

   rsContas_Pagar.Index = "Agenda"
   rsContas_Pagar.Seek ">", gnCodFilial, Data_Atual, 0, 0
   If Not rsContas_Pagar.NoMatch Then
     If rsContas_Pagar("Filial") = gnCodFilial Then
       If rsContas_Pagar("Vencimento") = Data_Atual Then
        If rsContas_Pagar("Valor Pago") = 0 Then
         frmAgenda.lstPend.AddItem "Existem contas a pagar vencendo hoje."
         i = i + 1
        End If
       End If
     End If
   End If


   rsEfetuados.Index = "Pend�ncias"
   Aux_Cliente = 0
   Aux_Sequ�ncia = 0
Lp_Pende:
   rsEfetuados.Seek ">", True, Aux_Cliente, Aux_Sequ�ncia
   If rsEfetuados.NoMatch Then GoTo Fim
   Aux_Cliente = rsEfetuados("Cliente")
   Aux_Sequ�ncia = rsEfetuados("Seq��ncia")
   If rsEfetuados("Pend�ncia") = False Then GoTo Fim
   If IsDate(rsEfetuados("Data Aviso")) Then
     If rsEfetuados("Data Aviso") = Data_Atual Then
       frmAgenda.lstPend.AddItem "Existem contatos pendentes com clientes / fornecedores para serem retomados hoje."
       i = i + 1
       GoTo Fim
     End If
   End If
   GoTo Lp_Pende


Fim:
   If Weekday(Data_Atual) = 6 Then
     frmAgenda.lstPend.AddItem "Hoje � sexta-feira, um bom dia para se fazer uma c�pia de seguran�a!!"
     i = i + 1
   End If
   

End Sub


Sub Calcula_Lucro(C_Venda_Valor As Double, _
  C_Venda_ICM_P As Double, C_Venda_ICM_V As Double, _
  C_Venda_IPI_P As Double, C_Venda_IPI_V As Double, _
  C_Venda_Imp_P As Double, C_Venda_Imp_V As Double, _
  C_Venda_Outros_V As Double, C_Venda_Outros_P As Double, _
  C_Venda_Sem_Nota As Double, C_Compra_Sem_Nota As Double, _
  Fixa_ICM_V_Perc As Boolean, Fixa_IPI_V_Perc As Boolean, _
  Fixa_Imp_V_Perc As Boolean, Fixa_Outros_V_Perc As Boolean, _
  Val_Lucro As Double, _
  C_Compra_Valor As Double, C_Compra_Desc_V As Double, _
  C_Compra_Frete_V As Double, C_Compra_Finan_V As Double, _
  C_Compra_Outros_V As Double, C_Compra_ICM_V As Double, _
  C_Compra_IPI_V As Double)

 Dim Lucro As Double
 Dim Imposto As Double
 Dim Perc_Imposto As Double
 Dim Aux As Double
 

 If Fixa_ICM_V_Perc = True Then
    C_Venda_ICM_V = Format((C_Venda_Valor * C_Venda_ICM_P / 100), "###,###,##0.00")
 Else
    If CDbl(C_Venda_Valor) = 0 Then
       C_Venda_ICM_P = 0
    Else
       C_Venda_ICM_P = C_Venda_ICM_V / C_Venda_Valor * 100
    End If
 End If
 
 
 If Fixa_IPI_V_Perc = True Then
    C_Venda_IPI_V = Format((C_Venda_Valor * C_Venda_IPI_P / 100), "###,###,##0.00")
 Else
    If CDbl(C_Venda_Valor) = 0 Then
       C_Venda_IPI_P = 0
    Else
       C_Venda_IPI_P = C_Venda_IPI_V / C_Venda_Valor * 100
    End If
 End If

 
 If Fixa_Imp_V_Perc = True Then
    C_Venda_Imp_V = C_Venda_Valor * C_Venda_Imp_P / 100
 Else
    If CDbl(C_Venda_Valor) = 0 Then
      C_Venda_Imp_P = 0
    Else
      C_Venda_Imp_P = CDbl(C_Venda_Imp_V) / CDbl(C_Venda_Valor) * 100
    End If
 End If
   
   
 If Fixa_Outros_V_Perc = True Then
    C_Venda_Outros_V = C_Venda_Valor * C_Venda_Outros_P / 100
 Else
    If CDbl(C_Venda_Valor) = 0 Then
      C_Venda_Outros_P = 0
    Else
      C_Venda_Outros_P = CDbl(C_Venda_Outros_V) / CDbl(C_Venda_Valor) * 100
    End If
 End If
 
 
 
 Aux = C_Venda_Valor * CDbl(C_Venda_ICM_P)
 Aux = Aux / 100
 C_Venda_ICM_V = Format(Aux, "########0.00")
 
 Aux = C_Venda_Valor * CDbl(C_Venda_IPI_P)
 Aux = Aux / 100
 C_Venda_IPI_V = Format(Aux, "########0.00")
 
 Aux = C_Venda_Valor * CDbl(C_Venda_Imp_P)
 Aux = Aux / 100
 C_Venda_Imp_V = Format(Aux, "########0.00")
 
 Aux = C_Venda_Valor * CDbl(C_Venda_Outros_P)
 Aux = Aux / 100
 C_Venda_Outros_V = Format(Aux, "########0.00")
 
  
 Lucro = C_Venda_Valor - CDbl(C_Compra_Valor)
 
 Lucro = Lucro + CDbl(C_Compra_Desc_V)
 
 Lucro = Lucro - CDbl(C_Compra_Frete_V)
 
 Lucro = Lucro - CDbl(C_Compra_Finan_V)
 
 Lucro = Lucro - CDbl(C_Compra_Outros_V)
 
 
 Perc_Imposto = 1 - (CDbl(C_Compra_Sem_Nota) / CDbl(100))
 Imposto = CDbl(C_Compra_ICM_V) * Perc_Imposto
 
 Lucro = Lucro + Imposto
 
 
 Imposto = CDbl(C_Compra_IPI_V) * Perc_Imposto
 
 Lucro = Lucro - Imposto
 
 If CDbl(C_Venda_IPI_V) > 0 Then
   Lucro = Lucro + Imposto
 End If
 'Else
 '  Lucro = Lucro - Imposto
 'End If


 Perc_Imposto = 1 - (CDbl(C_Venda_Sem_Nota) / CDbl(100))
 Imposto = CDbl(C_Venda_ICM_V) * Perc_Imposto

 Lucro = Lucro - Imposto
 
 Imposto = CDbl(C_Venda_IPI_V) * Perc_Imposto

 Lucro = Lucro - Imposto
 
 
 Lucro = Lucro - CDbl(C_Venda_Outros_V)
 
 Lucro = Lucro - CDbl(C_Venda_Imp_V)
 
 
 'Lucro = Format(Lucro, "#########0.00")
 
 Val_Lucro = Lucro
 
End Sub


Sub Acha_Produto(C�digo As String, Produto As String, Tamanho As Integer, _
                  Cor As Integer, Edi��o As Long, Tipo As Integer, Erro As Integer)
  Dim rsProdutos As Recordset
  Dim rsGrade As Recordset
  Dim rsEdicoes As Recordset
  Dim C�d As String
  Dim Aux_Str As String
  Dim Edic As Long
  
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsGrade = db.OpenRecordset("C�digos da Grade", , dbReadOnly)
  Set rsEdicoes = db.OpenRecordset("Edi��es", , dbReadOnly)
  
  'Tipo 0 = Produto Normal
  'Tipo 1 = Produto com Grade
  'Tipo 2 = Produto com Edi��o
  
  'erro 0 = OK
  'erro 1 = Produto n�o encontrado
  'erro 2 = Produto com grade sem tamanho e cor
  'erro 3 = Produto com edi��o sem edi��o
  'erro 4 = Produto com grade sem produto principal
  C�d = Trim(C�digo)
  
  C�digo = Trim(C�digo)
  
  rsProdutos.Index = "C�digo"
  rsProdutos.Seek "=", C�d
  If Not rsProdutos.NoMatch Then
    If rsProdutos("Tipo") = "G" Then
      Produto = 0
      Tamanho = 0
      Cor = 0
      Edi��o = 0
      Tipo = 1
      Erro = 2
      Exit Sub
    End If
    If rsProdutos("Tipo") = "E" Then
      Produto = 0
      Tamanho = 0
      Cor = 0
      Edi��o = 0
      Tipo = 2
      Erro = 3
      Exit Sub
    End If
    Produto = C�digo
    Tamanho = 0
    Cor = 0
    Edi��o = 0
    Tipo = 0
    If C�digo = "0" Then
       Erro = 1
    Else
       Erro = 0
    End If
    Exit Sub
  End If
     
  rsGrade.Index = "C�digo"
  rsGrade.Seek "=", C�digo
  
  If Not rsGrade.NoMatch Then
    C�d = rsGrade("C�digo Original")
    rsProdutos.Seek "=", C�d
    If rsProdutos.NoMatch Then
      Produto = 0
      Tamanho = 0
      Cor = 0
      Edi��o = 0
      Tipo = 1
      Erro = 4
      Exit Sub
    End If
    Produto = rsGrade("C�digo Original")
    Aux_Str = Trim(Right(C�digo, 6))
    Tamanho = Val(Left(Aux_Str, 3))
    Cor = Val(Right(Aux_Str, 3))
    Edi��o = 0
    Tipo = 1
    Erro = 0
    Exit Sub
  Else  ' Tente Edi��o
    If Len(C�digo) <> 18 Then
      Produto = 0
      Tamanho = 0
      Cor = 0
      Edi��o = 0
      Tipo = 0
      Erro = 1
      Exit Sub
    End If
  End If

  C�d = Left(C�digo, 13)
  Edic = Val(Trim(Right(C�digo, 5)))
  
  rsEdicoes.Index = "Produto"
  rsEdicoes.Seek "=", C�d, Edic
  If Not rsEdicoes.NoMatch Then
    rsProdutos.Index = "C�digo"
    rsProdutos.Seek "=", C�d
    If rsProdutos.NoMatch Then
      Produto = 0
      Tamanho = 0
      Cor = 0
      Edi��o = 0
      Tipo = 2
      Erro = 4
      Exit Sub
    End If
    Produto = C�d
    Tamanho = 0
    Cor = 0
    Edi��o = Edic
    Tipo = 2
    Erro = 0
    Exit Sub
  End If
  
  Produto = 0
  Tamanho = 0
  Cor = 0
  Edi��o = 0
  Tipo = 0
  Erro = 1

End Sub

Public Function gbProdutoComEdicao(ByVal sCodProd As String) As Boolean
  Dim rsEdicoes As Recordset
  Dim rsProdutos As Recordset
  Dim sCodPrefix As String
  Dim sCodSufix As String
  
  Set rsEdicoes = db.OpenRecordset("Edi��es", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
    
  sCodPrefix = Left(sCodProd, 13)
  sCodSufix = Val(Trim(Right(sCodProd, 5)))
  
  rsEdicoes.Index = "Produto"
  rsEdicoes.Seek "=", sCodPrefix, sCodSufix
  If Not rsEdicoes.NoMatch Then
    rsProdutos.Index = "C�digo"
    rsProdutos.Seek "=", sCodPrefix
    gbProdutoComEdicao = Not rsProdutos.NoMatch
  Else
    gbProdutoComEdicao = False
  End If

  rsEdicoes.Close
  rsProdutos.Close
  Set rsEdicoes = Nothing
  Set rsProdutos = Nothing

End Function

Public Function gbProdutoComGrade(ByVal sCodProd As String) As Boolean
  Dim rsGrade As Recordset
  
  Set rsGrade = db.OpenRecordset("C�digos da Grade", , dbReadOnly)
  
  rsGrade.Index = "C�digo"
  rsGrade.Seek "=", sCodProd
  
  gbProdutoComGrade = Not rsGrade.NoMatch

  rsGrade.Close
  Set rsGrade = Nothing
  
End Function

Public Sub gsGetDescProd(ByVal sCodProd As String, ByRef sDescricao As String, ByRef sUnidVenda As String)
  Dim rsProdutos2 As Recordset
  
  Set rsProdutos2 = rsProdutos.Clone
  
  sCodProd = Trim(sCodProd)
  
  rsProdutos2.FindFirst "C�digo = '" & sCodProd & "'"
  If Not rsProdutos2.NoMatch Then
    sDescricao = rsProdutos2("Nome") & ""
    sUnidVenda = rsProdutos2("Unidade Venda") & ""
  Else
    sDescricao = sCodProd
    sUnidVenda = ""
  End If
  
  rsProdutos2.Close
  Set rsProdutos2 = Nothing
  
End Sub

Function Arredonda_Valor(Pre�o As Double, Arredondar As String) As Double
  
  Dim Aux1 As Double
  Dim Pre�o_Str As String
  
   'Arrendondar
  '"005" = arredonda para 0.05
  '"010" = arredonda para 0.10
  '"050" = arredonda para 0.50
  '"100" = arredonda para 1.00
  '"500" = arredonda para 5.00
  '"1000" = arredonda para 10.00
    
  If Arredondar = "000" Then
    Arredonda_Valor = Pre�o
    Exit Function
  End If
    
  Pre�o_Str = Trim(Format(Pre�o, "##########0.00"))
    
  If Arredondar = "005" Then
    Aux1 = CDbl(Right(Pre�o_Str, 1))
    Aux1 = 10 - Aux1
    If Aux1 = 10 Then Aux1 = 0
    If Aux1 > 5 Then Aux1 = Aux1 - 5
    Aux1 = Aux1 / 100
    Pre�o = Pre�o + Aux1
  End If

  If Arredondar = "010" Then
    Aux1 = CDbl(Right(Pre�o_Str, 1))
    Aux1 = 10 - Aux1
    If Aux1 = 10 Then Aux1 = 0
    If Aux1 > 10 Then Aux1 = Aux1 - 10
    Aux1 = Aux1 / 100
    Pre�o = Pre�o + Aux1
  End If
  
  If Arredondar = "050" Then
    Aux1 = CDbl(Right(Pre�o_Str, 2))
    Aux1 = 100 - Aux1
    If Aux1 = 100 Then Aux1 = 0
    If Aux1 > 50 Then Aux1 = Aux1 - 50
    Aux1 = Aux1 / 100
    Pre�o = Pre�o + Aux1
  End If

  If Arredondar = "100" Then
    Aux1 = CDbl(Right(Pre�o_Str, 2))
    Aux1 = 100 - Aux1
    If Aux1 = 100 Then Aux1 = 0
    Aux1 = Aux1 / 100
    Pre�o = Pre�o + Aux1
  End If

  If Arredondar = "500" Then
    Aux1 = CDbl(Right(Pre�o_Str, 4))
    Aux1 = Aux1 * 100
    Aux1 = 1000 - Aux1
    If Aux1 = 1000 Then Aux1 = 0
    If Aux1 > 500 Then Aux1 = Aux1 - 500
    Aux1 = Aux1 / 100
    Pre�o = Pre�o + Aux1
  End If
                                 
  If Arredondar = "1000" Then
    Aux1 = CDbl(Right(Pre�o_Str, 4))
    Aux1 = Aux1 * 100
    Aux1 = 1000 - Aux1
    If Aux1 = 1000 Then Aux1 = 0
    Aux1 = Aux1 / 100
    Pre�o = Pre�o + Aux1
  End If

  Arredonda_Valor = Pre�o
  

End Function

'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado o suporte a transa��o com tratamento a erro
'Implementado a atualiza��o de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Public Sub Grava_Estoque_Final(ByVal Filial As Integer, _
      ByVal Produto As String, _
      ByVal Tamanho As Integer, _
      ByVal Cor As Integer, _
      ByVal Edi��o As Long, _
      ByVal Estoque As Single, _
      Optional Data As Date)
  
  Dim rsEstoque_Final As Recordset
  Dim rsProdutos As Recordset
  Dim blnOnTransaction As Boolean
  
  On Error GoTo ErrHandler
  
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  
  rsProdutos.Index = "C�digo"
  rsProdutos.Seek "=", Produto
  If rsProdutos.NoMatch Then
    MsgBox "Erro ao atualizar estoque final. Produto inexistente." + Produto
    rsProdutos.Close
    Set rsProdutos = Nothing
    Exit Sub
  End If
  
  Call ws.BeginTrans
  blnOnTransaction = True
  
  Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
  
  With rsEstoque_Final
    .Index = "Produto"
    .Seek "=", Filial, Produto, Tamanho, Cor, Edi��o
    If .NoMatch Then
      .AddNew
      .Fields("Filial") = Filial
      .Fields("Produto") = Produto
      .Fields("Tamanho") = Tamanho
      .Fields("Cor") = Cor
      .Fields("Edi��o") = Edi��o
    Else
      .LockEdits = True
      .Edit
    End If
    .Fields("Estoque Atual") = Estoque
    .Fields("Classe") = rsProdutos("Classe")
    .Fields("Sub Classe") = rsProdutos("Sub Classe")
     
    If IsDate(Data) Then
       .Fields("�ltima Data") = Data
    Else
       .Fields("�ltima Data") = ""
    End If
    .Update
    .LockEdits = False
    .Close
  End With

  Set rsEstoque_Final = Nothing
  
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(Produto)
  
  Call ws.CommitTrans
  blnOnTransaction = False
  
  rsProdutos.Close
  Set rsProdutos = Nothing
  Exit Sub

ErrHandler:
  If blnOnTransaction Then ws.Rollback
  'Repassa o erro para a fun��o de origem
  Err.Raise Err.Number, "Grava Estoque Final", Err.Description
  
End Sub

Public Sub Grava_Temp_Sa�das(Filial As Integer, Sequ�ncia As Long, CodProduto As String)
  Dim rsSaidas As Recordset
  Dim rsOp_Sa�das As Recordset
  Dim rsFuncionarios As Recordset
  Dim rsClientes As Recordset
  
  Dim rsSaidas_Prod As Recordset
  Dim rsProdutos As Recordset
  Dim rsSaidas_Serv As Recordset
  
  Dim rsTempo As Recordset

  Dim Aux_Contador As Long
  Dim sSql As String
  Dim Aux_Linha As Long
  Dim Tipo As String
  Dim Ordem As Integer
 
  Set rsSaidas = db.OpenRecordset("Sa�das", , dbReadOnly)
  Set rsOp_Sa�das = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsTempo = dbTemp.OpenRecordset("Sa�das")
  Set rsSaidas_Serv = db.OpenRecordset("Sa�das - Servi�os", , dbReadOnly)
  
  
  rsSaidas.Index = "Sequ�ncia"
  rsSaidas.Seek "=", Filial, Sequ�ncia
  If rsSaidas.NoMatch Then Exit Sub
  
  
  rsOp_Sa�das.Index = "C�digo"
  rsOp_Sa�das.Seek "=", rsSaidas("Opera��o")
  
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", rsSaidas("Digitador")
  
  rsClientes.Index = "C�digo"
  rsClientes.Seek "=", rsSaidas("Cliente")
  
  sSql = "DELETE * FROM Sa�das WHERE Sequ�ncia = " & str(Sequ�ncia)
  dbTemp.Execute sSql

  Ordem = 1
  
  rsSaidas_Prod.Index = "Sequ�ncia"
  Aux_Linha = 0
Lp_Prod:
  rsSaidas_Prod.Seek ">", Filial, Sequ�ncia, Aux_Linha
  If rsSaidas_Prod.NoMatch Then GoTo Ver_Servi�o
  If rsSaidas_Prod("Filial") <> Filial Then GoTo Ver_Servi�o
  If rsSaidas_Prod("Sequ�ncia") <> Sequ�ncia Then GoTo Ver_Servi�o
  
  ' ******************************* Tratamento inclu�do em Dez/2019
  ' Condi��o que traz somente vendas que contenham pelo menos um item do produto selecionado
  ' S�o 3 situa��es:
  ' SITUA��O 1: O usu�rio digitou um produto Normal                        (Exemplo 1550)
  ' SITUA��O 2: O usu�rio digitou um produto com Grade                     (Exemplo 10777001003)
  ' SITUA��O 3: O usu�rio digitou um produto com Grade, mas sem tam e cor  (Exemplo 10777)
  Dim boPula As Boolean
  boPula = True
  
  If Not IsNull(CodProduto) And Trim(CodProduto) <> "" Then
      rsProdutos.Index = "C�digo"
      rsProdutos.Seek "=", CodProduto
      If rsProdutos.NoMatch Then
          'Tenta achar se � um produto com grade
          If Len(CodProduto) > 6 Then
              rsProdutos.Index = "C�digo"
              rsProdutos.Seek "=", Mid(CodProduto, 1, Len(CodProduto) - 6)
              If rsProdutos.NoMatch = False Then
                  'Achou...era um produto com grade...
                  'ESTA � A ***** SITUA��O 2 ******
                  While Sequ�ncia = rsSaidas_Prod("Sequ�ncia") And boPula = True
                      If rsSaidas_Prod("C�digo") <> CodProduto Then
                          rsSaidas_Prod.MoveNext
                          boPula = True
                      Else
                          rsSaidas_Prod.MoveNext
                          boPula = False
                      End If
                      If rsSaidas_Prod.EOF Then
                          GoTo Ver_Servi�o
                      End If
                  Wend
              End If
          Else
              rsSaidas_Prod.MoveNext
              boPula = True
              If rsSaidas_Prod.EOF Then
                  GoTo Ver_Servi�o
              End If
          End If
      Else
          If rsProdutos.Fields("Tipo").Value = "G" Then
              'ESTA � A ***** SITUA��O 3 ******
              While Sequ�ncia = rsSaidas_Prod("Sequ�ncia") And boPula = True
                  If Len(rsSaidas_Prod("C�digo")) > 6 Then
                      If Mid(rsSaidas_Prod("C�digo"), 1, Len(rsSaidas_Prod("C�digo")) - 6) <> CodProduto Then
                          rsSaidas_Prod.MoveNext
                          boPula = True
                      Else
                          rsSaidas_Prod.MoveNext
                          boPula = False
                      End If
                      If rsSaidas_Prod.EOF Then
                          GoTo Ver_Servi�o
                      End If
                  Else
                      rsSaidas_Prod.MoveNext
                      boPula = True
                  
                      If rsSaidas_Prod.EOF Then
                          GoTo Ver_Servi�o
                      End If
                  End If
              Wend
          Else
              'ESTA � A ***** SITUA��O 1 ******
              While Sequ�ncia = rsSaidas_Prod("Sequ�ncia") And boPula = True
                  If rsSaidas_Prod("C�digo") <> CodProduto Then
                      rsSaidas_Prod.MoveNext
                      boPula = True
                  Else
                      rsSaidas_Prod.MoveNext
                      boPula = False
                  End If
                  If rsSaidas_Prod.EOF Then
                      GoTo Ver_Servi�o
                  End If
              Wend
          End If
      End If
  
      rsSaidas_Prod.MovePrevious
  End If
  If Not IsNull(CodProduto) And Trim(CodProduto) <> "" Then
      If boPula = True Then
          GoTo Ver_Servi�o
      End If
  End If
  ' fim condi��o
  ' *******************************
  
  Aux_Linha = rsSaidas_Prod("Linha")
  
  Tipo = "P"
  
  GoSub Grava_Tempo
  
  GoTo Lp_Prod
  
  
Ver_Servi�o:
  rsSaidas_Serv.Index = "Sequ�ncia"
  Aux_Linha = 0
Lp_Serv:
  rsSaidas_Serv.Seek ">", Filial, Sequ�ncia, Aux_Linha
  If rsSaidas_Serv.NoMatch Then GoTo Fim
  If rsSaidas_Serv("Filial") <> Filial Then GoTo Fim
  If rsSaidas_Serv("Sequ�ncia") <> Sequ�ncia Then GoTo Fim
  Aux_Linha = rsSaidas_Serv("Linha")
  
  Tipo = "S"
  
  GoSub Grava_Tempo
  
  GoTo Lp_Serv


Fim:
  Exit Sub
 
  
  
Grava_Tempo:
  rsTempo.AddNew
    rsTempo("Sequ�ncia") = rsSaidas("Sequ�ncia")
    rsTempo("Data") = rsSaidas("Data")
    rsTempo("C�d Opera��o") = rsSaidas("Opera��o")
    rsTempo("Nome Opera��o") = ""
    If Not rsOp_Sa�das.NoMatch Then rsTempo("Nome Opera��o") = rsOp_Sa�das("Nome") & ""
    rsTempo("Tabela") = rsSaidas("Tabela")
    rsTempo("C�d Digitador") = rsSaidas("Digitador")
    rsTempo("Nome Digitador") = ""
    If Not rsFuncionarios.NoMatch Then rsTempo("Nome Digitador") = rsFuncionarios("Nome") & ""
    rsTempo("C�d Cliente") = rsSaidas("Cliente")
    rsTempo("Nome Cliente") = ""
    If Not rsClientes.NoMatch Then rsTempo("Nome Cliente") = rsClientes("Nome") & ""
    
    rsTempo("Observa��es") = rsSaidas("Observa��es")
    rsTempo("Ref Interna") = rsSaidas("Refer�ncia")
    rsTempo("Efetivada") = rsSaidas("Efetivada")
    rsTempo("Total Produtos") = rsSaidas("Produtos")
    rsTempo("Total Desconto") = rsSaidas("Desconto")
    rsTempo("Total IPI") = rsSaidas("IPI")
    rsTempo("Total Frete") = rsSaidas("Frete")
    rsTempo("Total B ICM") = rsSaidas("Base ICM")
    rsTempo("Total ICM") = rsSaidas("Valor ICM")
    rsTempo("Total B ICM Subs") = rsSaidas("Base ICM Subs")
    rsTempo("Total ICM Subs") = rsSaidas("Valor ICM Subs")
    rsTempo("Total Nota") = rsSaidas("Total")
    rsTempo("Total Servi�os") = rsSaidas("Servi�os")
    rsTempo("Total ISS") = rsSaidas("Valor ISS")
    rsTempo("Nota") = rsSaidas("Nota Impressa")
    rsTempo("Nota Cancelada") = rsSaidas("Nota Cancelada")
    
    
    '02/09/2003 - mpdea
    'Desconto no Sub Total
    rsTempo.Fields("DescontoSubTotal").Value = rsSaidas.Fields("DescontoSubTotal").Value
    
    
    If Ordem = 1 Then
     rsTempo("Conta") = True
     Ordem = 0
    End If
    'rsTempo("Dinheiro") = ""
    'rsTempo("Cart�o") = ""
    'rsTempo("Vale") = ""
    
        
    
    
    If Tipo = "P" Then
       rsProdutos.Index = "C�digo"
       rsProdutos.Seek "=", rsSaidas_Prod("C�digo Sem Grade")
       
       rsTempo("Tipo Prod") = "P"
       rsTempo("C�digo") = rsSaidas_Prod("C�digo")
       rsTempo("Qtde") = rsSaidas_Prod("Qtde")
       If Not rsProdutos.NoMatch Then rsTempo("Nome") = rsProdutos("Nome")
       rsTempo("Pre�o") = rsSaidas_Prod("Pre�o")
       rsTempo("Desconto") = rsSaidas_Prod("Desconto")
       rsTempo("ICM") = rsSaidas_Prod("ICM")
       rsTempo("IPI") = rsSaidas_Prod("IPI")
       rsTempo("Pre�o Final") = rsSaidas_Prod("Pre�o Final")
       rsTempo("Etiqueta") = rsSaidas_Prod("Etiqueta")
       If Not rsProdutos.NoMatch Then rsTempo("Fracionado") = rsProdutos("Fracionado")
    End If
    If Tipo = "S" Then
       rsTempo("Tipo Prod") = "S"
       rsTempo("C�digo") = rsSaidas_Serv("C�digo")
       rsTempo("Qtde") = rsSaidas_Serv("Tempo")
       rsTempo("Nome") = Left(rsSaidas_Serv("Descri��o"), 50)
       rsTempo("Pre�o") = rsSaidas_Serv("Pre�o")
       rsTempo("Desconto") = 0
       rsTempo("ICM") = 0
       rsTempo("IPI") = 0
       rsTempo("Pre�o Final") = rsSaidas_Serv("Pre�o")
       rsTempo("Etiqueta") = False
       rsTempo("Fracionado") = False
    End If
    
    
    
  rsTempo.Update
    
  Return

End Sub

Sub Grava_Temp_Entradas(Filial As Integer, Sequ�ncia As Long)
  Dim rsEntradas As Recordset
  Dim rsOp_Entradas As Recordset
  Dim rsFuncionarios As Recordset
  Dim rsClientes As Recordset
  
  Dim rsEntradas_Prod As Recordset
  Dim rsProdutos As Recordset
   
  Dim rsTempo As Recordset

  Dim Aux_Contador As Long
  Dim sSql As String
  Dim Aux_Linha As Long
  Dim Tipo As String
 
  Dim Ordem As Integer
  
  Dim str_codigo As String
  Dim Str_Aux As String
  
  
  Set rsEntradas = db.OpenRecordset("Entradas", , dbReadOnly)
  Set rsOp_Entradas = db.OpenRecordset("Opera��es Entrada", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsEntradas_Prod = db.OpenRecordset("Entradas - Produtos", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsTempo = dbTemp.OpenRecordset("Entradas")
  
  rsEntradas.Index = "Sequ�ncia"
  rsEntradas.Seek "=", Filial, Sequ�ncia
  If rsEntradas.NoMatch Then Exit Sub
  
  
  sSql = "Delete * From Entradas Where Sequ�ncia = " + str(Sequ�ncia)
  dbTemp.Execute sSql
  
  Ordem = 1
  
  rsOp_Entradas.Index = "C�digo"
  rsFuncionarios.Index = "C�digo"
  rsClientes.Index = "C�digo"
  rsEntradas_Prod.Index = "Sequ�ncia"
  Aux_Linha = 0
Lp_Prod:
  rsEntradas_Prod.Seek ">", Filial, Sequ�ncia, Aux_Linha
  If rsEntradas_Prod.NoMatch Then GoTo Fim
  If rsEntradas_Prod("Filial") <> Filial Then GoTo Fim
  If rsEntradas_Prod("Sequ�ncia") <> Sequ�ncia Then GoTo Fim
  Aux_Linha = rsEntradas_Prod("Linha")
  GoSub Grava_Tempo
  
  GoTo Lp_Prod
  
Fim:
  Exit Sub
 
Grava_Tempo:
  With rsTempo
    .AddNew
    .Fields("Sequ�ncia") = rsEntradas("Sequ�ncia")
    .Fields("Data") = rsEntradas("Data")
    
    .Fields("C�d Opera��o") = rsEntradas("Opera��o")
    rsOp_Entradas.Seek "=", rsEntradas("Opera��o")
    .Fields("Nome Opera��o") = ""
    If Not rsOp_Entradas.NoMatch Then
      .Fields("Nome Opera��o") = rsOp_Entradas("Nome") & ""
    End If
        
    .Fields("C�d Digitador") = rsEntradas("Digitador")
    .Fields("Nome Digitador") = ""
    rsFuncionarios.Seek "=", rsEntradas("Digitador")
    If Not rsFuncionarios.NoMatch Then
      .Fields("Nome Digitador") = rsFuncionarios("Apelido") & ""
    End If
    
    .Fields("C�d Fornecedor") = rsEntradas("Fornecedor")
    .Fields("Nome Fornecedor") = ""
    rsClientes.Seek "=", rsEntradas("Fornecedor")
    If Not rsClientes.NoMatch Then
      .Fields("Nome Fornecedor") = rsClientes("Nome")
    End If
    
    .Fields("Observa��es") = rsEntradas("Observa��es")
    .Fields("Nota Fiscal") = rsEntradas("Nota Fiscal")
    .Fields("Pedido") = rsEntradas("Pedido")
    .Fields("Data Emiss�o") = rsEntradas("Data Emiss�o")
    .Fields("Efetivada") = rsEntradas("Efetivada")
    .Fields("Total Produtos") = rsEntradas("Produtos")
    .Fields("Total Desconto") = rsEntradas("Desconto")
    .Fields("Total IPI") = rsEntradas("IPI")
    .Fields("Total Frete") = rsEntradas("Frete")
    .Fields("Total B ICM") = rsEntradas("Base ICM")
    .Fields("Total ICM") = rsEntradas("Valor ICM")
    .Fields("Total B ICM Subs") = rsEntradas("Base ICM Subs")
    .Fields("Total ICM Subs") = rsEntradas("Valor ICM Subs")
    .Fields("Total Nota") = rsEntradas("Total")
    .Fields("Nota") = rsEntradas("Nota Impressa")
    '20/01/2004 - Daniel
    'Tratamento para os campos Entradas.CentroCusto e Entradas.NomeCentroCusto
    If IsNumeric(rsEntradas("CentroCusto").Value) Then
      .Fields("CentroCusto").Value = rsEntradas("CentroCusto").Value
      .Fields("NomeCentroCusto").Value = strGetNomeCentroCusto(rsEntradas("CentroCusto").Value) & ""
    End If
    
    rsProdutos.Index = "C�digo"
    rsProdutos.Seek "=", rsEntradas_Prod("C�digo Sem Grade")
    
    '29/10/2009 - mpdea
    'Inclu�do informa��es sobre a grade
    If Not rsProdutos.NoMatch Then
      If rsProdutos.Fields("Tipo").Value = "G" Then
        str_codigo = rsEntradas_Prod("C�digo")
        Str_Aux = str_codigo & Space(1) & gsGetNameTamanho(Mid(str_codigo, Len(str_codigo) - 5, 3))
        Str_Aux = Str_Aux & Space(1) & gsGetNameCor(Right(str_codigo, 3))
      Else
        Str_Aux = rsEntradas_Prod("C�digo")
      End If
    Else
      Str_Aux = rsEntradas_Prod("C�digo")
    End If
    .Fields("C�digo") = Str_Aux
    
    .Fields("Qtde") = rsEntradas_Prod("Qtde")
    
    If rsProdutos.NoMatch Then
      rsProdutos.Seek "=", 0
    End If
    
    '25/05/2005 - Daniel
    'rsProdutos("Nome") concatenamos com & ""
    .Fields("Nome") = rsProdutos("Nome") & ""
    .Fields("Unidade Venda") = rsProdutos("Unidade Venda")
    If Not rsProdutos.NoMatch Then .Fields("Nome") = rsProdutos("Nome")
    .Fields("Pre�o") = rsEntradas_Prod("Pre�o")
    .Fields("Desconto") = rsEntradas_Prod("Desconto")
    .Fields("ICM") = rsEntradas_Prod("ICM")
    .Fields("IPI") = rsEntradas_Prod("IPI")
    .Fields("Pre�o Final") = rsEntradas_Prod("Pre�o Final")
    .Fields("Etiqueta") = rsEntradas_Prod("Etiqueta")
    If Not rsProdutos.NoMatch Then .Fields("Fracionado") = rsProdutos("Fracionado")
    
    If Ordem = 1 Then
      .Fields("Fracionado") = True
      Ordem = 0
    Else
      .Fields("Fracionado") = False
    End If
    
    .Fields("CodUsuarioOwner") = gnUserCode
      
    .Update
    
  End With
    
  Return

End Sub


Function Inverte_Data(Data_Str As String) As String
 Dim Dia, M�s, Ano As String

 If Not IsDate(Data_Str) Then
   Inverte_Data = ""
   Exit Function
 End If
 
 Dia = Trim(str(Day(CDate(Data_Str))))
 M�s = Trim(str(Month(CDate(Data_Str))))
 Ano = Trim(str(Year(CDate(Data_Str))))
 
 Inverte_Data = M�s + "/" + Dia + "/" + Ano
 
 
End Function


Function Ajusta_Data(Dia As String) As String
 Dim Aux1 As String
 Dim Aux2 As String
 Dim Aux3 As String
 
 On Error GoTo Processa_Erro
 
 If IsNull(Dia) Then
    Ajusta_Data = ""
    Exit Function
 End If
 
 If Len(Dia) = 8 Then Dia = Dia + "  "
 
 Aux1 = Right$(Dia, 2)
 If Aux1 <> "__" And Aux1 <> "  " Then
    Ajusta_Data = Dia
    Exit Function
 End If
 Aux3 = Left$(Dia, 6)
 Aux1 = Right$(Dia, 4)
 Aux2 = Left$(Aux1, 2)
 
 If Aux2 = "96" Then
   Aux3 = Aux3 + "1996"
   Ajusta_Data = Aux3
   Exit Function
 End If
 
 If Aux2 = "97" Then
   Aux3 = Aux3 + "1997"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "98" Then
   Aux3 = Aux3 + "1998"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "99" Then
   Aux3 = Aux3 + "1999"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "00" Then
   Aux3 = Aux3 + "2000"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "01" Then
   Aux3 = Aux3 + "2001"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "02" Then
   Aux3 = Aux3 + "2002"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "03" Then
   Aux3 = Aux3 + "2003"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "04" Then
   Aux3 = Aux3 + "2004"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "05" Then
   Aux3 = Aux3 + "2005"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "06" Then
   Aux3 = Aux3 + "2006"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "07" Then
   Aux3 = Aux3 + "2007"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "08" Then
   Aux3 = Aux3 + "2008"
   Ajusta_Data = Aux3
   Exit Function
 End If

 If Aux2 = "09" Then
   Aux3 = Aux3 + "2009"
   Ajusta_Data = Aux3
   Exit Function
 End If


 Ajusta_Data = Dia
 Exit Function
 
Processa_Erro:
  Screen.MousePointer = vbDefault
  Select Case frmErro.gnShowErr(Err.Number, "Ajustar data")
    Case 0 'Repetir
      Resume
    Case 1 'Prosseguir
      Resume Next
    Case 2 'Sair
      Exit Function
    Case 3 'Encerrar
      End
  End Select

End Function


Function Acha_Estoque(Filial As Integer, Produto As String, Tamanho As Integer, Cor As Integer, Edi��o As Long, Erro As Integer) As Double
  Dim Estoque As Double
  Dim rsEstoque_Final As Recordset
  Dim rsProdutos As Recordset
  Dim bFracionado As Boolean
  Dim sMask As String
  Dim nCasasDecimais As Integer
 
  Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
  'Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  
  rsProdutos.Index = "C�digo"
  rsProdutos.Seek "=", Produto
  If Not rsProdutos.NoMatch Then
    bFracionado = rsProdutos("Fracionado")
    nCasasDecimais = gsHandleNull(rsProdutos("QtdeCasasDecimais"))
  Else
    bFracionado = False
    nCasasDecimais = 0
  End If
  
  Rem Verifica se tem estoque suficiente
  Estoque = 0
  rsEstoque_Final.Index = "Produto"
  rsEstoque_Final.Seek "=", Filial, Produto, Tamanho, Cor, Edi��o
  If Not rsEstoque_Final.NoMatch Then
    Estoque = rsEstoque_Final("Estoque Atual")
    If bFracionado Then
      sMask = String(8, "#") & "0"
      If nCasasDecimais > 0 Then
        sMask = sMask & "." & String(nCasasDecimais, "0")
      End If
      Estoque = Format(Estoque, sMask)
    End If
    Erro = 0
    GoTo Sai_fun��o
  End If
  
  
  Estoque = 0
  rsProdutos.Seek "=", Produto
  If rsProdutos.NoMatch Then
      Erro = 4
      GoTo Sai_fun��o
  End If
     
  If rsProdutos("Tipo") = "G" Then
    If Tamanho = 0 And Cor = 0 Then Erro = 2
    GoTo Sai_fun��o
  End If
     
  If rsProdutos("Tipo") = "E" Then
    If Edi��o = 0 Then Erro = 3
    GoTo Sai_fun��o
  End If
  
  Erro = 1
  

Sai_fun��o:
   Acha_Estoque = Estoque
End Function



Function �_S�_N�mero(C�digo As String) As Integer
 Dim i As Integer


  If Len(C�digo) = 0 Then
    �_S�_N�mero = True
    Exit Function
  End If
  
  
  For i = 1 To Len(C�digo)
   If Mid(C�digo, i, 1) < "0" Or Mid(C�digo, i, 1) > "9" Then
     �_S�_N�mero = False
     Exit Function
   End If
  Next i
  
  �_S�_N�mero = True
  
  

End Function

Function Acha_Estoque_Grade(Filial As Integer, Produto As String, Tamanho As Integer, Cor As Integer, Edi��o As Long, Erro As Integer) As Double
 Dim Estoque As Double
 Dim rsEstoque_Final As Recordset
 Dim rsProdutos As Recordset
 
   Set rsEstoque_Final = db.OpenRecordset("Estoque Final")
   Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
   
   
   Estoque = 0
   rsProdutos.Index = "C�digo"
   rsProdutos.Seek "=", Produto
   If rsProdutos.NoMatch Then
       Erro = 4
       GoTo Sai_fun��o
   End If
      
'   If rsProdutos("Tipo") = "G" Then
 '    If Tamanho = 0 And Cor = 0 Then Erro = 2
 '    GoTo Sai_fun��o
 '  End If
      
   If rsProdutos("Tipo") = "E" Then
     If Edi��o = 0 Then Erro = 3
     GoTo Sai_fun��o
   End If


   
   
   

   Rem Verifica se tem estoque suficiente
   Estoque = 0
   Erro = 0
   Edi��o = -1
   rsProdutos.Index = "C�digo"
   rsEstoque_Final.Index = "Produto"
Lp1:
   rsEstoque_Final.Seek ">", Filial, Produto, Tamanho, Cor, Edi��o
   If rsEstoque_Final.NoMatch Then GoTo Sai_fun��o
   If rsEstoque_Final("Filial") <> Filial Then GoTo Sai_fun��o
   If rsEstoque_Final("Produto") <> Produto Then GoTo Sai_fun��o
   
   If rsEstoque_Final("Estoque Atual") > 0 Then
      Estoque = Estoque + rsEstoque_Final("Estoque Atual")
   End If
     
   Tamanho = rsEstoque_Final("Tamanho")
   Cor = rsEstoque_Final("Cor")
   Edi��o = rsEstoque_Final("Edi��o")
     
   GoTo Lp1
   
Sai_fun��o:
   Acha_Estoque_Grade = Estoque
   
End Function

Public Function Gera_Ordena��o(ByVal sCodigo As String) As String
  Dim sAUX As String
  
  If �_S�_N�mero(sCodigo) Then
    sAUX = String(20, "+")
  Else
    sAUX = ""
  End If
  
'  sAux = sAux & C�digo
'  sAux = Trim(Right$(sAux, 20))
'  Aux = Right$(Aux, 20)
  
  Gera_Ordena��o = Trim(Right$(sAUX & sCodigo, 20))

End Function

Sub Limpa_Faturas()
 Dim i As Integer
 
 For i = 0 To 49
  Tab_Fat(i).N�mero = 0
  Tab_Fat(i).Vencimento = CDate("01/01/01")
  Tab_Fat(i).Valor = 0
 Next i
 
End Sub

Sub Limpa_Produtos()
 Dim i As Integer
 
 For i = 0 To 499
    Tab_Prod(i).C�digo = 0
    Tab_Prod(i).C�digo_Prod_Forn = ""
    Tab_Prod(i).Nome = ""
    Tab_Prod(i).C_Fiscal = ""
    Tab_Prod(i).S_Tribut�ria = ""
    Tab_Prod(i).Unid = ""
    Tab_Prod(i).Qtde = 0
    Tab_Prod(i).Valor_Unit = 0
    Tab_Prod(i).Valor_Total = 0
    Tab_Prod(i).Desconto_Perc = 0
    Tab_Prod(i).Aliq_ICM = 0
    Tab_Prod(i).Aliq_IPI = 0
    Tab_Prod(i).Valor_IPI = 0
    Tab_Prod(i).Cor = 0
    Tab_Prod(i).Nome_Cor = ""
    Tab_Prod(i).Tamanho = 0
    Tab_Prod(i).Nome_Tamanho = ""
    Tab_Prod(i).Local = ""
    Tab_Prod(i).Descr_Adicional = ""
    '27/04/2005 - Daniel
    'Campo Fabricante
    Tab_Prod(i).Fabricante = ""
 Next i

End Sub

Sub Limpa_Servi�os()
  Dim i As Integer
  
  For i = 0 To 49
    Tab_Serv(i).C�digo = 0
    Tab_Serv(i).Descri��o = ""
    Tab_Serv(i).Conclu�do = False
    Tab_Serv(i).Pre�o_Total = 0
    '27/07/2005 - Daniel
    'CST (C�digo de Situa��o Tribut�ria)
    'Finalidade: Atender a realidade da empresa W.V. Hidroan�lise Ltda (J.R. Hidroqu�mica)
    Tab_Serv(i).CST = ""
  Next i
End Sub


Function Pega_Atrasado_Cliente(Cliente As Long) As Double
' 06/12/2007 - Celso
' Fun��o para calcular o valor de contas em atraso de cliente
'
  Dim rsContas_Receber As Recordset
  Dim Total As Double
  Dim Contador As Long
  
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber", , dbReadOnly)

  Total = 0
  Contador = 0
  
  rsContas_Receber.Index = "Cliente2"

Lp1:
  rsContas_Receber.Seek ">", Cliente, Contador
  If rsContas_Receber.NoMatch Then GoTo Fim
  Contador = rsContas_Receber("Contador")
  If rsContas_Receber("Cliente") <> Cliente Then GoTo Fim
  If rsContas_Receber("Valor Recebido") <> 0 Then GoTo Lp1
  If rsContas_Receber("Vencimento") > Data_Atual Then GoTo Lp1
  If rsContas_Receber("Tipo") = "C" And rsContas_Receber("Processado") Then GoTo Lp1
  If rsContas_Receber("Tipo") = "O" Then GoTo Lp1
  '-----------------------------------------------------------------------
  
  Total = Total + rsContas_Receber("Valor") + rsContas_Receber("Acr�scimo")
  GoTo Lp1
  
  
Fim:
  Pega_Atrasado_Cliente = Total
  

End Function


Function Pega_Limite_Usado(Cliente As Long) As Double

  'Dim rsCliFor As Recordset
  Dim rsContas_Receber As Recordset
  Dim rsConta_Cliente As Recordset
  Dim Total As Double
  Dim Contador As Long
  
  'Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsContas_Receber = db.OpenRecordset("Contas a Receber", , dbReadOnly)
  Set rsConta_Cliente = db.OpenRecordset("Conta Cliente", , dbReadOnly)


  Total = 0
  Contador = 0
  
  rsContas_Receber.Index = "Cliente2"

Lp1:
  rsContas_Receber.Seek ">", Cliente, Contador
  If rsContas_Receber.NoMatch Then GoTo Pega_Conta
  Contador = rsContas_Receber("Contador")
  If rsContas_Receber("Cliente") <> Cliente Then GoTo Pega_Conta
  If rsContas_Receber("Valor Recebido") <> 0 Then GoTo Lp1
  
  '12/08/2003 - maikel
  '             Adicionada a cl�usula abaixo para que cheques que j� foram processados n�o sejam computados no c�lculo do limite de cr�dito
  If rsContas_Receber("Tipo") = "C" And rsContas_Receber("Processado") Then GoTo Lp1
  '-----------------------------------------------------------------------
  If rsContas_Receber("Tipo") = "O" And rsContas_Receber("Valor Cart�o") > 0 Then GoTo Lp1
  Total = Total + rsContas_Receber("Valor") + rsContas_Receber("Acr�scimo")
  GoTo Lp1
  
  
  
Pega_Conta:
  
  rsConta_Cliente.Index = "Cliente"
  Contador = 0

LP2:
  rsConta_Cliente.Seek ">", Cliente, Contador
  If rsConta_Cliente.NoMatch Then GoTo Fim
  Contador = rsConta_Cliente("Contador")
  If rsConta_Cliente("Cliente") <> Cliente Then GoTo Fim
  If rsConta_Cliente("Valor") = rsConta_Cliente("Valor Pago") Then GoTo LP2
  
  Total = Total + (rsConta_Cliente("Valor") - rsConta_Cliente("Valor Pago"))
  GoTo LP2
  
  
  
Fim:
  Pega_Limite_Usado = Total
  

End Function

Public Function Retira_Zeros(C�digo As String) As String
  Dim i As Integer
  Dim J As Integer
  Dim K As Integer
  
  i = Len(C�digo)
   
  If CStr(C�digo) = "0" Then
     i = 0
  End If
  
  If i = 0 Then
    Retira_Zeros = C�digo
    Exit Function
  End If
  
  For J = 1 To i
    If Left$(C�digo, 1) = "0" Then
      C�digo = Right$(C�digo, Len(C�digo) - 1)
    Else
      Retira_Zeros = C�digo
      Exit Function
    End If
  Next J
  
  
  '21/01/2004 - mpdea
  'Corrigido RT-5 ao continuar com o c�digo em branco
  If C�digo = "" Then
    Retira_Zeros = C�digo
    Exit Function
  End If
  
  
  'Procura algo diferente de numeros
  For J = 1 To i
   K = Asc(Mid(C�digo, J, 1))
   If K < 48 Or K > 57 Then
     Retira_Zeros = C�digo
     Exit Function
   End If
  Next J
End Function

Function Retorna_Valor(Texto As String)
  Dim Texto_Num As Variant
  Dim Tamanho As Integer
  Dim Pos As Integer
  Dim Letra As String
  Dim Tempo As String
  Dim Decimal1 As String
  
  Tempo = Format$(2.2, "##0.00")
  Decimal1 = Mid$(Tempo, 2, 1)
  
    
  Tamanho = Len(Texto)
  If Tamanho = 0 Then
    Retorna_Valor = 0
    Exit Function
  End If
  
  For Pos = 1 To Tamanho
    Letra = Mid(Texto, Pos, 1)
    If (Asc(Letra) >= 48 And Asc(Letra) <= 57) Or Letra = Decimal1 Or Letra = "-" Then
      If Letra <> Decimal1 Then Texto_Num = Texto_Num + Letra
      If Letra = Decimal1 Then
        Texto_Num = Texto_Num + Letra
        Decimal1 = ""
      End If
    End If
   Next Pos
 
   Retorna_Valor = CDbl(Texto_Num)
End Function

Function Verifica_DV(Texto As String) As Integer
 
 Dim DV As Long
 Dim i, J As Integer
 
  For i = 1 To Len(Texto)
    J = Asc(Mid(Texto, i, 1))
    DV = DV + J
    If DV > 255 Then DV = DV - 255
  Next i
  
  Do While DV < 161
    DV = DV + 10
  Loop

  Verifica_DV = DV

End Function

'Public Function Mostra_Erro(Erro As Integer, M�dulo As String) As Integer
'  frmErro.Num_Erro.Caption = Erro
'  frmErro.M�dulo = M�dulo
'  'frmErro.Show vbModal
'  Mostra_Erro = Val(frmErro.Retorno.Caption)
'End Function

Function Verifica_Tecla_C�digo(ByVal KeyAscii As Integer) As Integer
  If KeyAscii = 8 Then
    Verifica_Tecla_C�digo = 8  'backspace
    Exit Function
  End If
  
  If KeyAscii = 44 Then  ' ,
    Verifica_Tecla_C�digo = 45
    Exit Function
  End If
  
  If KeyAscii = 45 Then   '  -
    Verifica_Tecla_C�digo = 45
    Exit Function
  End If
  
  If KeyAscii = 46 Then   ' .
    Verifica_Tecla_C�digo = 45
    Exit Function
  End If
  
  If KeyAscii = 47 Then   ' /
    Verifica_Tecla_C�digo = 45
    Exit Function
  End If
  
  If KeyAscii = 92 Then   ' \
    Verifica_Tecla_C�digo = 45
    Exit Function
  End If
  
  If KeyAscii = 95 Then ' "_"
    Verifica_Tecla_C�digo = 45
    Exit Function
  End If
  
  If KeyAscii >= 65 And KeyAscii <= 90 Then  ' A - Z
    Verifica_Tecla_C�digo = KeyAscii
    Exit Function
  End If
  
  If KeyAscii >= 97 And KeyAscii <= 122 Then  'a - z
    Verifica_Tecla_C�digo = (KeyAscii - 32)
    Exit Function
  End If
  
  If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 13 Then Verifica_Tecla_C�digo = 0
    Exit Function
  End If
  Verifica_Tecla_C�digo = KeyAscii
  

End Function

Function Verifica_Tecla_Data(KeyAscii As Integer) As Integer
  If KeyAscii = 8 Then
    Verifica_Tecla_Data = 8  'backspace
    Exit Function
  End If
  If KeyAscii = 47 Then
    Verifica_Tecla_Data = 47  '/
    Exit Function
  End If
  
  If KeyAscii < 48 Or KeyAscii > 57 Then
    Verifica_Tecla_Data = 0
    Exit Function
  End If
  Verifica_Tecla_Data = KeyAscii
End Function

Function Verifica_Tecla_Integer(KeyAscii As Integer) As Integer
  If KeyAscii = 8 Then
    Verifica_Tecla_Integer = 8  'backspace
    Exit Function
  End If
  If KeyAscii < 48 Or KeyAscii > 57 Then
    Verifica_Tecla_Integer = 0
    Exit Function
  End If
  Verifica_Tecla_Integer = KeyAscii
End Function

Function gnGotCurrency(KeyAscii) As Integer
  Select Case KeyAscii
    Case 8, 44, 45, 46
      gnGotCurrency = KeyAscii
    Case Is < 48
      gnGotCurrency = 0
    Case Is > 57
      gnGotCurrency = 0
    Case Else
      gnGotCurrency = KeyAscii
  End Select
End Function

Function Apaga_Aspas(Texto As String) As String
  Dim Pos, Tamanho As Integer
  Dim Texto2 As String
  Dim Letra As String
  Dim Aspas As String
  
  Tamanho = Len(Texto)
  If Tamanho = 0 Then Exit Function
  
  Texto2 = ""
  Letra = ""
 Aspas = Chr(34)
  For Pos = 1 To Tamanho
    Letra = Mid(Texto, Pos, 1)
    If Letra <> Aspas Then Texto2 = Texto2 + Letra
  Next Pos
  
  Apaga_Aspas = Texto2
      
End Function

Private Function strGetNomeCentroCusto(ByVal CodCentroCusto As Integer) As String
  '20/01/2004 - Daniel
  'Tratamento para os campos Entradas.CentroCusto e Entradas.NomeCentroCusto
  'na Sub Grava_Temp_Entradas
  Dim rstCentroCusto As Recordset
  Dim strSQL         As String
  
  strSQL = "SELECT Nome FROM [Centros de Custo] "
  strSQL = strSQL & " WHERE C�digo = " & CodCentroCusto
  
  Set rstCentroCusto = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstCentroCusto
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      strGetNomeCentroCusto = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  
  Set rstCentroCusto = Nothing
  
End Function

Public Function g_blnVerificarUsoCodFornece() As Boolean
  '06/05/2005 - Daniel
  '
  'Implementa��o.: Trabalhar com o c�digo para fornecedor cadastrado na tela de produtos.
  '                Impacto: Ao entrar com o c�digo para o fornecedor no campo c�digo do produto
  '                o sistema dever� trazer o c�digo do produto que estiver amarrado nele
  'Solicita��o...: Cristiano Pavinato - PSI RS
  Dim rstParametros As Recordset
  Dim strSQL        As String
  
  On Error GoTo TratarErro
  
  strSQL = "SELECT UtilizarCodFornec FROM [Par�metros Filial] WHERE Filial = " & gnCodFilial
  
  Set rstParametros = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstParametros
    If Not (.BOF And .EOF) Then
      .MoveFirst
      g_blnVerificarUsoCodFornece = .Fields("UtilizarCodFornec").Value
    End If
    .Close
  End With

  Set rstParametros = Nothing
  
  Exit Function
  
TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  g_blnVerificarUsoCodFornece = False
  Exit Function

End Function

Public Function g_strBuscarCodProd(ByVal CodParaFornec As String) As String
  '06/05/2005 - Daniel
  '
  'Implementa��o.: Trabalhar com o c�digo para fornecedor cadastrado na tela de produtos.
  '                Impacto: Ao entrar com o c�digo para o fornecedor no campo c�digo do produto
  '                o sistema dever� trazer o c�digo do produto que estiver amarrado nele
  'Solicita��o...: Cristiano Pavinato - PSI RS
  Dim rstProdutos As Recordset
  Dim strSQL      As String
  
  On Error GoTo TratarErro
  
  g_strBuscarCodProd = ""
  
  strSQL = "SELECT C�digo, [C�digo do Fornecedor] FROM Produtos WHERE [C�digo do Fornecedor] = '" & CodParaFornec & "'"
  
  Set rstProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstProdutos.RecordCount = 0 Then
      MsgBox "Para o C�digo do Fornecedor " & CodParaFornec & " n�o h� nenhum produto vinculado.", vbExclamation, "Aten��o"
      '07/12/2006 - Anderson
      'Alterado pois causando problemas quando o c�digo do produto fornecedor era igual ao c�digo do produto
      'g_strBuscarCodProd = CodParaFornec 'Devolve o mesmo valor que veio...
      g_strBuscarCodProd = ""
      rstProdutos.Close
      Set rstProdutos = Nothing
      Exit Function
  End If
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      'Conforme an�lise sempre pegar o primeiro e '�nico' c�digo do produto
      'segundo a TI Brasil (Pavinato)
      .MoveFirst
      g_strBuscarCodProd = .Fields("C�digo").Value & ""
    End If
    .Close
  End With
  
  Set rstProdutos = Nothing
  
  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  g_strBuscarCodProd = ""
  Exit Function

End Function

'23/05/2007 - Anderson
'Exporta��o de Dados para sistema da Brasil Inform�tica
'Solicitante: Anistex Ind. e Com. Ltda (QS31935-863)
Public Function g_blnExportarDadosBrasilInformatica(ByVal Filial As Byte, ByVal DataInicial As Date, ByVal DataFinal As Date, ByVal Data As expBrasilInformaticaData, ByVal Tipo As expBrasilInformaticaTipo, ByRef ARQUIVO As String) As Boolean
On Error GoTo TratarErro:

  Dim lngArquivoSaida As Long 'Informa o n�mero do arquivo dispon�vel
  Dim strCache As String      'Recebe o valor para imprimir a linha do arquivo texto
  Dim strAux As String        'Auxilia na forma��o da string
  Dim strSQL As String        'Monta a string de consulta SQL para gera��o dos dados
  Dim intContador As Integer  'Auxilia em estruturas de repeti��o
  Dim rsSaidas As Recordset   'Abre a tabela de Saidas
  Dim rsEntradas As Recordset 'Abre a tabela de Entradas
  Dim strCaminho As String    'Informa o caminho onde deve ser salvo o arquivo
  Dim strRet As String        'Obtem retorno do arquivo ini
  Dim rsFiscal As Recordset   'Abre tabela com informa��es da impressora fiscal
  Dim rsParametros As Recordset 'Abre a tabela de Parametros da Empresa Filial
  Dim rsFiscalAnalitico As Recordset 'Abre a tabela com as informa��es detalhadas do movimento ECF
  
  lngArquivoSaida = FreeFile
  
  strCaminho = gsDefaultPath
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    'Path da aplica��o
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "SISTEMA", "ExportarDadosBrasilInformatica")
    If strRet <> "" Then strCaminho = strRet
  End If
  
  If ARQUIVO = "" Then
    ARQUIVO = strCaminho & Format(Now(), "yyyyMMddhhnnss") & ".txt"
  End If
  
  Open ARQUIVO For Output As #lngArquivoSaida
  
  '**************************************************************************************
  'Sa�das
  '**************************************************************************************
  If Tipo = Saidas Or Tipo = Todos Then
  
    strSQL = ""
    strSQL = strSQL & "SELECT [Sa�das].*, [Opera��es Sa�da].[C�digo Fiscal], [Cli_For].CGC, [Cli_For].Inscri��o, [Cli_For].Estado, [Cli_For].Cidade "
    strSQL = strSQL & "FROM ([Opera��es Sa�da] INNER JOIN Sa�das ON [Opera��es Sa�da].C�digo = Sa�das.Opera��o) "
    strSQL = strSQL & "                        INNER JOIN Cli_For ON Sa�das.Cliente = Cli_For.C�digo "
    strSQL = strSQL & "WHERE [Nota Impressa]>0 "
    strSQL = strSQL & "  AND [Nota Cancelada]=0 "
    strSQL = strSQL & "  AND [Movimenta��o Desfeita]=0 "
    
    'Verifica o filtro da Data de Emiss�o ou Data de Entrada
    If Data = DataEmissao Then
      strSQL = strSQL & "  AND [DataEmissaoNota]>=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
      strSQL = strSQL & "  AND [DataEmissaoNota]<=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
    Else
      strSQL = strSQL & "  AND Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
      strSQL = strSQL & "  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
    End If
    
    'Verifica se o filtro � por filial
    If Filial > 0 Then
      strSQL = strSQL & "  AND Filial =" & Filial & " "
    End If
    
    strSQL = strSQL & "ORDER BY Sequ�ncia "
    
    Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    Do Until rsSaidas.EOF
  
      'ORDEM CAMPO                    TP TM DC POSI��ES
      '---------------------------------------------------
      '0001  IDENTIFICADOR             C 03  0 0001 A 0003
      strCache = "LCT"
      
      '0002  NUMERO DO LAN�AMENTO      N 05  0 0004 A 0008
      strCache = strCache & Right(String(5, "0") & "" & rsSaidas("Sequ�ncia"), 5)
      
      '0003  NUMERO DA NOTA FISCAL     C 10  0 0009 A 0018
      strCache = strCache & Left("" & rsSaidas("Nota Impressa") & String(10, " "), 10)
      
      '0004  NUMERO DA NOTA FISCAL     C 10  0 0019 A 0028
      strCache = strCache & String(10, "0")
  
      '0005  ESP�CIE DA NOTA FISCAL    C 05  0 0029 A 0033
      strCache = strCache & Left("NF" & String(5, " "), 5)
      
      '0006  S�RIE DA NOTA FISCAL      C 03  0 0034 A 0036
      strCache = strCache & Left("" & rsSaidas("SerieNF") & String(3, " "), 3)
      
      '0007  DATA DE EMISS�O           D 08  0 0037 A 0044
      strCache = strCache & Right(String(8, " ") & Format("" & rsSaidas("DataEmissaoNota"), "YYYYMMDD"), 8)
      
      '0008  DATA DE ENTRADA           D 08  0 0045 A 0052
      strCache = strCache & Right(String(8, " ") & Format("" & rsSaidas("Data"), "YYYYMMDD"), 8)
      
      '0009  NATUREZA DE OPERA��O      C 06  0 0053 A 0058
      'Fun��o para retornar apenas os n�meros
      strAux = ""
      For intContador = 1 To Len("" & rsSaidas("C�digo Fiscal"))
        If IsNumeric(Mid("" & rsSaidas("C�digo Fiscal"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsSaidas("C�digo Fiscal"), intContador, 1)
        End If
      Next
      strCache = strCache & Right(String(4, " ") & Left(strAux, 4), 4) & "00"
      
      '0010  SUBSTITUI��O TRIBUTARIA   L 01  0 0059 A 0059
      If rsSaidas("Base ICM Subs") > 0 Then
        strCache = strCache & "T"
      Else
        strCache = strCache & "F"
      End If
      
      '0011  CODIGO DO FOR/CLI         C 12  0 0060 A 0071
      strCache = strCache & String(12, " ")
      
      '0012  VENDA A CONSUMIDOR        L 01  0 0072 A 0072
      strAux = ""
      For intContador = 1 To Len("" & rsSaidas("C�digo Fiscal"))
        If IsNumeric(Mid("" & rsSaidas("C�digo Fiscal"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsSaidas("C�digo Fiscal"), intContador, 1)
        End If
      Next
      If (Len("" & rsSaidas("Inscri��o")) > 0 And UCase(Trim("" & rsSaidas("Inscri��o"))) <> "ISENTO") And Left(strAux, 1) = "6" Then
        strCache = strCache & "T"
      Else
        strCache = strCache & "F"
      End If
      
      '0013  CNPJ DO FOR/CLI           C 18  0 0073 A 0090
      strAux = ""
      For intContador = 1 To Len("" & rsSaidas("CGC"))
        If IsNumeric(Mid("" & rsSaidas("CGC"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsSaidas("CGC"), intContador, 1)
        End If
      Next
      
      strAux = String(14, "0") & strAux
      
      strCache = strCache & Left(Right(strAux, 14), 2) & "." & _
                            Left(Right(strAux, 12), 3) & "." & _
                            Left(Right(strAux, 9), 3) & "/" & _
                            Left(Right(strAux, 6), 4) & "-" & _
                            Right(strAux, 2)
      
      '0014  CONTA A  D�BITO           N 05  0 0091 A 0095
      strCache = strCache & String(5, " ")
      
      '0015  CONTA A CR�DITO           N 05  0 0096 A 0100
      strCache = strCache & String(5, " ")
      
      '0016  C.CUSTO A D�BITO          N 06  0 0101 A 0106
      strCache = strCache & String(6, " ")
      
      '0017  C.CUSTO A CR�DITO         N 06  0 0107 A 0112
      strCache = strCache & String(6, " ")
      
      '0018  SETOR  A D�BITO           N 06  0 0113 A 0118
      strCache = strCache & String(6, " ")
      
      '0019  SETOR A CR�DITO           N 06  0 0119 A 0124
      strCache = strCache & String(6, " ")
      
      '0020  VALOR  TOTAL DA NF        N 16  2 0125 A 0140
      strAux = Replace(Right(Format("" & rsSaidas("Total"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0021  VALOR CONT�BIL            N 16  2 0141 A 0156
      strCache = strCache & "0000000000000.00"
      
      '0022  UF                        C 03  0 0157 A 0159
      strCache = strCache & Right(String(3, " ") & UCase("" & rsSaidas("Estado")), 3)
      
      '0023  C�DIGO DO SERVI�O         N 09  0 0160 A 0168
      strCache = strCache & String(9, " ")
      
      '0024  VALOR DOS MATERIAIS       N 16  2 0169 A 0184
      strCache = strCache & "0000000000000.00"
      
      '0025  VALOR SUB-EMPREITADA      N 16  2 0185 A 0200
      strCache = strCache & "0000000000000.00"
      
      '0026  ROTINA DE CALCULO  1      N 05  0 0201 A 0205
      strCache = strCache & String(5, "0")
      
      '0027  BSC.ICMS  1               N 16  2 0206 A 0221
      strAux = Replace(Right(Format("" & rsSaidas("Base ICM"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0028  ALIQUOTA DE ICMS  1       N 05  2 0222 A 0226
      If rsSaidas("Valor ICM") > 0 And rsSaidas("Base ICM") > 0 Then
        strCache = strCache & Right(Format(Int(((rsSaidas("Valor ICM") / rsSaidas("Base ICM")) * 100)), "00.00"), 5)
      Else
        strCache = strCache & "00.00"
      End If
      
      '0029  VALOR DO ICMS  1          N 16  2 0227 A 0242
      strAux = Replace(Right(Format(rsSaidas("Valor ICM"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0030  VALOR ISENTA DE ICM  1    N 16  2 0243 A 0258
      strCache = strCache & "0000000000000.00"
      
      '0031  VALOR OUTRAS DE ICM 1     N 16  2 0259 A 0274
      strCache = strCache & "0000000000000.00"
      
      '0032  ROTINA DE CALCULO  2      N 05  0 0275 A 0279
      strCache = strCache & String(5, "0")
      
      '0033  BSC.ICMS  2               N 16  2 0280 A 0295
      strCache = strCache & "0000000000000.00"
      
      '0034  ALIQUOTA DE ICMS  2       N 05  2 0296 A 0300
      strCache = strCache & "00.00"
      
      '0035  VALOR DO ICMS  2          N 16  2 0301 A 0316
      strCache = strCache & "0000000000000.00"
      
      '0036  VALOR ISENTA DE ICM  2    N 16  2 0317 A 0332
      strCache = strCache & "0000000000000.00"
      
      '0037  VALOR OUTRAS DE ICM 2     N 16  2 0333 A 0348
      strCache = strCache & "0000000000000.00"
      
      '0038  ROTINA DE CALCULO  3      N 05  0 0349 A 0353
      strCache = strCache & String(5, "0")
      
      '0039  BSC.ICMS  3               N 16  2 0354 A 0369
      strCache = strCache & "0000000000000.00"
      
      '0040  ALIQUOTA DE ICMS  3       N 05  2 0370 A 0374
      strCache = strCache & "00.00"
      
      '0041  VALOR DO ICMS  3          N 16  2 0375 A 0390
      strCache = strCache & "0000000000000.00"
      
      '0042  VALOR ISENTA DE ICM  3    N 16  2 0391 A 0406
      strCache = strCache & "0000000000000.00"
      
      '0043  VALOR OUTRAS DE ICM 3     N 16  2 0407 A 0422
      strCache = strCache & "0000000000000.00"
      
      '0044  ROTINA DE CALCULO  4      N 05  0 0423 A 0427
      strCache = strCache & String(5, "0")
      
      '0045  BSC.ICMS  4               N 16  2 0428 A 0443
      strCache = strCache & "0000000000000.00"
      
      '0046  ALIQUOTA DE ICMS  4       N 05  2 0444 A 0448
      strCache = strCache & "00.00"
      
      '0047  VALOR DO ICMS  4          N 16  2 0449 A 0464
      strCache = strCache & "0000000000000.00"
      
      '0048  VALOR ISENTA DE ICM  4    N 16  2 0465 A 0480
      strCache = strCache & "0000000000000.00"
      
      '0049  VALOR OUTRAS DE ICM 4     N 16  2 0481 A 0496
      strCache = strCache & "0000000000000.00"
      
      '0050  ROTINA DE CALCULO  5      N 05  0 0497 A 0501
      strCache = strCache & String(5, "0")
      
      '0051  BSC.ICMS  5               N 16  2 0502 A 0517
      strCache = strCache & "0000000000000.00"
      
      '0052  ALIQUOTA DE ICMS  5       N 05  2 0518 A 0522
      strCache = strCache & "00.00"
      
      '0053  VALOR DO ICMS  5          N 16  2 0523 A 0538
      strCache = strCache & "0000000000000.00"
      
      '0054  VALOR ISENTA DE ICM  5    N 16  2 0539 A 0554
      strCache = strCache & "0000000000000.00"
      
      '0055  VALOR OUTRAS DE ICM 5     N 16  2 0555 A 0570
      strCache = strCache & "0000000000000.00"
      
      '0056  VLR TOTAL DA BSC.IPI      N 16  2 0571 A 0586
      strCache = strCache & "0000000000000.00"
      
      '0057  PERCENTUAL DE IPI         N 05  2 0587 A 0591
      strCache = strCache & "00.00"
      
      '0058  VALOR DO IPI              N 16  2 0592 A 0607
      strAux = Replace(Right(Format("" & rsSaidas("IPI"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0059  VALOR DE ISENTAS DE IPI   N 16  2 0608 A 0623
      strCache = strCache & "0000000000000.00"
      
      '0060  VALOR DE OUTRAS DE IPI    N 16  2 0624 A 0639
      strCache = strCache & "0000000000000.00"
      
      '0061  PERC.  I.R.  S/ SERVI�OS  N 05  2 0640 A 0644
      strAux = Replace(Right(Format("" & rsSaidas("Perc IR Sobre ISS"), "00.00"), 5), ",", ".")
      strCache = strCache & strAux
      
      '0062  ICMS RETIDO NA FONTE      N 16  2 0645 A 0660
      strAux = Replace(Right(Format("" & rsSaidas("Total IRRF"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0063  OBS. DE IPI               N 16  2 0661 A 0676
      strCache = strCache & "0000000000000.00"
      
      '0064  BASE DE CALCULO INSS      N 16  2 0677 A 0692
      strCache = strCache & "0000000000000.00"
      
      '0065  PERCENTUAL DO INSS        N 05  2 0693 A 0697
      strCache = strCache & "00.00"
      
      '0066  VALOR DO INSS             N 16  2 0698 A 0713
      strCache = strCache & "0000000000000.00"
      
      '0067  BASE DE CALC. S. TRIB.    N 16  2 0714 A 0729
      strAux = Replace(Right(Format("" & rsSaidas("Base ICM Subs"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0068  VALOR DA S.TRIBUTARIA     N 16  2 0730 A 0745
      strAux = Replace(Right(Format("" & rsSaidas("Valor ICM Subs"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0069  CODIGO DA ZFM             C 05  0 0746 A 0750
      Select Case UCase(Mid("" & rsSaidas("Estado"), 1, 2))
      
        Case "AM"
        
          Select Case UCase("" & rsSaidas("Cidade"))
            Case "MANAUS": strCache = strCache & "00255"
            Case "PRESIDENTE FIGUEIREDO": strCache = strCache & "09841"
            Case "RIO PRETO DA EVA": strCache = strCache & "09843"
            Case "TABATINGA": strCache = strCache & "09847"
            Case Else: strCache = strCache & "0000"
          End Select
        
        Case "AC"
        
          Select Case UCase("" & rsSaidas("Cidade"))
            Case "BRASIL�IA": strCache = strCache & "00105"
            Case "CRUZEIRO DO SUL": strCache = strCache & "00107"
            Case "EPIT�CIOLANDIA": strCache = strCache & "99998"
            Case Else: strCache = strCache & "00000"
          End Select
        
        Case "AP"
        
          Select Case UCase("" & rsSaidas("Cidade"))
            Case "MACAPA": strCache = strCache & "00605"
            Case "SANTANA": strCache = strCache & "00615"
            Case Else: strCache = strCache & "00000"
          End Select
          
        Case "RO"
        
          Select Case UCase("" & rsSaidas("Cidade"))
            Case "GUARAJA MIRIM": strCache = strCache & "00001"
            Case Else: strCache = strCache & "00000"
          End Select
        
        Case "RR"
          Select Case UCase("" & rsSaidas("Cidade"))
            Case "BONFIM": strCache = strCache & "00307"
            Case "PACARAIMA": strCache = strCache & "99999"
            Case Else: strCache = strCache & "00000"
          End Select
          
        Case Else: strCache = strCache & "00000"
      
      End Select
      
      '0070  OBSERVA��ES NECESSARIAS   C 40  0 0751 A 0790
      strCache = strCache & Left("" & rsSaidas("Observa��es") & String(40, " "), 40)
      
      '0071  FLAG DE ATUALIZA��O       L 01  0 0791 A 0791
      strCache = strCache & "F"
      
      '0072  NUMERO DA ESTA��O         C 03  0 0792 A 0794
      strCache = strCache & "001"
      
      '0073  OBSERVA��O 2              C 40  0 0795 A 0834
      strCache = strCache & String(40, " ")
      
      '0074  OBSERVA��O 3              C 40  0 0835 A 0874
      strCache = strCache & String(40, " ")
      
      '0075  CIF_FOB                   C 01  0 0875 A 0875
      If "" & rsSaidas("obs_FretePago") <> 2 Then
        strCache = strCache & "1"
      Else
        strCache = strCache & "2"
      End If
      
      '0076  SITNOTA                   C 01  0 0876 A 0876
      strCache = strCache & " "
      
      '0077  BSCISSRET                 N 16  2 0877 A 0892
      strAux = Replace(Right(Format("" & rsSaidas("Base ISS"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0078  VLRISSRET                 N 16  2 0893 A 0908
      strAux = Replace(Right(Format("" & rsSaidas("Valor ISS"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0079  ALQISSRET                 N  5  2 0909 A 0913
      strAux = Replace(Right(Format("" & rsSaidas("Perc IR Sobre ISS"), "00.00"), 16), ",", ".")
      strCache = strCache & strAux
    
      Print #lngArquivoSaida, strCache
      strCache = ""
      rsSaidas.MoveNext
      
    Loop
    
    rsSaidas.Close
    
    '26/10/2007 - Anderson
    'Implementa��o de ECF
    '**************************************************************************************
    'Sa�das - Cupom Fiscal
    '**************************************************************************************
    '      0001    IDENTIFICADOR           C 03  0 0001 a 0003
    '      0002    NUMERO DO LANCAMENTO    N 05  0 0004 a 0008
    '      0003    CODIGO DA MAQUINA REG   C 03  0 0009 a 0011
    '      0004    VLR. CANCELAMENTOS      N 16  2 0012 a 0027
    '      0005    VLR. DESCONTOS          N 16  2 0028 a 0043
    '      0006    VLR ISS                 N 16  2 0044 a 0059
    '      0007    LEITURA Z               N 06  0 0060 a 0065
    '      0008    CUPOM INICIAL           C 06  0 0066 a 0071
    '      0009    CUPOM FINAL             C 06  0 0072 a 0077
    '      0010    CRO                     N 03  0 0078 A 0080
   
    strSQL = ""
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM FISReg60Mestre "
    strSQL = strSQL & "WHERE Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
    strSQL = strSQL & "  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
    
    'Verifica se o filtro � por filial
    If Filial > 0 Then
      strSQL = strSQL & "  AND Filial =" & Filial & " "
    End If
    
    Set rsFiscal = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    strSQL = ""
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM [Par�metros Filial] "
    strSQL = strSQL & "WHERE Filial =" & Filial & " "
    
    Set rsParametros = db.OpenRecordset(strSQL, dbOpenSnapshot)

    Do Until rsFiscal.EOF
    
      strSQL = ""
      strSQL = strSQL & "SELECT Sa�das.Filial, Sa�das.Data, Sum(Sa�das.[Base ISS]) AS SomaDeBaseISS, Sum(Sa�das.[Valor ISS]) AS SomaDeValorISS, Sum(Sa�das.[Base ICM]) AS SomaDeBaseICM, Sum(Sa�das.[Valor ICM]) AS SomaDeValorICM, Sum(Sa�das.Total) AS SomaDeTotal, Sum(Sa�das.[Perc IR Sobre ISS]) AS SomaDePercIRSobreISS, Sum(Sa�das.[Base ICM Subs]) AS SomaDeBaseICMSubs, Sum(Sa�das.[Total IRRF]) AS SomaDeTotalIRRF, Sum(Sa�das.[Valor ICM Subs]) AS SomaDeValorICMSubs, [Opera��es Sa�da].[C�digo Fiscal] "
      strSQL = strSQL & "FROM [Opera��es Sa�da] INNER JOIN Sa�das ON [Opera��es Sa�da].C�digo = Sa�das.Opera��o "
      strSQL = strSQL & "GROUP BY Sa�das.Filial, Sa�das.Data, [Opera��es Sa�da].[C�digo Fiscal], Sa�das.[Cupom Fiscal Impresso], Sa�das.[Nota Cancelada], Sa�das.[Movimenta��o Desfeita], Sa�das.Data, Sa�das.Filial "
      strSQL = strSQL & "HAVING [Cupom Fiscal Impresso]<>0 "
      strSQL = strSQL & "  AND [Nota Cancelada]=0 "
      strSQL = strSQL & "  AND [Movimenta��o Desfeita]=0 "
      strSQL = strSQL & "  AND Data =#" & Format(rsFiscal("Data"), "mm/dd/yyyy") & "# "
      strSQL = strSQL & "  AND Filial =" & rsFiscal("Filial") & " "

      Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
      
      If Not rsSaidas.EOF Then

        'ORDEM CAMPO                    TP TM DC POSI��ES
        '---------------------------------------------------
        '0001  IDENTIFICADOR             C 03  0 0001 A 0003
        strCache = "LCT"
        
        '0002  NUMERO DO LAN�AMENTO      N 05  0 0004 A 0008
        strCache = strCache & "00000"
        
        '0003  NUMERO DA NOTA FISCAL     C 10  0 0009 A 0018
        strCache = strCache & Left("" & rsFiscal("NrCOOInicioDia") & String(10, " "), 10)
        
        '0004  NUMERO DA NOTA FISCAL     C 10  0 0019 A 0028
        strCache = strCache & Left("" & rsFiscal("NrCOOFimDia") & String(10, " "), 10)
    
        '0005  ESP�CIE DA NOTA FISCAL    C 05  0 0029 A 0033
        strCache = strCache & Left("CF" & String(5, " "), 5)
        
        '0006  S�RIE DA NOTA FISCAL      C 03  0 0034 A 0036
        strCache = strCache & Left("ECF" & String(3, " "), 3)

        '0007  DATA DE EMISS�O           D 08  0 0037 A 0044
        strCache = strCache & Right(String(8, " ") & Format("" & rsFiscal("Data"), "YYYYMMDD"), 8)

        '0008  DATA DE ENTRADA           D 08  0 0045 A 0052
        strCache = strCache & Right(String(8, " ") & Format("" & rsFiscal("Data"), "YYYYMMDD"), 8)

        '0009  NATUREZA DE OPERA��O      C 06  0 0053 A 0058
        'Fun��o para retornar apenas os n�meros
        strAux = ""
        For intContador = 1 To Len("" & rsSaidas("C�digo Fiscal"))
          If IsNumeric(Mid("" & rsSaidas("C�digo Fiscal"), intContador, 1)) Then
            strAux = strAux & Mid("" & rsSaidas("C�digo Fiscal"), intContador, 1)
          End If
        Next
        strCache = strCache & Right(String(4, " ") & Left(strAux, 4), 4) & "00"

        '0010  SUBSTITUI��O TRIBUTARIA   L 01  0 0059 A 0059
        If rsSaidas("SomaDeBaseICMSubs") > 0 Then
          strCache = strCache & "T"
        Else
          strCache = strCache & "F"
        End If

        '0011  CODIGO DO FOR/CLI         C 12  0 0060 A 0071
        strCache = strCache & String(12, " ")

        '0012  VENDA A CONSUMIDOR        L 01  0 0072 A 0072
        strCache = strCache & "T"

        '0013  CNPJ DO FOR/CLI           C 18  0 0073 A 0090
        strAux = String(14, "0")

        strCache = strCache & Left(Right(strAux, 14), 2) & "." & _
                              Left(Right(strAux, 12), 3) & "." & _
                              Left(Right(strAux, 9), 3) & "/" & _
                              Left(Right(strAux, 6), 4) & "-" & _
                              Right(strAux, 2)

        '0014  CONTA A  D�BITO           N 05  0 0091 A 0095
        strCache = strCache & String(5, " ")

        '0015  CONTA A CR�DITO           N 05  0 0096 A 0100
        strCache = strCache & String(5, " ")

        '0016  C.CUSTO A D�BITO          N 06  0 0101 A 0106
        strCache = strCache & String(6, " ")

        '0017  C.CUSTO A CR�DITO         N 06  0 0107 A 0112
        strCache = strCache & String(6, " ")

        '0018  SETOR  A D�BITO           N 06  0 0113 A 0118
        strCache = strCache & String(6, " ")

        '0019  SETOR A CR�DITO           N 06  0 0119 A 0124
        strCache = strCache & String(6, " ")

        '0020  VALOR  TOTAL DA NF        N 16  2 0125 A 0140
        strAux = Replace(Right(Format("" & rsSaidas("SomaDeTotal"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux

        '0021  VALOR CONT�BIL            N 16  2 0141 A 0156
        strCache = strCache & "0000000000000.00"

        '0022  UF                        C 03  0 0157 A 0159
        strCache = strCache & Right(String(3, " ") & UCase("" & rsParametros("Estado")), 3)

        '0023  C�DIGO DO SERVI�O         N 09  0 0160 A 0168
        strCache = strCache & String(9, " ")

        '0024  VALOR DOS MATERIAIS       N 16  2 0169 A 0184
        strCache = strCache & "0000000000000.00"

        '0025  VALOR SUB-EMPREITADA      N 16  2 0185 A 0200
        strCache = strCache & "0000000000000.00"

        '0026  ROTINA DE CALCULO  1      N 05  0 0201 A 0205
        strCache = strCache & String(5, "0")

        '0027  BSC.ICMS  1               N 16  2 0206 A 0221
        strAux = Replace(Right(Format("" & rsSaidas("SomaDeBaseICM"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux

        '0028  ALIQUOTA DE ICMS  1       N 05  2 0222 A 0226
        If rsSaidas("SomaDeValorICM") > 0 And rsSaidas("SomaDeBaseICM") > 0 Then
          strCache = strCache & Right(Format(Int(((rsSaidas("SomaDeValorICM") / rsSaidas("SomaDeBaseICM")) * 100)), "00.00"), 5)
        Else
          strCache = strCache & "00.00"
        End If

        '0029  VALOR DO ICMS  1          N 16  2 0227 A 0242
        strAux = Replace(Right(Format(rsSaidas("SomaDeValorICM"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux

        '0030  VALOR ISENTA DE ICM  1    N 16  2 0243 A 0258
        strCache = strCache & "0000000000000.00"

        '0031  VALOR OUTRAS DE ICM 1     N 16  2 0259 A 0274
        strCache = strCache & "0000000000000.00"

        '0032  ROTINA DE CALCULO  2      N 05  0 0275 A 0279
        strCache = strCache & String(5, "0")

        '0033  BSC.ICMS  2               N 16  2 0280 A 0295
        strCache = strCache & "0000000000000.00"

        '0034  ALIQUOTA DE ICMS  2       N 05  2 0296 A 0300
        strCache = strCache & "00.00"

        '0035  VALOR DO ICMS  2          N 16  2 0301 A 0316
        strCache = strCache & "0000000000000.00"

        '0036  VALOR ISENTA DE ICM  2    N 16  2 0317 A 0332
        strCache = strCache & "0000000000000.00"

        '0037  VALOR OUTRAS DE ICM 2     N 16  2 0333 A 0348
        strCache = strCache & "0000000000000.00"

        '0038  ROTINA DE CALCULO  3      N 05  0 0349 A 0353
        strCache = strCache & String(5, "0")

        '0039  BSC.ICMS  3               N 16  2 0354 A 0369
        strCache = strCache & "0000000000000.00"

        '0040  ALIQUOTA DE ICMS  3       N 05  2 0370 A 0374
        strCache = strCache & "00.00"

        '0041  VALOR DO ICMS  3          N 16  2 0375 A 0390
        strCache = strCache & "0000000000000.00"

        '0042  VALOR ISENTA DE ICM  3    N 16  2 0391 A 0406
        strCache = strCache & "0000000000000.00"

        '0043  VALOR OUTRAS DE ICM 3     N 16  2 0407 A 0422
        strCache = strCache & "0000000000000.00"

        '0044  ROTINA DE CALCULO  4      N 05  0 0423 A 0427
        strCache = strCache & String(5, "0")

        '0045  BSC.ICMS  4               N 16  2 0428 A 0443
        strCache = strCache & "0000000000000.00"

        '0046  ALIQUOTA DE ICMS  4       N 05  2 0444 A 0448
        strCache = strCache & "00.00"

        '0047  VALOR DO ICMS  4          N 16  2 0449 A 0464
        strCache = strCache & "0000000000000.00"

        '0048  VALOR ISENTA DE ICM  4    N 16  2 0465 A 0480
        strCache = strCache & "0000000000000.00"

        '0049  VALOR OUTRAS DE ICM 4     N 16  2 0481 A 0496
        strCache = strCache & "0000000000000.00"

        '0050  ROTINA DE CALCULO  5      N 05  0 0497 A 0501
        strCache = strCache & String(5, "0")

        '0051  BSC.ICMS  5               N 16  2 0502 A 0517
        strCache = strCache & "0000000000000.00"

        '0052  ALIQUOTA DE ICMS  5       N 05  2 0518 A 0522
        strCache = strCache & "00.00"

        '0053  VALOR DO ICMS  5          N 16  2 0523 A 0538
        strCache = strCache & "0000000000000.00"

        '0054  VALOR ISENTA DE ICM  5    N 16  2 0539 A 0554
        strCache = strCache & "0000000000000.00"

        '0055  VALOR OUTRAS DE ICM 5     N 16  2 0555 A 0570
        strCache = strCache & "0000000000000.00"

        '0056  VLR TOTAL DA BSC.IPI      N 16  2 0571 A 0586
        strCache = strCache & "0000000000000.00"

        '0057  PERCENTUAL DE IPI         N 05  2 0587 A 0591
        strCache = strCache & "00.00"

        '0058  VALOR DO IPI              N 16  2 0592 A 0607
        strAux = "0000000000000.00"
        strCache = strCache & strAux

        '0059  VALOR DE ISENTAS DE IPI   N 16  2 0608 A 0623
        strCache = strCache & "0000000000000.00"

        '0060  VALOR DE OUTRAS DE IPI    N 16  2 0624 A 0639
        strCache = strCache & "0000000000000.00"

        '0061  PERC.  I.R.  S/ SERVI�OS  N 05  2 0640 A 0644
        strAux = "00.00"
        strCache = strCache & strAux

        '0062  ICMS RETIDO NA FONTE      N 16  2 0645 A 0660
        strAux = Replace(Right(Format("0" & rsSaidas("SomaDeTotalIRRF"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux

        '0063  OBS. DE IPI               N 16  2 0661 A 0676
        strCache = strCache & "0000000000000.00"

        '0064  BASE DE CALCULO INSS      N 16  2 0677 A 0692
        strCache = strCache & "0000000000000.00"

        '0065  PERCENTUAL DO INSS        N 05  2 0693 A 0697
        strCache = strCache & "00.00"

        '0066  VALOR DO INSS             N 16  2 0698 A 0713
        strCache = strCache & "0000000000000.00"

        '0067  BASE DE CALC. S. TRIB.    N 16  2 0714 A 0729
        strAux = Replace(Right(Format("" & rsSaidas("SomaDeBaseICMSubs"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux

        '0068  VALOR DA S.TRIBUTARIA     N 16  2 0730 A 0745
        strAux = Replace(Right(Format("" & rsSaidas("SomaDeValorICMSubs"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux

        '0069  CODIGO DA ZFM             C 05  0 0746 A 0750
        Select Case UCase(Mid("" & rsParametros("Estado"), 1, 2))

          Case "AM"

            Select Case UCase("" & rsParametros("Cidade"))
              Case "MANAUS": strCache = strCache & "00255"
              Case "PRESIDENTE FIGUEIREDO": strCache = strCache & "09841"
              Case "RIO PRETO DA EVA": strCache = strCache & "09843"
              Case "TABATINGA": strCache = strCache & "09847"
              Case Else: strCache = strCache & "0000"
            End Select

          Case "AC"

            Select Case UCase("" & rsParametros("Cidade"))
              Case "BRASIL�IA": strCache = strCache & "00105"
              Case "CRUZEIRO DO SUL": strCache = strCache & "00107"
              Case "EPIT�CIOLANDIA": strCache = strCache & "99998"
              Case Else: strCache = strCache & "00000"
            End Select

          Case "AP"

            Select Case UCase("" & rsParametros("Cidade"))
              Case "MACAPA": strCache = strCache & "00605"
              Case "SANTANA": strCache = strCache & "00615"
              Case Else: strCache = strCache & "00000"
            End Select

          Case "RO"

            Select Case UCase("" & rsParametros("Cidade"))
              Case "GUARAJA MIRIM": strCache = strCache & "00001"
              Case Else: strCache = strCache & "00000"
            End Select

          Case "RR"
            Select Case UCase("" & rsParametros("Cidade"))
              Case "BONFIM": strCache = strCache & "00307"
              Case "PACARAIMA": strCache = strCache & "99999"
              Case Else: strCache = strCache & "00000"
            End Select

          Case Else: strCache = strCache & "00000"

        End Select

        '0070  OBSERVA��ES NECESSARIAS   C 40  0 0751 A 0790
        strCache = strCache & String(40, " ")

        '0071  FLAG DE ATUALIZA��O       L 01  0 0791 A 0791
        strCache = strCache & "F"

        '0072  NUMERO DA ESTA��O         C 03  0 0792 A 0794
        strCache = strCache & "001"

        '0073  OBSERVA��O 2              C 40  0 0795 A 0834
        strCache = strCache & String(40, " ")

        '0074  OBSERVA��O 3              C 40  0 0835 A 0874
        strCache = strCache & String(40, " ")

        '0075  CIF_FOB                   C 01  0 0875 A 0875
        strCache = strCache & "0"

        '0076  SITNOTA                   C 01  0 0876 A 0876
        strCache = strCache & " "

        '0077  BSCISSRET                 N 16  2 0877 A 0892
        strAux = Replace(Right(Format("" & rsSaidas("SomaDeBaseISS"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux

        '0078  VLRISSRET                 N 16  2 0893 A 0908
        strAux = Replace(Right(Format("" & rsSaidas("SomaDeValorISS"), "0000000000000.00"), 16), ",", ".")
        strCache = strCache & strAux
        
        '0079  ALQISSRET                 N  5  2 0909 A 0913
        strAux = "00.00"
        strCache = strCache & strAux
    
        Print #lngArquivoSaida, strCache
        strCache = ""
        
        'ORDEM CAMPO                    TP TM DC POSI��ES
        '---------------------------------------------------
        '0001    IDENTIFICADOR           C 03  0 0001 a 0003
        strCache = "ECF"
        
        '0002    NUMERO DO LANCAMENTO    N 05  0 0004 a 0008
        strCache = strCache & "00000"
        
        '0003    CODIGO DA MAQUINA REG   C 03  0 0009 a 0011
        strCache = strCache & Right(String(3, " ") & UCase("" & rsFiscal("NrECF")), 3)
        
        '0004    VLR. CANCELAMENTOS      N 16  2 0012 a 0027
        strSQL = ""
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM FISReg60Analitico "
        strSQL = strSQL & "WHERE Filial =" & Filial & " "
        strSQL = strSQL & "  AND Data =#" & Format(rsFiscal("Data"), "mm/dd/yyyy") & "# "
        strSQL = strSQL & "  AND ST_Aliquota ='CANC'"
        Set rsFiscalAnalitico = db.OpenRecordset(strSQL, dbOpenSnapshot)
        If Not rsFiscalAnalitico.EOF Then
          strCache = strCache & Replace(Right(Format("" & rsFiscalAnalitico("VlrAcumulado"), "0000000000000.00"), 16), ",", ".")
        Else
          strCache = strCache & "0000000000000.00"
        End If
        rsFiscalAnalitico.Close
        
        '0005    VLR. DESCONTOS          N 16  2 0028 a 0043
        strSQL = ""
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM FISReg60Analitico "
        strSQL = strSQL & "WHERE Filial =" & Filial & " "
        strSQL = strSQL & "  AND Data =#" & Format(rsFiscal("Data"), "mm/dd/yyyy") & "# "
        strSQL = strSQL & "  AND ST_Aliquota ='DESC'"
        Set rsFiscalAnalitico = db.OpenRecordset(strSQL, dbOpenSnapshot)
        If Not rsFiscalAnalitico.EOF Then
          strCache = strCache & Replace(Right(Format("" & rsFiscalAnalitico("VlrAcumulado"), "0000000000000.00"), 16), ",", ".")
        Else
          strCache = strCache & "0000000000000.00"
        End If
        rsFiscalAnalitico.Close
        
        '0006    VLR ISS                 N 16  2 0044 a 0059
        strSQL = ""
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM FISReg60Analitico "
        strSQL = strSQL & "WHERE Filial =" & Filial & " "
        strSQL = strSQL & "  AND Data =#" & Format(rsFiscal("Data"), "mm/dd/yyyy") & "# "
        strSQL = strSQL & "  AND ST_Aliquota ='ISS'"
        Set rsFiscalAnalitico = db.OpenRecordset(strSQL, dbOpenSnapshot)
        If Not rsFiscalAnalitico.EOF Then
          strCache = strCache & Replace(Right(Format("" & rsFiscalAnalitico("VlrAcumulado"), "0000000000000.00"), 16), ",", ".")
        Else
          strCache = strCache & "0000000000000.00"
        End If
        rsFiscalAnalitico.Close
        
        '0007    LEITURA Z               N 06  0 0060 a 0065
        strCache = strCache & Right(Format("0" & rsFiscal("NrContReducaoZ"), "000000"), 6)
        
        '0008    CUPOM INICIAL           C 06  0 0066 a 0071
        strCache = strCache & Left("0" & rsFiscal("NrCOOInicioDia") & String(6, " "), 6)

        '0009    CUPOM FINAL             C 06  0 0072 a 0077
        strCache = strCache & Left("0" & rsFiscal("NrCOOFimDia") & String(6, " "), 6)
        
        '0010    CRO                     N 03  0 0078 A 0080
        strCache = strCache & Right(Format("0" & rsFiscal("NrCRO"), "000"), 3)
        
        Print #lngArquivoSaida, strCache
        strCache = ""

      End If
      
      rsSaidas.Close
      
      rsFiscal.MoveNext
    Loop
    
    rsFiscal.Close
    rsParametros.Close
    Set rsFiscal = Nothing
    Set rsFiscalAnalitico = Nothing
    Set rsParametros = Nothing
    Set rsSaidas = Nothing

  End If
  
  '**************************************************************************************
  'Entradas
  '**************************************************************************************
  If Tipo = Entradas Or Tipo = Todos Then
  
    strSQL = ""
    strSQL = strSQL & "SELECT [Entradas].*, [Opera��es Entrada].[C�digo Fiscal], [Cli_For].CGC, [Cli_For].Inscri��o, [Cli_For].Estado, [Cli_For].Cidade "
    strSQL = strSQL & "FROM (Entradas INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo) "
    strSQL = strSQL & "               INNER JOIN Cli_For ON Entradas.Fornecedor = Cli_For.C�digo "
    strSQL = strSQL & "WHERE [Nota Cancelada]=0 "
    
    'Verifica o filtro da Data de Emiss�o ou Data de Entrada
    If Data = DataEmissao Then
      strSQL = strSQL & "  AND [Data Emiss�o]>=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
      strSQL = strSQL & "  AND [Data Emiss�o]<=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
    Else
      strSQL = strSQL & "  AND Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
      strSQL = strSQL & "  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
    End If
    
    'Verifica se o filtro � por filial
    If Filial > 0 Then
      strSQL = strSQL & "  AND Filial =" & Filial & " "
    End If
    
    strSQL = strSQL & "ORDER BY Sequ�ncia "
    
    Set rsEntradas = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    Do Until rsEntradas.EOF
  
      'ORDEM CAMPO                    TP TM DC POSI��ES
      '---------------------------------------------------
      '0001  IDENTIFICADOR             C 03  0 0001 A 0003
      strCache = "LCT"
      
      '0002  NUMERO DO LAN�AMENTO      N 05  0 0004 A 0008
      strCache = strCache & Right(String(5, "0") & "" & rsEntradas("Sequ�ncia"), 5)
      
      '0003  NUMERO DA NOTA FISCAL     C 10  0 0009 A 0018
      strCache = strCache & Left("" & rsEntradas("Nota Fiscal") & String(10, " "), 10)
      
      '0004  NUMERO DA NOTA FISCAL     C 10  0 0019 A 0028
      strCache = strCache & String(10, "0")
  
      '0005  ESP�CIE DA NOTA FISCAL    C 05  0 0029 A 0033
      strCache = strCache & Left("NF" & String(5, " "), 5)
      
      '0006  S�RIE DA NOTA FISCAL      C 03  0 0034 A 0036
      strCache = strCache & Left("" & rsEntradas("SerieNF") & String(3, " "), 3)
      
      '0007  DATA DE EMISS�O           D 08  0 0037 A 0044
      strCache = strCache & Right(String(8, " ") & Format("" & rsEntradas("Data Emiss�o"), "YYYYMMDD"), 8)
      
      '0008  DATA DE ENTRADA           D 08  0 0045 A 0052
      strCache = strCache & Right(String(8, " ") & Format("" & rsEntradas("Data"), "YYYYMMDD"), 8)
      
      '0009  NATUREZA DE OPERA��O      C 06  0 0053 A 0058
      'Fun��o para retornar apenas os n�meros
      strAux = ""
      For intContador = 1 To Len("" & rsEntradas("C�digo Fiscal"))
        If IsNumeric(Mid("" & rsEntradas("C�digo Fiscal"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsEntradas("C�digo Fiscal"), intContador, 1)
        End If
      Next
      strCache = strCache & Right(String(4, " ") & Left(strAux, 4), 4) & "00"
      
      '0010  SUBSTITUI��O TRIBUTARIA   L 01  0 0059 A 0059
      If rsEntradas("Base ICM Subs") > 0 Then
        strCache = strCache & "T"
      Else
        strCache = strCache & "F"
      End If
      
      '0011  CODIGO DO FOR/CLI         C 12  0 0060 A 0071
      strCache = strCache & String(12, " ")
      
      '0012  VENDA A CONSUMIDOR        L 01  0 0072 A 0072
      strAux = ""
      For intContador = 1 To Len("" & rsEntradas("C�digo Fiscal"))
        If IsNumeric(Mid("" & rsEntradas("C�digo Fiscal"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsEntradas("C�digo Fiscal"), intContador, 1)
        End If
      Next
      If (Len("" & rsEntradas("Inscri��o")) > 0 And UCase(Trim("" & rsEntradas("Inscri��o"))) <> "ISENTO") And Left(strAux, 1) = "6" Then
        strCache = strCache & "T"
      Else
        strCache = strCache & "F"
      End If
      
      '0013  CNPJ DO FOR/CLI           C 18  0 0073 A 0090
      strAux = ""
      For intContador = 1 To Len("" & rsEntradas("CGC"))
        If IsNumeric(Mid("" & rsEntradas("CGC"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsEntradas("CGC"), intContador, 1)
        End If
      Next
      
      strAux = String(14, "0") & strAux
      
      strCache = strCache & Left(Right(strAux, 14), 2) & "." & _
                            Left(Right(strAux, 12), 3) & "." & _
                            Left(Right(strAux, 9), 3) & "/" & _
                            Left(Right(strAux, 6), 4) & "-" & _
                            Right(strAux, 2)
      
      '0014  CONTA A  D�BITO           N 05  0 0091 A 0095
      strCache = strCache & String(5, " ")
      
      '0015  CONTA A CR�DITO           N 05  0 0096 A 0100
      strCache = strCache & String(5, " ")
      
      '0016  C.CUSTO A D�BITO          N 06  0 0101 A 0106
      strCache = strCache & String(6, " ")
      
      '0017  C.CUSTO A CR�DITO         N 06  0 0107 A 0112
      strCache = strCache & String(6, " ")
      
      '0018  SETOR  A D�BITO           N 06  0 0113 A 0118
      strCache = strCache & String(6, " ")
      
      '0019  SETOR A CR�DITO           N 06  0 0119 A 0124
      strCache = strCache & String(6, " ")
      
      '0020  VALOR  TOTAL DA NF        N 16  2 0125 A 0140
      strAux = Replace(Right(Format("" & rsEntradas("Total"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0021  VALOR CONT�BIL            N 16  2 0141 A 0156
      strCache = strCache & "0000000000000.00"
      
      '0022  UF                        C 03  0 0157 A 0159
      strCache = strCache & Right(String(3, " ") & UCase("" & rsEntradas("Estado")), 3)
      
      '0023  C�DIGO DO SERVI�O         N 09  0 0160 A 0168
      strCache = strCache & String(9, " ")
      
      '0024  VALOR DOS MATERIAIS       N 16  2 0169 A 0184
      strCache = strCache & "0000000000000.00"
      
      '0025  VALOR SUB-EMPREITADA      N 16  2 0185 A 0200
      strCache = strCache & "0000000000000.00"
      
      '0026  ROTINA DE CALCULO  1      N 05  0 0201 A 0205
      strCache = strCache & String(5, "0")
      
      '0027  BSC.ICMS  1               N 16  2 0206 A 0221
      strAux = Replace(Right(Format("" & rsEntradas("Base ICM"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0028  ALIQUOTA DE ICMS  1       N 05  2 0222 A 0226
      If rsEntradas("Valor ICM") > 0 And rsEntradas("Base ICM") > 0 Then
        strCache = strCache & Right(Format(Int(((rsEntradas("Valor ICM") / rsEntradas("Base ICM")) * 100)), "00.00"), 5)
      Else
        strCache = strCache & "00.00"
      End If
      
      '0029  VALOR DO ICMS  1          N 16  2 0227 A 0242
      strAux = Replace(Right(Format(rsEntradas("Valor ICM"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0030  VALOR ISENTA DE ICM  1    N 16  2 0243 A 0258
      strCache = strCache & "0000000000000.00"
      
      '0031  VALOR OUTRAS DE ICM 1     N 16  2 0259 A 0274
      strCache = strCache & "0000000000000.00"
      
      '0032  ROTINA DE CALCULO  2      N 05  0 0275 A 0279
      strCache = strCache & String(5, "0")
      
      '0033  BSC.ICMS  2               N 16  2 0280 A 0295
      strCache = strCache & "0000000000000.00"
      
      '0034  ALIQUOTA DE ICMS  2       N 05  2 0296 A 0300
      strCache = strCache & "00.00"
      
      '0035  VALOR DO ICMS  2          N 16  2 0301 A 0316
      strCache = strCache & "0000000000000.00"
      
      '0036  VALOR ISENTA DE ICM  2    N 16  2 0317 A 0332
      strCache = strCache & "0000000000000.00"
      
      '0037  VALOR OUTRAS DE ICM 2     N 16  2 0333 A 0348
      strCache = strCache & "0000000000000.00"
      
      '0038  ROTINA DE CALCULO  3      N 05  0 0349 A 0353
      strCache = strCache & String(5, "0")
      
      '0039  BSC.ICMS  3               N 16  2 0354 A 0369
      strCache = strCache & "0000000000000.00"
      
      '0040  ALIQUOTA DE ICMS  3       N 05  2 0370 A 0374
      strCache = strCache & "00.00"
      
      '0041  VALOR DO ICMS  3          N 16  2 0375 A 0390
      strCache = strCache & "0000000000000.00"
      
      '0042  VALOR ISENTA DE ICM  3    N 16  2 0391 A 0406
      strCache = strCache & "0000000000000.00"
      
      '0043  VALOR OUTRAS DE ICM 3     N 16  2 0407 A 0422
      strCache = strCache & "0000000000000.00"
      
      '0044  ROTINA DE CALCULO  4      N 05  0 0423 A 0427
      strCache = strCache & String(5, "0")
      
      '0045  BSC.ICMS  4               N 16  2 0428 A 0443
      strCache = strCache & "0000000000000.00"
      
      '0046  ALIQUOTA DE ICMS  4       N 05  2 0444 A 0448
      strCache = strCache & "00.00"
      
      '0047  VALOR DO ICMS  4          N 16  2 0449 A 0464
      strCache = strCache & "0000000000000.00"
      
      '0048  VALOR ISENTA DE ICM  4    N 16  2 0465 A 0480
      strCache = strCache & "0000000000000.00"
      
      '0049  VALOR OUTRAS DE ICM 4     N 16  2 0481 A 0496
      strCache = strCache & "0000000000000.00"
      
      '0050  ROTINA DE CALCULO  5      N 05  0 0497 A 0501
      strCache = strCache & String(5, "0")
      
      '0051  BSC.ICMS  5               N 16  2 0502 A 0517
      strCache = strCache & "0000000000000.00"
      
      '0052  ALIQUOTA DE ICMS  5       N 05  2 0518 A 0522
      strCache = strCache & "00.00"
      
      '0053  VALOR DO ICMS  5          N 16  2 0523 A 0538
      strCache = strCache & "0000000000000.00"
      
      '0054  VALOR ISENTA DE ICM  5    N 16  2 0539 A 0554
      strCache = strCache & "0000000000000.00"
      
      '0055  VALOR OUTRAS DE ICM 5     N 16  2 0555 A 0570
      strCache = strCache & "0000000000000.00"
      
      '0056  VLR TOTAL DA BSC.IPI      N 16  2 0571 A 0586
      strCache = strCache & "0000000000000.00"
      
      '0057  PERCENTUAL DE IPI         N 05  2 0587 A 0591
      strCache = strCache & "00.00"
      
      '0058  VALOR DO IPI              N 16  2 0592 A 0607
      strAux = Replace(Right(Format("" & rsEntradas("IPI"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0059  VALOR DE ISENTAS DE IPI   N 16  2 0608 A 0623
      strCache = strCache & "0000000000000.00"
      
      '0060  VALOR DE OUTRAS DE IPI    N 16  2 0624 A 0639
      strCache = strCache & "0000000000000.00"
      
      '0061  PERC.  I.R.  S/ SERVI�OS  N 05  2 0640 A 0644
      strCache = strCache & "00.00"
      
      '0062  ICMS RETIDO NA FONTE      N 16  2 0645 A 0660
      strCache = strCache & "0000000000000.00"
      
      '0063  OBS. DE IPI               N 16  2 0661 A 0676
      strCache = strCache & "0000000000000.00"
      
      '0064  BASE DE CALCULO INSS      N 16  2 0677 A 0692
      strCache = strCache & "0000000000000.00"
      
      '0065  PERCENTUAL DO INSS        N 05  2 0693 A 0697
      strCache = strCache & "00.00"
      
      '0066  VALOR DO INSS             N 16  2 0698 A 0713
      strCache = strCache & "0000000000000.00"
      
      '0067  BASE DE CALC. S. TRIB.    N 16  2 0714 A 0729
      strAux = Replace(Right(Format("" & rsEntradas("Base ICM Subs"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0068  VALOR DA S.TRIBUTARIA     N 16  2 0730 A 0745
      strAux = Replace(Right(Format("" & rsEntradas("Valor ICM Subs"), "0000000000000.00"), 16), ",", ".")
      strCache = strCache & strAux
      
      '0069  CODIGO DA ZFM             C 05  0 0746 A 0750
      Select Case UCase(Mid("" & rsEntradas("Estado"), 1, 2))
      
        Case "AM"
        
          Select Case UCase("" & rsEntradas("Cidade"))
            Case "MANAUS": strCache = strCache & "00255"
            Case "PRESIDENTE FIGUEIREDO": strCache = strCache & "09841"
            Case "RIO PRETO DA EVA": strCache = strCache & "09843"
            Case "TABATINGA": strCache = strCache & "09847"
            Case Else: strCache = strCache & "0000"
          End Select
        
        Case "AC"
        
          Select Case UCase("" & rsEntradas("Cidade"))
            Case "BRASIL�IA": strCache = strCache & "00105"
            Case "CRUZEIRO DO SUL": strCache = strCache & "00107"
            Case "EPIT�CIOLANDIA": strCache = strCache & "99998"
            Case Else: strCache = strCache & "00000"
          End Select
        
        Case "AP"
        
          Select Case UCase("" & rsEntradas("Cidade"))
            Case "MACAPA": strCache = strCache & "00605"
            Case "SANTANA": strCache = strCache & "00615"
            Case Else: strCache = strCache & "00000"
          End Select
          
        Case "RO"
        
          Select Case UCase("" & rsEntradas("Cidade"))
            Case "GUARAJA MIRIM": strCache = strCache & "00001"
            Case Else: strCache = strCache & "00000"
          End Select
        
        Case "RR"
          Select Case UCase("" & rsEntradas("Cidade"))
            Case "BONFIM": strCache = strCache & "00307"
            Case "PACARAIMA": strCache = strCache & "99999"
            Case Else: strCache = strCache & "00000"
          End Select
          
        Case Else: strCache = strCache & "00000"
      
      End Select
      
      '0070  OBSERVA��ES NECESSARIAS   C 40  0 0751 A 0790
      strCache = strCache & Left("" & rsEntradas("Observa��es") & String(40, " "), 40)
      
      '0071  FLAG DE ATUALIZA��O       L 01  0 0791 A 0791
      strCache = strCache & "F"
      
      '0072  NUMERO DA ESTA��O         C 03  0 0792 A 0794
      strCache = strCache & "001"
      
      '0073  OBSERVA��O 2              C 40  0 0795 A 0834
      strCache = strCache & String(40, " ")
      
      '0074  OBSERVA��O 3              C 40  0 0835 A 0874
      strCache = strCache & String(40, " ")
      
      '0075  CIF_FOB                   C 01  0 0875 A 0875
      If rsEntradas("obs_FretePago") <> 2 Then
        strCache = strCache & "1"
      Else
        strCache = strCache & "2"
      End If
      
      '0076  SITNOTA                   C 01  0 0876 A 0876
      strCache = strCache & " "
      
      '0077  BSCISSRET                 N 16  2 0877 A 0892
      strCache = strCache & "0000000000000.00"
      
      '0078  VLRISSRET                 N 16  2 0893 A 0908
      strCache = strCache & "0000000000000.00"
      
      '0079  ALQISSRET                 N  5  2 0909 A 0913
      strCache = strCache & "00.00"
    
      Print #lngArquivoSaida, strCache
      strCache = ""
      rsEntradas.MoveNext
      
    Loop
    
    rsEntradas.Close
    Set rsEntradas = Nothing
    
  End If
  
  Close #lngArquivoSaida
  
  g_blnExportarDadosBrasilInformatica = True
  
  Exit Function
  
TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Close #lngArquivoSaida
  g_blnExportarDadosBrasilInformatica = False
  Exit Function

End Function

'19/07/2007 - Anderson
'Exporta��o de Dados para sistema da Sadig Web
'Solicitante: Gurgel e Leite
Public Function g_blnExportarDadosSadigWeb(ByVal Filial As Byte, ByVal DataInicial As Date, ByVal DataFinal As Date, ByVal Data As expSadigWebData, ByVal Tipo As expSadigWebTipo, ByRef ARQUIVO As String) As Boolean
On Error GoTo TratarErro:

  Dim lngArquivoSaida As Long   'Informa o n�mero do arquivo dispon�vel
  Dim strCache As String        'Recebe o valor para imprimir a linha do arquivo texto
  Dim strAux As String          'Auxilia na forma��o da string
  Dim strSQL As String          'Monta a string de consulta SQL para gera��o dos dados
  Dim intContador As Integer    'Auxilia em estruturas de repeti��o
  Dim rsSaidas As Recordset     'Abre a tabela de Saidas
  Dim rsParametros As Recordset 'Abre a tabela de Parametros para obter informa��es sobre a filial.
  Dim strCaminho As String      'Informa o caminho onde deve ser salvo o arquivo
  Dim strRet As String          'Obtem retorno do arquivo ini
  
  lngArquivoSaida = FreeFile
  
  strCaminho = gsDefaultPath
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    'Path da aplica��o
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "SISTEMA", "ExportarDadosSadigWeb")
    If strRet <> "" Then strCaminho = strRet
  End If
  
  If ARQUIVO = "" Then
    ARQUIVO = strCaminho & Format(Now(), "yyyyMMddhhnnss") & ".txt"
  End If
  
  Open ARQUIVO For Output As #lngArquivoSaida
  
  Set rsParametros = db.OpenRecordset("SELECT * FROM [Par�metros Filial] WHERE Filial=" & Filial)
  
  '**************************************************************************************
  'Sa�das
  '**************************************************************************************
  If Tipo = SaidasSadigWeb Then
  
    strSQL = ""
    strSQL = strSQL & "SELECT Sa�das.Filial, Sa�das.Data, Sa�das.Sequ�ncia, [Sa�das - Produtos].Linha, Cli_For.CGC, Cli_For.Nome, Cli_For.Endere�o, Cli_For.Cidade, Cli_For.CEP, Cli_For.Estado, Cli_For.[Fone 1], Cli_For.SadigWeb_Tipo, [Opera��es Sa�da].SadigWeb_Tipo, [Sa�das - Produtos].[C�digo sem Grade], Produtos.Nome, [Sa�das - Produtos].Qtde, Funcion�rios.Nome, Funcion�rios.SadigWeb_CDRC "
    strSQL = strSQL & "FROM (((([Opera��es Sa�da] INNER JOIN Sa�das ON [Opera��es Sa�da].C�digo = Sa�das.Opera��o) INNER JOIN Cli_For ON Sa�das.Cliente = Cli_For.C�digo) INNER JOIN Funcion�rios ON Sa�das.Digitador = Funcion�rios.C�digo) INNER JOIN [Sa�das - Produtos] ON (Sa�das.Filial = [Sa�das - Produtos].Filial) AND (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia)) INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo "
    strSQL = strSQL & "WHERE [Nota Impressa]>0 "
    strSQL = strSQL & "  AND [Nota Cancelada]=0 "
    strSQL = strSQL & "  AND [Movimenta��o Desfeita]=0 "

    'Verifica o filtro da Data de Emiss�o ou Data de Entrada
    If Data = DataEmissaoSadigWeb Then
      strSQL = strSQL & "  AND [DataEmissaoNota]>=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
      strSQL = strSQL & "  AND [DataEmissaoNota]<=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
    Else
      strSQL = strSQL & "  AND Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
      strSQL = strSQL & "  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
    End If
    
    'Verifica se o filtro � por filial
    If Filial > 0 Then
      strSQL = strSQL & "  AND Sa�das.Filial =" & Filial & " "
    End If
    
    strSQL = strSQL & "ORDER BY Sa�das.Sequ�ncia, [Sa�das - Produtos].Linha  "
    
    Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    Do Until rsSaidas.EOF
  
      'CNPJ Distribuidor - Campo destinado ao c�digo do CGC do Distribuidor sem m�scara. Trazer os 0 (zero) a  esquerda (Quando houver).Ex.: 02557889000128
      strAux = ""
      For intContador = 1 To Len("" & rsParametros("CGC"))
        If IsNumeric(Mid("" & rsParametros("CGC"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsParametros("CGC"), intContador, 1)
        End If
      Next
      strAux = String(14, "0") & strAux
      strCache = Right(strAux, 14)

      'Nome da Empresa - Campo destinado a raz�o social da distribuidora
      strCache = strCache & Left(rsParametros("nome") & String(50, " "), 50)
      
      'CNPJ Cliente - Campo destinado ao CNPJ ou CPF do cliente da distribuidora sem mascara. Ex. 62527619000156 ou 06757010829, Trazer os 0 (zero) a esquerda e no caso do CPF alinhado a esquerda com espa�os em branco a direita
      strAux = ""
      For intContador = 1 To Len("" & rsSaidas("CGC"))
        If IsNumeric(Mid("" & rsSaidas("CGC"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsSaidas("CGC"), intContador, 1)
        End If
      Next
      'Se for CPF, preencher com espa�os em branco.
      If Len(strAux) < 14 Then
        strAux = String(14, " ") & strAux
      Else
        strAux = String(14, "0") & strAux
      End If
      strCache = strCache & Right(strAux, 14)
      
      'Nome Cliente - Campo destinado � raz�o social do cliente
      strCache = strCache & Left(rsSaidas("Cli_For.Nome") & String(60, " "), 60)
  
      'Endere�o - Endere�o do cliente
      strCache = strCache & Left(rsSaidas("Endere�o") & String(50, " "), 50)
  
      'Cidade - Cidade do cliente
      strCache = strCache & Left(rsSaidas("Cidade") & String(50, " "), 50)
      
      'CEP - Campo destinado ao n�mero do CEP sem m�scara. Ex. 13690000, com zeros a esquerda se houver
      strAux = ""
      For intContador = 1 To Len("" & rsSaidas("CEP"))
        If IsNumeric(Mid("" & rsSaidas("CEP"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsSaidas("CEP"), intContador, 1)
        End If
      Next
      strAux = String(8, "0") & strAux
      strCache = strCache & Right(strAux, 8)
      
      'Estado - Campo destinado ao estado do cliente
      strCache = strCache & Right(String(2, " ") & rsSaidas("Estado"), 2)
      
      'Telefone - Campo com numero do telefone com DDD. Ex. 1935839000
      strAux = ""
      For intContador = 1 To Len("" & rsSaidas("Fone 1"))
        If IsNumeric(Mid("" & rsSaidas("Fone 1"), intContador, 1)) Then
          strAux = strAux & Mid("" & rsSaidas("Fone 1"), intContador, 1)
        End If
      Next
      strAux = strAux & String(20, " ")
      strCache = strCache & Left(strAux, 20)
      
      'Tipo Cliente - Campo destinado � segmenta��o do cliente
      strCache = strCache & Left("" & rsSaidas("Cli_For.SadigWeb_Tipo") & String(40, " "), 40)
      
      'Tipo Saida - Campo destinado ao tipo de sa�da do produto
      strCache = strCache & Left("" & rsSaidas("Opera��es Sa�da.SadigWeb_Tipo") & String(15, " "), 15)
      
      'CODIGO Produto Royal - Campo destinado ao c�digo de produto da Royal Canin, ou seja, se na CDRC o c�digo Special Croc 15KG � numero 100, ter� que ser enviado o c�digo da RCB  - 013115
      strCache = strCache & Left("" & rsSaidas("C�digo sem Grade") & String(14, " "), 14)
      
      'Descri��o Produto Royal - Descri��o do Produto vendido
      strCache = strCache & Left(rsSaidas("Produtos.Nome") & String(50, " "), 50)
      
      'Qtde. Saida - Quantidade de CX/SC/LT/PC sa�das. Obs. Sem separador de milhares Ex. 999999
      strCache = strCache & Left(String(6, " ") & Int(rsSaidas("Qtde") * 100), 6)
      
      'Data Movimento - Data da sa�da do produto, formato dd/mm/aaaa
      strCache = strCache & Right(String(10, " ") & Format("" & rsSaidas("Data"), "DD/MM/YYYY"), 10)
      
      'C�digo Vendedor - C�digo do vendedor na CDRC
      strCache = strCache & Left("" & rsSaidas("SadigWeb_CDRC") & String(20, " "), 20)

      'Nome Vendedor - Nome do Vendedor
      strCache = strCache & Left("" & rsSaidas("Funcion�rios.Nome") & String(30, " "), 30)
      
      Print #lngArquivoSaida, strCache
      strCache = ""
      rsSaidas.MoveNext
      
    Loop
    
    rsSaidas.Close
    Set rsSaidas = Nothing

  End If
  
  Close #lngArquivoSaida
  
  g_blnExportarDadosSadigWeb = True
  
  Exit Function
  
TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Close #lngArquivoSaida
  g_blnExportarDadosSadigWeb = False
  Exit Function
End Function

'19/07/2007 - Anderson
'Gera��o de dados para relat�rio de Vendas por Fornecedor
'Solicitante: Nutricare
Public Function g_blnRelVendasFornecedor(ByVal Filial As Byte, ByVal DataInicial As Date, ByVal DataFinal As Date, ByVal Fornecedor As String, ByVal Vendedor As String, ByVal ProdutoClasse As String, ByVal ProdutoSubClasse As String, ByVal Cidade As String, ByVal Estado As String, Optional ByRef BarraProgresso As ProgressBar) As Boolean
On Error GoTo TratarErro:

  Dim strSQL As String          'Monta a string de consulta SQL para gera��o dos dados
  Dim rsSaidas As Recordset     'Abre a tabela de Saidas
  Dim rsVendasFornecedor As Recordset ' abre a tabela tempor�ria para adi��o e dados
  
  strSQL = "SELECT Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Cli_For_1.C�digo, Cli_For_1.Nome, Cli_For.Nome, Cli_For.CGC, Cli_For.Cidade, Cli_For.Estado, Produtos.Nome, Classes.C�digo, Classes.Nome, [Sub Classes].C�digo, [Sub Classes].Nome, Sum([Sa�das - Produtos].Qtde) AS TotalQuantidade, Sum([Sa�das - Produtos].[Pre�o Final]) AS TotalPrecoFinal "
  strSQL = strSQL & " FROM ([Sub Classes] INNER JOIN (Classes INNER JOIN (((((Sa�das INNER JOIN [Sa�das - Produtos] ON (Sa�das.Filial = [Sa�das - Produtos].Filial) AND (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia)) INNER JOIN Cli_For ON Sa�das.Cliente = Cli_For.C�digo) INNER JOIN Forn_Prod ON [Sa�das - Produtos].[C�digo sem Grade] = Forn_Prod.Produto) INNER JOIN Cli_For AS Cli_For_1 ON Forn_Prod.Fornecedor = Cli_For_1.C�digo) INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo) ON Classes.C�digo = Produtos.Classe) ON [Sub Classes].C�digo = Produtos.[Sub Classe]) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
  strSQL = strSQL & "WHERE Sa�das.Efetivada = -1 And Sa�das.[Nota Cancelada] = 0 And [Opera��es Sa�da].Tipo = 'V' AND (Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "#) "
  
  If Vendedor <> "0" Then
    strSQL = strSQL & "AND Sa�das.Digitador = " & Vendedor & " "
  End If
  
  strSQL = strSQL & "GROUP BY Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Cli_For_1.C�digo, Cli_For_1.Nome, Cli_For.Nome, Cli_For.CGC, Cli_For.Cidade, Cli_For.Estado, Produtos.Nome, Classes.C�digo, Classes.Nome, [Sub Classes].C�digo, [Sub Classes].Nome "
  strSQL = strSQL & "Having Sa�das.Filial = " & Filial & " "
  
  If Fornecedor <> "0" Then
    strSQL = strSQL & "AND Cli_For_1.C�digo =  " & Fornecedor & " "
  End If

  If Cidade <> "" Then
    strSQL = strSQL & "AND Cli_For.Cidade = '" & Cidade & "' "
  End If
  
  If Estado <> "" Then
    strSQL = strSQL & "AND Cli_For.Estado = '" & Estado & "' "
  End If
  
  If ProdutoClasse <> "0" Then
    strSQL = strSQL & "AND Classes.C�digo = " & ProdutoClasse & " "
  End If
  
  If ProdutoSubClasse <> "0" Then
    strSQL = strSQL & "AND [Sub Classes].C�digo = " & ProdutoSubClasse & " "
  End If

  strSQL = strSQL & "ORDER BY Cli_For_1.Nome, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade]"
  
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  dbTemp.Execute "DELETE * FROM RelVendasFornecedor", dbFailOnError
  
  Set rsVendasFornecedor = dbTemp.OpenRecordset("RelVendasFornecedor")
  
  With rsSaidas
  
    If (.BOF And .EOF) Then
      If Not BarraProgresso Is Nothing Then
        BarraProgresso.min = 0
        BarraProgresso.Max = 1
        BarraProgresso.Value = 0
      End If
      MsgBox "N�o h� informa��es para serem exibidas no relat�rio. Verifique se os filtros foram preenchidos corretamente.", vbInformation, "Quick Store"
      g_blnRelVendasFornecedor = False
      
      rsSaidas.Close
      rsVendasFornecedor.Close
      
      Set rsSaidas = Nothing
      Set rsVendasFornecedor = Nothing
      
      Exit Function
    End If
    
    .MoveLast
    .MoveFirst
    
    If Not BarraProgresso Is Nothing Then
      BarraProgresso.min = 0
      BarraProgresso.Max = .RecordCount + 1
    End If
    
    Do Until .EOF
      DoEvents
      rsVendasFornecedor.AddNew
      rsVendasFornecedor("FornecedorCodigo").Value = .Fields("Cli_For_1.C�digo")
      rsVendasFornecedor("FornecedorNome").Value = "" & .Fields("Cli_For_1.Nome")
      rsVendasFornecedor("ClienteCodigo").Value = .Fields("Cliente")
      rsVendasFornecedor("ClienteNome").Value = "" & .Fields("Cli_For.Nome")
      rsVendasFornecedor("ClienteCNPJCPF").Value = " " & .Fields("CGC")
      rsVendasFornecedor("ClienteCidade").Value = " " & .Fields("Cidade")
      rsVendasFornecedor("ClienteEstado").Value = " " & .Fields("Estado")
      rsVendasFornecedor("ProdutoCodigo").Value = " " & .Fields("C�digo sem Grade")
      rsVendasFornecedor("ProdutoNome").Value = "" & .Fields("Produtos.Nome")
      rsVendasFornecedor("ProdutoQuantidade").Value = .Fields("TotalQuantidade")
      rsVendasFornecedor("ProdutoValor").Value = .Fields("TotalPrecoFinal")
      rsVendasFornecedor("ProdutoClasseCodigo").Value = .Fields("Classes.C�digo")
      rsVendasFornecedor("ProdutoClasseNome").Value = "" & .Fields("Classes.Nome")
      rsVendasFornecedor("ProdutoSubClasseCodigo").Value = .Fields("Sub Classes.C�digo")
      rsVendasFornecedor("ProdutoSubClasseNome").Value = "" & .Fields("Sub Classes.Nome")
      rsVendasFornecedor.Update
      
      If Not BarraProgresso Is Nothing Then
        BarraProgresso.Value = .AbsolutePosition
      End If
      .MoveNext
    Loop
    rsVendasFornecedor.Close
    .Close
  End With
  
  If Not BarraProgresso Is Nothing Then
    BarraProgresso.min = 0
    BarraProgresso.Max = 1
    BarraProgresso.Value = 0
  End If
  
  Set rsVendasFornecedor = Nothing
  Set rsSaidas = Nothing
  
  g_blnRelVendasFornecedor = True
  
  Exit Function
  
TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Set rsVendasFornecedor = Nothing
  Set rsSaidas = Nothing
  g_blnRelVendasFornecedor = False

End Function

'10/09/2007 - Anderson
'Fun��o criada para gerar log do contas a receber
'Solicitado: Agrotama (Technomax)
Public Sub SystemLog(ByVal strData As String, ByVal strHora As String, ByVal strUsuario As String, ByVal intOperacao As enuSystemLog, ByVal strDescricao As String, ByVal strModulo As String, ByVal strTabela As String, Optional ByVal strArquivo As String)
On Error GoTo TratarErro

    Dim intFileNumber As Integer
    Dim strOperacao As String
    
    Select Case intOperacao
      Case 1: strOperacao = "Inserir"
      Case 2: strOperacao = "Alterar"
      Case 3: strOperacao = "Excluir"
      Case 9: strOperacao = "Outros"
    End Select

    If Len(strArquivo) = 0 Then
      strArquivo = gsDefaultPath & "system.log"
    End If

    intFileNumber = FreeFile
    Open strArquivo For Append As intFileNumber
    Print #intFileNumber, strData & ";" & strHora & ";" & strUsuario & ";" & strOperacao & ";" & strDescricao & ";" & strModulo & ";" & strTabela
    Close intFileNumber

TratarErro:
    Err.Clear
End Sub

'19/10/2007 - Anderson
'Fun��o criada para verificar o desconto permitido por classe de produto
'Solitante: Agrotama
Public Function PermiteDescontoMargemLucro(ByVal CodigoProduto As String, ByVal ValorDesconto As Double, ByVal Quantidade As Double, ByVal PrecoUnitario As Double) As Boolean
  Dim strSQL As String
  Dim rsCusto As Recordset
  Dim dblPrecoCusto As Double
  Dim dblPrecoMinimoPermitido As Double
  
  PermiteDescontoMargemLucro = False
  
  strSQL = strSQL & "SELECT Pre�os.Tabela, Pre�os.Produto, Pre�os.Pre�o, Produtos.Classe, Classes.Nome, Classes.LucroMinimoPermitido "
  strSQL = strSQL & "FROM (Classes INNER JOIN Produtos ON Classes.C�digo = Produtos.Classe) INNER JOIN Pre�os ON Produtos.C�digo = Pre�os.Produto "
  strSQL = strSQL & "WHERE Pre�os.Produto='" & CodigoProduto & "' AND Pre�os.Tabela='CUSTO' "

  Set rsCusto = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If Not rsCusto.EOF Then
    dblPrecoCusto = rsCusto("Pre�o").Value
    dblPrecoMinimoPermitido = (rsCusto("Pre�o").Value + (rsCusto("Pre�o").Value * rsCusto("LucroMinimoPermitido").Value / 100)) * Quantidade
  End If
  
  rsCusto.Close
  Set rsCusto = Nothing
  
  If (Quantidade * PrecoUnitario) - ((Quantidade * PrecoUnitario) * ValorDesconto / 100) >= dblPrecoMinimoPermitido Then
    PermiteDescontoMargemLucro = True
  End If
End Function

'30/10/2007 - Anderson
'Fun��o criada para gerar relat�rio de Produtos a comprar
'Solicitante: King Cross
Public Function g_bolRelatorioProdutosComprar(ByVal Filial As Integer, ByVal DataInicial As Date, ByVal DataFinal As Date, ByVal CodigoProduto As String, ByVal Fornecedor As String, ByVal Classe As Integer, ByVal SubClasse As Integer, ByVal AtivarEspacoFisico As Boolean, ByVal Fator As Long) As Boolean
'Fator - Indica a quantidade de dias prevista para compra de produtos
  Dim strSQL As String
  Dim rsRelatorio As Recordset
  Dim rsProdutosComprar As Recordset
  Dim rsProduto
  Dim lngDiferencaDias As Long
  Dim Erro As Integer
  Dim dblEstoqueAtual As Double
  
  g_bolRelatorioProdutosComprar = False
  
  lngDiferencaDias = DateDiff("d", DataInicial, DataFinal)
  
  If lngDiferencaDias <= 0 Then lngDiferencaDias = 1
  
  strSQL = ""
  strSQL = strSQL & "SELECT Sa�das.Filial, Produtos.C�digo, Produtos.Nome, Forn_Prod.Fornecedor, Cli_For.Fantasia, Produtos.Classe, Classes.Nome, Produtos.[Sub Classe], [Sub Classes].Nome, Produtos.EspacoFisicoTotal, Produtos.Fracionado, Sum([Sa�das - Produtos].Qtde) AS SomaDeQtde "
  strSQL = strSQL & "FROM (Sa�das INNER JOIN (Cli_For INNER JOIN ((((Produtos INNER JOIN [Sa�das - Produtos] ON Produtos.C�digo = [Sa�das - Produtos].[C�digo sem Grade]) INNER JOIN Classes ON Produtos.Classe = Classes.C�digo) INNER JOIN [Sub Classes] ON Produtos.[Sub Classe] = [Sub Classes].C�digo) INNER JOIN Forn_Prod ON Produtos.C�digo = Forn_Prod.Produto) ON Cli_For.C�digo = Forn_Prod.Fornecedor) ON (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia) AND (Sa�das.Filial = [Sa�das - Produtos].Filial)) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
  strSQL = strSQL & "WHERE Sa�das.Efetivada = -1 "
  strSQL = strSQL & "  AND Sa�das.[Nota Cancelada] = 0 "
  strSQL = strSQL & "  AND [Opera��es Sa�da].Tipo = 'V' "
  strSQL = strSQL & "  AND (Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "#  "
  strSQL = strSQL & "  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & "  AND Sa�das.Filial=" & Filial & " "
  
  If Fornecedor <> "0" Then
    strSQL = strSQL & "AND Forn_Prod.Fornecedor =  " & Fornecedor & " "
  End If
  
  If CodigoProduto <> "0" Then
    strSQL = strSQL & "AND Produtos.C�digo =  '" & CodigoProduto & "' "
  End If
  
  If Classe <> "0" Then
    strSQL = strSQL & "AND Produtos.Classe = " & Classe & " "
  End If
  
  If SubClasse <> "0" Then
    strSQL = strSQL & "AND Produtos.[Sub Classe] = " & SubClasse & " "
  End If
  
  strSQL = strSQL & "GROUP BY Sa�das.Filial, Produtos.C�digo, Produtos.Nome, Forn_Prod.Fornecedor, Cli_For.Fantasia, Produtos.Classe, Classes.Nome, Produtos.[Sub Classe], [Sub Classes].Nome, Produtos.EspacoFisicoTotal, Produtos.Fracionado "

  Set rsRelatorio = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  dbTemp.Execute "DELETE * FROM RelProdutosComprar", dbFailOnError
  
  Set rsProdutosComprar = dbTemp.OpenRecordset("RelProdutosComprar")

  Do Until rsRelatorio.EOF
  
    With rsProdutosComprar
      
      dblEstoqueAtual = Acha_Estoque(rsRelatorio("Filial"), rsRelatorio("C�digo"), 0, 0, 0, Erro)
      If Erro <> 0 Then dblEstoqueAtual = 0
      
      .AddNew
      .Fields("CodigoProduto").Value = rsRelatorio("C�digo")
      .Fields("Descricao").Value = rsRelatorio("Produtos.Nome")
      .Fields("CodigoClasse").Value = rsRelatorio("Classe")
      .Fields("Classe").Value = rsRelatorio("Classes.Nome")
      .Fields("CodigoSubClasse").Value = rsRelatorio("Sub Classe")
      .Fields("SubClasse").Value = rsRelatorio("Sub Classes.Nome")
      .Fields("CodigoFornecedor").Value = rsRelatorio("Fornecedor")
      .Fields("Fornecedor").Value = rsRelatorio("Fantasia")
      .Fields("EstoqueFisicoAtual").Value = dblEstoqueAtual
      If AtivarEspacoFisico Then
        .Fields("EstoqueFisicoTotal").Value = "0" & rsRelatorio("EspacoFisicoTotal")
        .Fields("EspacoFisicoDisponivel").Value = .Fields("EstoqueFisicoTotal").Value - .Fields("EstoqueFisicoAtual").Value
        
        If rsRelatorio("Fracionado") <> 0 Then
          .Fields("MediaVendas").Value = rsRelatorio("SomaDeQtde") / lngDiferencaDias
        Else
          .Fields("MediaVendas").Value = Int(rsRelatorio("SomaDeQtde") / lngDiferencaDias)
        End If
        
        If (Fator * .Fields("MediaVendas").Value) - .Fields("EstoqueFisicoAtual").Value > .Fields("EspacoFisicoDisponivel").Value Then
          .Fields("QuantidadeComprar").Value = .Fields("EspacoFisicoDisponivel").Value
        ElseIf (Fator * .Fields("MediaVendas").Value) - .Fields("EstoqueFisicoAtual").Value < 0 Then
          .Fields("QuantidadeComprar").Value = 0
        Else
          .Fields("QuantidadeComprar").Value = ((Fator * .Fields("MediaVendas").Value) - .Fields("EstoqueFisicoAtual").Value)
        End If
      Else
        .Fields("EstoqueFisicoTotal").Value = 0
        .Fields("EspacoFisicoDisponivel").Value = 0
        If rsRelatorio("Fracionado") <> 0 Then
          .Fields("MediaVendas").Value = rsRelatorio("SomaDeQtde") / lngDiferencaDias
        Else
          .Fields("MediaVendas").Value = Int(rsRelatorio("SomaDeQtde") / lngDiferencaDias)
        End If
        
        If (Fator * .Fields("MediaVendas").Value) - .Fields("EstoqueFisicoAtual").Value < 0 Then
          .Fields("QuantidadeComprar").Value = 0
        Else
          .Fields("QuantidadeComprar").Value = ((Fator * .Fields("MediaVendas").Value) - .Fields("EstoqueFisicoAtual").Value)
        End If

      End If
      
      .Update
      
    End With
  
    rsRelatorio.MoveNext
    
  Loop

  rsRelatorio.Close
  Set rsRelatorio = Nothing
  
  g_bolRelatorioProdutosComprar = True
  
End Function
'08/01/2008 - Anderson
'Exporta��o de Dados de vendas para cliente Pearson
'Solicitante: Rodrigo Technomax
Public Function g_blnExportarDadosPearson(ByVal Filial As Byte, ByVal DataInicial As Date, ByVal DataFinal As Date, ByVal strOperacao As String) As Boolean
  On Error GoTo TratarErro

  Dim strCPFCNPJ As String    'Obtem o CPF ou CNPJ
  Dim strSQL As String        'Monta a string de consulta SQL para gera��o dos dados
  Dim intContador As Integer  'Auxilia em estruturas de repeti��o
  Dim rsSaidas As Recordset   'Abre a tabela de Saidas
  Dim strODBC As String
  Dim strRet As String        'Obtem retorno do arquivo ini
  Dim dbDbase As Database
'  Dim rsORPBIE As Recordset
'  Dim rsTORPBIE As Recordset
  Dim rsSaidasProdutos As Recordset
  Dim wsDbase As Workspace
  Dim strICM_On_IPI As String
  Dim dblDescontoItem As Double
  
  strODBC = ""
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    'DSN ODBC
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "SISTEMA", "ODBCPearson")
    If strRet <> "" Then strODBC = strRet
  End If

  Set wsDbase = DBEngine.CreateWorkspace("ODBCDbase", "admin", "", dbUseODBC)
  Set dbDbase = wsDbase.OpenDatabase("", dbDriverComplete, False, "ODBC;DSN=" & strODBC)
  dbDbase.Execute "DELETE FROM ORPBIE"
  dbDbase.Execute "DELETE FROM TORPBIE"
'  Set rsORPBIE = dbDbase.OpenRecordset("ORPBIE")
'  Set rsTORPBIE = dbDbase.OpenRecordset("SELECT * FROM TORPBIE", dbOpenDynaset)
  
  '**************************************************************************************
  'Sa�das
  '**************************************************************************************
  
  '11/06/2008 - mpdea
  'Modificado m�todo de sele��o do COO
  '  de Mid([Sa�das]![Observa��es],1,21)='Venda Fiscal COO nr. ' para [Cupom Fiscal Impresso]
  'e de identifica��o do COO
  '  de Mid([Sa�das]![Observa��es],21) para Right([Sa�das]![Observa��es], 6)
  '19/03/2008 - mpdea
  'Inclu�do Nome e Fone 1
  strSQL = "SELECT [Sa�das].*, [Opera��es Sa�da].[C�digo Fiscal], [Cli_For].CGC, [Cli_For].C�digo, "
  strSQL = strSQL & "[Cli_For].Nome, [Cli_For].Inscri��o, [Cli_For].[Fone 1], [Cli_For].Estado, "
  strSQL = strSQL & "[Cli_For].Cidade, [Cli_For].Endere�o, [Cli_For].Bairro, [Cli_For].Cidade, "
  strSQL = strSQL & "[Cli_For].CEP, [Cli_For].Estado, Right([Sa�das]![Observa��es], 6) AS COO "
  strSQL = strSQL & "FROM ([Opera��es Sa�da] INNER JOIN Sa�das ON [Opera��es Sa�da].C�digo = Sa�das.Opera��o) "
  strSQL = strSQL & "INNER JOIN Cli_For ON Sa�das.Cliente = Cli_For.C�digo "
  strSQL = strSQL & "WHERE [Nota Cancelada] = 0 "
  strSQL = strSQL & "AND [Movimenta��o Desfeita] = 0 "
  strSQL = strSQL & "AND [Data]>=#" & Format(DataInicial, "mm/dd/yyyy") & "# "
  strSQL = strSQL & "AND [Data]<=#" & Format(DataFinal, "mm/dd/yyyy") & "# "
  strSQL = strSQL & "AND [Cupom Fiscal Impresso]"
  
  'Verifica se o filtro � por filial
  If Filial > 0 Then
    strSQL = strSQL & "  AND Filial =" & Filial & " "
  End If
  
  If strOperacao > 0 Then
    strSQL = strSQL & "  AND Opera��o =" & strOperacao & " "
  End If

  strSQL = strSQL & "ORDER BY Sequ�ncia "
  
  '11/06/2008 - mpdea
  'Inclu�do par�metro somente leitura
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  Do Until rsSaidas.EOF
  
    strCPFCNPJ = ""
    For intContador = 1 To Len("" & rsSaidas("CGC"))
      If IsNumeric(Mid("" & rsSaidas("CGC"), intContador, 1)) Then
        strCPFCNPJ = strCPFCNPJ & Mid("" & rsSaidas("CGC"), intContador, 1)
      End If
    Next

    '19/03/2008 - mpdea
    'Inclu�do Nome, Inscri��o Estadual (ou RG) e Fone 1
    'Formatado c�digo SQL para melhor exibi��o
    '
    'Atualiza o arquivo ORPBIE.DBF
    strSQL = ""
    strSQL = strSQL & "INSERT INTO ORPBIE "
    strSQL = strSQL & "(TipoNota, Tipo, Numero, Cliente, Emissao, Saida, Cadastro, Entrega, Representa, Condicao, "
    strSQL = strSQL & "TipoDoc, Conta, Quan_embal, Espe_embal, Pedido1, AC_Total, AC_merc, AC_IPI, AC_ICM, AC_Desc, "
    strSQL = strSQL & "Base_ICM, Base_IPI, Desc_Medio, Desc_Real, Aipi_medio, AC_p_liqui, Ac_p_bruto, Emitido, "
    strSQL = strSQL & "Cancelado, Grupo, Codigo, CIC_ou_CGC, Endereco, Bairro, CEP, Cidade, Estado, NOME, IE_OU_RG, FONE) "
    strSQL = strSQL & "VALUES ('PDV','NF','" & Format(rsSaidas("COO"), "0000000000") & "','" & rsSaidas("C�digo") & "',#"
    strSQL = strSQL & Format(rsSaidas("Data"), "MM/DD/YYYY") & "#,#" & Format(rsSaidas("Data"), "MM/DD/YYYY") & "#,#"
    strSQL = strSQL & Format(rsSaidas("Data"), "MM/DD/YYYY") & "#,#" & Format(rsSaidas("Data"), "MM/DD/YYYY")
    strSQL = strSQL & "#,'00','AV','FT001','005',' ',' ','" & rsSaidas("Sequ�ncia") & "',"
    strSQL = strSQL & Replace(rsSaidas("Total"), ",", ".") & "," & Replace(rsSaidas("Produtos"), ",", ".") & ","
    strSQL = strSQL & Replace(rsSaidas("IPI"), ",", ".") & "," & Replace(rsSaidas("Valor ICM"), ",", ".") & ","
    strSQL = strSQL & Replace(rsSaidas("Desconto") + rsSaidas("DescontoSubTotal"), ",", ".") & ","
    strSQL = strSQL & Replace(rsSaidas("Base ICM"), ",", ".") & ",0,0,0,0,0,0,True,False,'CLI',"
    strSQL = strSQL & rsSaidas("C�digo") & ",'" & strCPFCNPJ & "','" & rsSaidas("Endere�o") & "','"
    strSQL = strSQL & rsSaidas("Bairro") & "','" & rsSaidas("CEP") & "','" & rsSaidas("Cidade") & "','"
    strSQL = strSQL & rsSaidas("Estado") & "','" & rsSaidas("Nome") & "','" & rsSaidas("Inscri��o") & "','"
    strSQL = strSQL & rsSaidas("Fone 1") & "') "
  
    dbDbase.Execute strSQL
    
    Set rsSaidasProdutos = db.OpenRecordset("SELECT * FROM [Sa�das - Produtos] WHERE Filial=" & rsSaidas("Filial") & " AND Sequ�ncia=" & rsSaidas("Sequ�ncia") & " ORDER BY Filial, Sequ�ncia, Linha", dbOpenDynaset)
    
    Do Until rsSaidasProdutos.EOF
    
      If rsSaidasProdutos("ICM") > 0 Or rsSaidasProdutos("IPI") > 0 Then
        strICM_On_IPI = True
      Else
        strICM_On_IPI = False
      End If
      
      dblDescontoItem = ((rsSaidasProdutos("Pre�o") * rsSaidasProdutos("Qtde")) + ((rsSaidasProdutos("Pre�o") * rsSaidasProdutos("Qtde")) * (rsSaidasProdutos("IPI") / 100))) * (rsSaidasProdutos("Desconto") / 100)
    
      'Atualiza o arquivo TORPBIE.DBF
      strSQL = ""
      strSQL = strSQL & "INSERT INTO TORPBIE "
      strSQL = strSQL & "       (Tipo, Ordem, Numero, Item, Cliente, Entrega, Qt, Valor, Natureza, ICM, IPI, ICM_on_IPI, Des_Ou_ACR ) "
      strSQL = strSQL & "VALUES ('NF','" & rsSaidasProdutos("Linha") & "','" & Format(rsSaidas("COO"), "0000000000") & "','" & rsSaidasProdutos("C�digo") & "','" & rsSaidas("C�digo") & "',#" & Format(rsSaidas("Data"), "MM/DD/YYYY") & "#," & Replace(rsSaidasProdutos("Qtde"), ",", ".") & "," & Replace(rsSaidasProdutos("Pre�o"), ",", ".") & ",'" & rsSaidasProdutos("CFOP") & "'," & Replace(rsSaidasProdutos("ICM"), ",", ".") & "," & Replace(rsSaidasProdutos("IPI"), ",", ".") & "," & strICM_On_IPI & ",'" & Replace(Format(dblDescontoItem, "0.00"), ",", ".") & "') "
      
      dbDbase.Execute strSQL
      
      rsSaidasProdutos.MoveNext
      
    Loop
    
    rsSaidasProdutos.Close
    Set rsSaidasProdutos = Nothing
    
    rsSaidas.MoveNext
    
  Loop
  
  rsSaidas.Close
  Set rsSaidas = Nothing
  dbDbase.Close
  Set dbDbase = Nothing
  wsDbase.Close
  Set wsDbase = Nothing
  
  g_blnExportarDadosPearson = True
  
  Exit Function
  
TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  g_blnExportarDadosPearson = False
  Exit Function

End Function

'16/01/2008 - Anderson
'Gera��o de dados para relat�rio de Vendas
'Solicitante: LL Comercio de Ferramentas
Public Function g_blnRelVendasII(ByVal Filial As Byte, ByVal Cliente As String, ByVal Produto As String, ByVal ProdutoClasse As String, ByVal ProdutoSubClasse As String, ByVal Fornecedor As String, ByVal Vendedor As String, ByVal Operacao As String, ByVal DataInicial As Date, ByVal DataFinal As Date, Optional ByRef BarraProgresso As ProgressBar) As Boolean
On Error GoTo TratarErro:

  Dim strSQL As String          'Monta a string de consulta SQL para gera��o dos dados
  Dim rsSaidas As Recordset     'Abre a tabela de Saidas
  Dim rsEntradas As Recordset   'Abre a tabela de Entradas
  Dim rsRelVendas As Recordset ' abre a tabela tempor�ria para adi��o e dados
  Dim rsAux As Recordset
  
  strSQL = "SELECT [Sa�das - Produtos].[C�digo sem Grade], Produtos.Nome, Sum([Sa�das - Produtos].Qtde) AS Quantidade, Sum([Pre�o]*[Qtde]) AS ValorVenda, Sum([Pre�o Final]*([Sa�das - Produtos].[ICM]/100)) AS ICMSVenda "
  If Fornecedor <> "0" Then
    strSQL = strSQL & " FROM Cli_For INNER JOIN (((Sa�das INNER JOIN ([Sa�das - Produtos] INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo) ON (Sa�das.Filial = [Sa�das - Produtos].Filial) AND (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia)) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo) INNER JOIN Forn_Prod ON Produtos.C�digo = Forn_Prod.Produto) ON Cli_For.C�digo = Forn_Prod.Fornecedor "
  Else
    strSQL = strSQL & " FROM (Sa�das INNER JOIN ([Sa�das - Produtos] INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo) ON (Sa�das.Filial = [Sa�das - Produtos].Filial) AND (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia)) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
  End If
  strSQL = strSQL & "WHERE Sa�das.Filial = " & Filial & " AND Sa�das.Efetivada = -1 And Sa�das.[Nota Cancelada] = 0 And [Opera��es Sa�da].Tipo = 'V' AND (Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "#) "
  
  If Cliente <> "0" Then
    strSQL = strSQL & "AND Sa�das.Cliente=" & Cliente & " "
  End If
  
  If Produto <> "0" Then
    strSQL = strSQL & "AND [Sa�das - Produtos].[C�digo sem Grade]='" & Produto & "' "
  End If
  
  If ProdutoClasse <> "0" Then
    strSQL = strSQL & "AND Produtos.Classe = " & ProdutoClasse & " "
  End If
  
  If ProdutoSubClasse <> "0" Then
    strSQL = strSQL & "AND Produtos.[Sub Classe] = " & ProdutoSubClasse & " "
  End If
  
  If Fornecedor <> "0" Then
    strSQL = strSQL & "AND Cli_For.C�digo =  " & Fornecedor & " "
  End If
  
  If Vendedor <> "0" Then
    strSQL = strSQL & "AND Sa�das.Digitador = " & Vendedor & " "
  End If
  
  If Operacao <> "0" Then
    strSQL = strSQL & "AND Sa�das.Opera��o = " & Operacao & " "
  End If
  
  strSQL = strSQL & "GROUP BY [Sa�das - Produtos].[C�digo sem Grade], Produtos.Nome "
  strSQL = strSQL & "ORDER BY Produtos.Nome"
  
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  dbTemp.Execute "DELETE * FROM tblRelVendasII", dbFailOnError
  
  Set rsRelVendas = dbTemp.OpenRecordset("tblRelVendasII")
  
  With rsSaidas
  
    If (.BOF And .EOF) Then
      If Not BarraProgresso Is Nothing Then
        BarraProgresso.min = 0
        BarraProgresso.Max = 1
        BarraProgresso.Value = 0
      End If
      MsgBox "N�o h� informa��es para serem exibidas no relat�rio. Verifique se os filtros foram preenchidos corretamente.", vbInformation, "Quick Store"
      g_blnRelVendasII = False
      
      rsSaidas.Close
      rsRelVendas.Close
      
      Set rsSaidas = Nothing
      Set rsRelVendas = Nothing
      
      Exit Function
    End If
    
    .MoveLast
    .MoveFirst
    
    If Not BarraProgresso Is Nothing Then
      BarraProgresso.min = 0
      BarraProgresso.Max = .RecordCount + 1
    End If
    
    Do Until .EOF
      '15/08/2008 - mpdea
      'Realiza DoEvents somente em intervalos para n�o demorar a consulta em eventos de atualiza��o
      If .AbsolutePosition Mod 100 = 0 Then
        DoEvents
      End If
      
      strSQL = "SELECT Last([Entradas - Produtos].ICM) AS ICMS, Last([Entradas - Produtos].Pre�o) AS Preco "
      strSQL = strSQL & "FROM (Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequ�ncia = [Entradas - Produtos].Sequ�ncia)) INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo "
      strSQL = strSQL & "Where [Entradas - Produtos].C�digo='" & .Fields("C�digo Sem Grade") & "' And [Opera��es Entrada].Tipo = 'C' AND [Entradas - Produtos].Filial=" & Filial
      '31/03/2008 - mpdea
      'Comentado o c�digo abaixo conforme solicita��o do cliente
      'para que n�o filtre a entrada pelo per�odo informado, mas sempre pela �ltima
      'strSQL = strSQL & " AND (Data >=#" & Format(DataInicial, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal, "mm/dd/yyyy") & "#) "
      strSQL = strSQL & " GROUP BY [Entradas - Produtos].C�digo, [Opera��es Entrada].Tipo, [Entradas - Produtos].Filial "
      strSQL = strSQL & "ORDER BY Last(Entradas.Sequ�ncia)"

      '28/10/2008 - mpdea
      'Alterado para dynaset, read only (mais r�pido)
      Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

      rsRelVendas.AddNew
      rsRelVendas("Codigo").Value = .Fields("C�digo Sem Grade")
      rsRelVendas("Descricao").Value = "" & .Fields("Nome")
      If (rsEntradas.BOF And rsEntradas.EOF) Then
        '28/10/2008 - mpdea
        'Caso n�o encontre uma movimenta��o de compra, os dados de custos
        'ser�o obtidos a partir do cadastro de produtos e tabela de pre�os
        'Solicitado pelo cliente LL Ferramentas nesta data
        strSQL = "SELECT [Percentual Icm Entrada], Pre�o "
        strSQL = strSQL & "FROM Produtos INNER JOIN Pre�os ON Produtos.C�digo = Pre�os.Produto "
        strSQL = strSQL & "WHERE Tabela = 'CUSTO' AND C�digo = '" & .Fields("C�digo Sem Grade") & "'"
        Set rsAux = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
        If Not (rsAux.BOF And rsAux.EOF) Then
          rsRelVendas("ICMS").Value = rsAux.Fields("Percentual Icm Entrada").Value
          '07/11/2008 - mpdea
          'Inclu�do multiplica��o pela quantidade
          rsRelVendas("PrecoCusto").Value = .Fields("Quantidade") * rsAux.Fields("Pre�o").Value
        Else
          rsRelVendas("ICMS").Value = 0
          rsRelVendas("PrecoCusto").Value = 0
        End If
        rsAux.Close
        Set rsAux = Nothing
      Else
        rsRelVendas("ICMS").Value = rsEntradas("ICMS")
        rsRelVendas("PrecoCusto").Value = .Fields("Quantidade") * rsEntradas("Preco")
      End If
      rsRelVendas("ValorVenda").Value = "" & .Fields("ValorVenda")
      rsRelVendas("Quantidade").Value = "" & .Fields("Quantidade")
      rsRelVendas("ValorICMSVenda").Value = "" & .Fields("ICMSVenda")
      rsRelVendas.Update
      
      rsEntradas.Close
      Set rsEntradas = Nothing
      
      If Not BarraProgresso Is Nothing Then
        BarraProgresso.Value = .AbsolutePosition
      End If
      .MoveNext
    Loop
    rsRelVendas.Close
    .Close
  End With
  
  If Not BarraProgresso Is Nothing Then
    BarraProgresso.min = 0
    BarraProgresso.Max = 1
    BarraProgresso.Value = 0
  End If
  
  Set rsRelVendas = Nothing
  Set rsSaidas = Nothing
  
  g_blnRelVendasII = True
  
  Exit Function
  
TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Set rsRelVendas = Nothing
  Set rsSaidas = Nothing
  g_blnRelVendasII = False

End Function

Public Sub Importa_Gesto()
  Dim GestoBD As Database
  Dim Cfisc_Pgto As Recordset
  Dim TipoRecebimpgto As Recordset
  Dim Cfisc_Base As Recordset
  Dim Caixa As Recordset
  Dim CaixaAnterior As Recordset
  Dim Resumo_Di�rio_Financeiro As Recordset
  Dim Resumo_Di�rio As Recordset
  Dim Contas_Receber As Recordset
  Dim produtos As Recordset
  Dim cad_prod As Recordset
  Dim Estoque_Final As Recordset
  Dim Estoque As Recordset
  Dim Estoque_Anterior As Recordset
  Dim Saidas As Recordset
  Dim saidas_prod As Recordset
  Dim Parametros As Recordset
  Dim nSequencia As String
  Dim Cliente As Recordset
  Dim BaseICMS As Integer
  Dim ValorICMS As Integer
  Dim IPI As Integer
  Dim produtosGesto As Recordset
  
  If frmParametros.VerificaPAF = True Then
    
    IPI = 0
    BaseICMS = 0
    ValorICMS = 0
    Set rsParametros = db.OpenRecordset("Select * from [Par�metros Filial] Where Filial = " & gnCodFilial & "")
    Set GestoBD = OpenDatabase(rsParametros("BancoPDV").Value & "\Gesto.mde", False, False)
    Set Cfisc_Base = GestoBD.OpenRecordset("Select * from Cfisc_base where Importado_retaguarda = Falso and FIS_CANCELADO = Falso")
    Do Until Cfisc_Base.EOF
      Set produtos = GestoBD.OpenRecordset("Select * from Cfisc_Item where FIS_CONTROL = " & Cfisc_Base("FIS_CONTROL") & "")
      Do Until produtos.EOF
        Set produtosGesto = GestoBD.OpenRecordset("Select * from ItemEstoque Where CODIGO_FORNECEDOR = " & produtos("FIS_CODIGO") & "")
        Set cad_prod = db.OpenRecordset("Select * from Produtos Where C�digo = " & produtos("FIS_CODIGO") & "")
        IPI = IPI + (cad_prod("Percentual IPI") * produtos("FIS_TOTALITEM")) / 100
        If produtosGesto("situacaoTributaria") = "T" Then
          BaseICMS = BaseICMS + produtos("FIS_TOTALITEM")
          ValorICMS = ValorICMS + (BaseICMS * cad_prod("Percentual ICM")) / 100
        End If
        produtos.MoveNext
      Loop
      Set Parametros = db.OpenRecordset("Select * from Par�metros Filial where Filial = " & gnCodFilial & "")
      nSequencia = gnGetNextSequencia(gnCodFilial) 'gera proxima sequencia
      Parametros.Edit
      Parametros("�ltima Movimenta��o") = nSequencia
      Parametros.Update
      Set Cliente = db.OpenRecordset("Select * from Cli_For where Nome  = " & Cfisc_Base("FIS_CLIENTE") & "")
      Set Saidas = db.OpenRecordset("Sa�das")
      Saidas.AddNew
      Saidas("Filial") = gnCodFilial
      Saidas("Data") = Format$(Data_Atual, "dd/mm/yyyy")
      Saidas("Sequ�ncia") = nSequencia
      Saidas("Opera��o") = Parametros("VR C�digo Opera��o")
      Saidas("Caixa") = 1
      Saidas("Tabela") = Parametros("Tabela 1")
      Saidas("Digitador") = Cfisc_Base("CODIGO_ATENDENTE")
      Saidas("Operador") = Cfisc_Base("CODIGO_ATENDENTE")
      Saidas("Cliente") = Cliente("C�digo")
      Saidas("Produtos") = Cfisc_Base("FIS_TOTALVENDA")
      Saidas("Desconto") = Cfisc_Base("DESCONTO")
    Cfisc_Base.MoveNext
    Loop
  
  End If

End Sub

'14/11/2014 - Eduardo
'Gera��o de dados para relat�rio de Vendas por Vendedor
'Solicitante: Info Social
Public Function g_blnRelVendasPorVendedor(ByVal Filial As Byte, ByVal Vendedor As String, ByVal Operacao As String, ByVal DataInicial1 As Date, ByVal DataFinal1 As Date, ByVal DataInicial2 As Date, ByVal DataFinal2 As Date, ByVal DataInicial3 As Date, ByVal DataFinal3 As Date, Optional ByRef BarraProgresso As ProgressBar) As Boolean
On Error GoTo TratarErro:

  Dim strSQL As String          'Monta a string de consulta SQL para gera��o dos dados
  Dim rsSaidas As Recordset     'Abre a tabela de Saidas
  Dim rsRelVendas As Recordset ' abre a tabela tempor�ria para adi��o e dados
  Dim rsTblVendasVendedor As Recordset
  Dim rsSaidas2 As Recordset
  Dim SomaMes As Double
  Dim strSQL2 As String
  
  SomaMes = 0
  
  strSQL = "SELECT Filial, Cli_For.Vendedor, Opera��o, Cliente, SUM(Total) AS SumMes1 FROM Sa�das "
  
  strSQL2 = "SELECT DISTINCT (Sequ�ncia),Filial, Cli_For.Vendedor, Opera��o, Cliente, Total FROM Sa�das "
  
  strSQL = strSQL & "INNER JOIN Cli_For ON Cli_For.Vendedor = Sa�das.Digitador "
  
  strSQL2 = strSQL2 & "INNER JOIN Cli_For ON Cli_For.Vendedor = Sa�das.Digitador "
  
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND (Data >=#" & Format(DataInicial1, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal1, "mm/dd/yyyy") & "#) "

  strSQL2 = strSQL2 & "WHERE Filial = " & Filial & " AND (Data >=#" & Format(DataInicial1, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal1, "mm/dd/yyyy") & "#) "
    
  If Vendedor <> "0" Then
    strSQL = strSQL & "AND Cli_For.Vendedor = " & Vendedor & " "
    strSQL2 = strSQL2 & "AND Cli_For.Vendedor = " & Vendedor & " "
  End If
  
  If Operacao <> "0" Then
    strSQL = strSQL & "AND Opera��o = " & Operacao & " "
    strSQL2 = strSQL2 & "AND Opera��o = " & Operacao & " "
  End If
  
  strSQL = strSQL & "AND Cliente IN (SELECT C�digo FROM Cli_For WHERE Vendedor <> 0 AND Inativo = false AND Tipo = 'C') "
  
  strSQL2 = strSQL2 & "AND Cliente IN (SELECT C�digo FROM Cli_For WHERE Vendedor <> 0 AND Inativo = false AND Tipo = 'C') "
  
  strSQL = strSQL & "GROUP BY Filial, Cli_For.Vendedor, Opera��o, Cliente "
  strSQL = strSQL & "ORDER BY Cliente"
  
  strSQL2 = strSQL2 & "ORDER BY Cliente"
  
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  Set rsSaidas2 = db.OpenRecordset(strSQL2, dbOpenDynaset, dbReadOnly)
  
  dbTemp.Execute "DELETE * FROM tblRelVendasPorVendedor", dbFailOnError
  
  Set rsRelVendas = dbTemp.OpenRecordset("tblRelVendasPorVendedor")
  
  With rsSaidas
  
    If (.BOF And .EOF) Then
      If Not BarraProgresso Is Nothing Then
        BarraProgresso.min = 0
        BarraProgresso.Max = 1
        BarraProgresso.Value = 0
      End If
      MsgBox "N�o h� informa��es para serem exibidas no relat�rio. Verifique se os filtros foram preenchidos corretamente.", vbInformation, "Quick Store"
      g_blnRelVendasPorVendedor = False

      rsSaidas.Close
      rsRelVendas.Close

      Set rsSaidas = Nothing
      Set rsRelVendas = Nothing

      Exit Function
    End If
    
    .MoveLast
    .MoveFirst
    
    If Not BarraProgresso Is Nothing Then
      BarraProgresso.min = 0
      BarraProgresso.Max = .RecordCount + 1
    End If
    
    Do Until .EOF
      'Realiza DoEvents somente em intervalos para n�o demorar a consulta em eventos de atualiza��o
      If .AbsolutePosition Mod 100 = 0 Then
        DoEvents
      End If
      
      rsSaidas2.MoveFirst
      
      Do Until rsSaidas2.EOF
        If rsSaidas("Cliente").Value = rsSaidas2("Cliente").Value Then
          SomaMes = SomaMes + rsSaidas2("Total")
        End If
        rsSaidas2.MoveNext
      Loop
      
      'Alterado para dynaset, read only (mais r�pido)
      'Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      rsRelVendas.AddNew
      rsRelVendas("Filial").Value = .Fields("Filial")
      rsRelVendas("Vendedor").Value = "" & .Fields("Vendedor")
      rsRelVendas("DataIni1").Value = "" & DataInicial1
      rsRelVendas("DataFim1").Value = "" & DataFinal1
      rsRelVendas("DataIni2").Value = "" & DataInicial2
      rsRelVendas("DataFim2").Value = "" & DataFinal2
      rsRelVendas("DataIni3").Value = "" & DataInicial3
      rsRelVendas("DataFim3").Value = "" & DataFinal3
      rsRelVendas("Operacao").Value = "" & .Fields("Opera��o")
      rsRelVendas("Cliente").Value = "" & .Fields("Cliente")
      rsRelVendas("SumMes1").Value = "" & SomaMes
      rsRelVendas("SumMes2").Value = "0"
      rsRelVendas("SumMes3").Value = "0"
      rsRelVendas("SumMeses").Value = "" & SomaMes
      rsRelVendas.Update
      
      'rsSaidas.Close
      
      SomaMes = 0
      
      If Not BarraProgresso Is Nothing Then
        BarraProgresso.Value = .AbsolutePosition
      End If
      .MoveNext
    Loop
    
    Set rsSaidas = Nothing
    
    Set rsSaidas2 = Nothing
    
      strSQL = ""
      
      strSQL2 = ""

      strSQL = "SELECT Filial, Cli_For.Vendedor, Opera��o, Cliente, SUM(Total) AS SumMes2 FROM Sa�das "
      
      strSQL2 = "SELECT DISTINCT (Sequ�ncia),Filial, Cli_For.Vendedor, Opera��o, Cliente, Total FROM Sa�das "
      
      strSQL = strSQL & "INNER JOIN Cli_For ON Cli_For.Vendedor = Sa�das.Digitador "
      
      strSQL2 = strSQL2 & "INNER JOIN Cli_For ON Cli_For.Vendedor = Sa�das.Digitador "

      strSQL = strSQL & "WHERE Filial = " & Filial & " AND (Data >=#" & Format(DataInicial2, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal2, "mm/dd/yyyy") & "#) "
      
      strSQL2 = strSQL2 & "WHERE Filial = " & Filial & " AND (Data >=#" & Format(DataInicial2, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal2, "mm/dd/yyyy") & "#) "

      If Vendedor <> "0" Then
        strSQL = strSQL & "AND Cli_For.Vendedor = " & Vendedor & " "
        strSQL2 = strSQL2 & "AND Cli_For.Vendedor = " & Vendedor & " "
      End If

      If Operacao <> "0" Then
        strSQL = strSQL & "AND Opera��o = " & Operacao & " "
        strSQL2 = strSQL2 & "AND Opera��o = " & Operacao & " "
      End If

      strSQL = strSQL & "AND Cliente IN (SELECT C�digo FROM Cli_For WHERE Vendedor <> 0 AND Inativo = false AND Tipo = 'C') "
      
      strSQL2 = strSQL2 & "AND Cliente IN (SELECT C�digo FROM Cli_For WHERE Vendedor <> 0 AND Inativo = false AND Tipo = 'C') "
      
      strSQL = strSQL & "GROUP BY Filial, Cli_For.Vendedor, Opera��o, Cliente "
      strSQL = strSQL & "ORDER BY Cliente"
      
      strSQL2 = strSQL2 & "ORDER BY Cliente"
      
      Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot, dbReadOnly)
      Set rsSaidas2 = db.OpenRecordset(strSQL2, dbOpenDynaset, dbReadOnly)
      
    Do Until rsSaidas.EOF
      If .AbsolutePosition Mod 100 = 0 Then
        DoEvents
      End If
      
      rsSaidas2.MoveFirst
      
      Do Until rsSaidas2.EOF
        If rsSaidas("Cliente").Value = rsSaidas2("Cliente").Value Then
          SomaMes = SomaMes + rsSaidas2("Total")
        End If
        rsSaidas2.MoveNext
      Loop
      
      
      strSQL = ""
      strSQL = "SELECT * FROM tblRelVendasPorVendedor WHERE Filial = " & rsSaidas("Filial").Value & " AND Vendedor = " & rsSaidas("Vendedor").Value
      strSQL = strSQL & " AND Operacao = " & rsSaidas("Opera��o").Value & " AND Cliente = " & rsSaidas("Cliente").Value & ""
      
      Set rsTblVendasVendedor = dbTemp.OpenRecordset(strSQL)
      
      If rsTblVendasVendedor.EOF Then
        rsTblVendasVendedor.AddNew
        rsTblVendasVendedor("Filial").Value = rsSaidas("Filial")
        rsTblVendasVendedor("Vendedor").Value = "" & rsSaidas("Vendedor")
        rsTblVendasVendedor("DataIni1").Value = "" & DataInicial1
        rsTblVendasVendedor("DataFim1").Value = "" & DataFinal1
        rsTblVendasVendedor("DataIni2").Value = "" & DataInicial2
        rsTblVendasVendedor("DataFim2").Value = "" & DataFinal2
        rsTblVendasVendedor("DataIni3").Value = "" & DataInicial3
        rsTblVendasVendedor("DataFim3").Value = "" & DataFinal3
        rsTblVendasVendedor("Operacao").Value = "" & rsSaidas("Opera��o")
        rsTblVendasVendedor("Cliente").Value = "" & rsSaidas("Cliente")
        rsTblVendasVendedor("SumMes1").Value = "0"
        rsTblVendasVendedor("SumMes2").Value = "" & SomaMes
        rsTblVendasVendedor("SumMes3").Value = "0"
        rsTblVendasVendedor("SumMeses").Value = "" & SomaMes
        rsTblVendasVendedor.Update
      Else
        rsTblVendasVendedor.Edit
        rsTblVendasVendedor("SumMes2") = "" & SomaMes
        rsTblVendasVendedor("SumMeses").Value = rsTblVendasVendedor("SumMes1") + SomaMes
        rsTblVendasVendedor.Update
      End If
      
      SomaMes = 0
      
      rsSaidas.MoveNext

    Loop
    
    Set rsTblVendasVendedor = Nothing
    
    Set rsSaidas = Nothing
    Set rsSaidas2 = Nothing
  
    strSQL = ""
    
    strSQL2 = ""

    strSQL = "SELECT Filial, Cli_For.Vendedor, Opera��o, Cliente, SUM(Total) AS SumMes3 FROM Sa�das "
    
    strSQL2 = "SELECT DISTINCT (Sequ�ncia),Filial, Cli_For.Vendedor, Opera��o, Cliente, Total FROM Sa�das "

    strSQL = strSQL & "INNER JOIN Cli_For ON Cli_For.Vendedor = Sa�das.Digitador "
    
    strSQL2 = strSQL2 & "INNER JOIN Cli_For ON Cli_For.Vendedor = Sa�das.Digitador "
    
    strSQL = strSQL & "WHERE Filial = " & Filial & " AND (Data >=#" & Format(DataInicial3, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal3, "mm/dd/yyyy") & "#) "
    
    strSQL2 = strSQL2 & "WHERE Filial = " & Filial & " AND (Data >=#" & Format(DataInicial3, "mm/dd/yyyy") & "#  AND Data <=#" & Format(DataFinal3, "mm/dd/yyyy") & "#) "

    If Vendedor <> "0" Then
      strSQL = strSQL & "AND Cli_For.Vendedor = " & Vendedor & " "
      strSQL2 = strSQL2 & "AND Cli_For.Vendedor = " & Vendedor & " "
    End If

    If Operacao <> "0" Then
      strSQL = strSQL & "AND Opera��o = " & Operacao & " "
      strSQL2 = strSQL2 & "AND Opera��o = " & Operacao & " "
    End If
    

    strSQL = strSQL & "AND Cliente IN (SELECT C�digo FROM Cli_For WHERE Vendedor <> 0 AND Inativo = false AND Tipo = 'C') "
    
    strSQL2 = strSQL2 & "AND Cliente IN (SELECT C�digo FROM Cli_For WHERE Vendedor <> 0 AND Inativo = false AND Tipo = 'C') "
    
    strSQL = strSQL & "GROUP BY Filial, Cli_For.Vendedor, Opera��o, Cliente "
    strSQL = strSQL & "ORDER BY Cliente"
    
    strSQL2 = strSQL2 & "ORDER BY Cliente"
    
    Set rsSaidas = db.OpenRecordset(strSQL, dbOpenSnapshot, dbReadOnly)
    Set rsSaidas2 = db.OpenRecordset(strSQL2, dbOpenSnapshot, dbReadOnly)
    
  Do Until rsSaidas.EOF
    If .AbsolutePosition Mod 100 = 0 Then
      DoEvents
    End If
    
    rsSaidas2.MoveFirst
      
      Do Until rsSaidas2.EOF
        If rsSaidas("Cliente").Value = rsSaidas2("Cliente").Value Then
          SomaMes = SomaMes + rsSaidas2("Total")
        End If
        rsSaidas2.MoveNext
      Loop
    
    strSQL = ""
    strSQL = "SELECT * FROM tblRelVendasPorVendedor WHERE Filial = " & rsSaidas("Filial").Value & " AND Vendedor = " & rsSaidas("Vendedor").Value
    strSQL = strSQL & " AND Operacao = " & rsSaidas("Opera��o").Value & " AND Cliente = " & rsSaidas("Cliente").Value & ""
    
    Set rsTblVendasVendedor = dbTemp.OpenRecordset(strSQL)
    
    If rsTblVendasVendedor.EOF Then
      rsTblVendasVendedor.AddNew
      rsTblVendasVendedor("Filial").Value = rsSaidas("Filial")
      rsTblVendasVendedor("Vendedor").Value = "" & rsSaidas("Vendedor")
      rsTblVendasVendedor("DataIni1").Value = "" & DataInicial1
      rsTblVendasVendedor("DataFim1").Value = "" & DataFinal1
      rsTblVendasVendedor("DataIni2").Value = "" & DataInicial2
      rsTblVendasVendedor("DataFim2").Value = "" & DataFinal2
      rsTblVendasVendedor("DataIni3").Value = "" & DataInicial3
      rsTblVendasVendedor("DataFim3").Value = "" & DataFinal3
      rsTblVendasVendedor("Operacao").Value = "" & rsSaidas("Opera��o")
      rsTblVendasVendedor("Cliente").Value = "" & rsSaidas("Cliente")
      rsTblVendasVendedor("SumMes1").Value = "0"
      rsTblVendasVendedor("SumMes2").Value = "0"
      rsTblVendasVendedor("SumMes3").Value = "" & SomaMes
      rsTblVendasVendedor("SumMeses").Value = "" & SomaMes
      rsTblVendasVendedor.Update
    Else
      rsTblVendasVendedor.Edit
      rsTblVendasVendedor("SumMes3") = "" & SomaMes
      rsTblVendasVendedor("SumMeses").Value = "" & SomaMes + rsTblVendasVendedor("SumMes1") + rsTblVendasVendedor("SumMes2")
      rsTblVendasVendedor.Update
    End If
    
    SomaMes = 0
    
    rsSaidas.MoveNext

  Loop
    Dim rsCarteira As Recordset
    
    strSQL = ""
  If Vendedor <> 0 Then
    strSQL = "SELECT C�digo, Vendedor FROM Cli_For WHERE Vendedor = " & Vendedor & " AND Inativo = false AND Tipo = 'C' ORDER BY Vendedor"
  Else
    strSQL = "SELECT C�digo, Vendedor FROM Cli_For WHERE Vendedor <> 0 AND Inativo = false AND Tipo = 'C' ORDER BY Vendedor"
  End If
      Set rsCarteira = db.OpenRecordset(strSQL)
    Do Until rsCarteira.EOF
     strSQL = ""
     strSQL = "SELECT * FROM tblRelVendasPorVendedor WHERE Cliente = " & rsCarteira("C�digo") & " AND Vendedor = " & rsCarteira("Vendedor")
     Set rsRelVendas = Nothing
     Set rsRelVendas = dbTemp.OpenRecordset(strSQL)
     
     If rsRelVendas.EOF Then
        rsRelVendas.AddNew
        rsRelVendas("Filial").Value = Filial
        rsRelVendas("Vendedor").Value = "" & rsCarteira("Vendedor")
        rsRelVendas("DataIni1").Value = "" & DataInicial1
        rsRelVendas("DataFim1").Value = "" & DataFinal1
        rsRelVendas("DataIni2").Value = "" & DataInicial2
        rsRelVendas("DataFim2").Value = "" & DataFinal2
        rsRelVendas("DataIni3").Value = "" & DataInicial3
        rsRelVendas("DataFim3").Value = "" & DataFinal3
        rsRelVendas("Operacao").Value = Operacao
        rsRelVendas("Cliente").Value = "" & rsCarteira("C�digo")
        rsRelVendas("SumMes1").Value = "0"
        rsRelVendas("SumMes2").Value = "0"
        rsRelVendas("SumMes3").Value = "0"
        rsRelVendas("SumMeses").Value = "0"
        rsRelVendas.Update
     End If
     rsCarteira.MoveNext
    Loop
  
    rsRelVendas.Close
    .Close
  End With
  
  If Not BarraProgresso Is Nothing Then
    BarraProgresso.min = 0
    BarraProgresso.Max = 1
    BarraProgresso.Value = 0
  End If
  
  Set rsRelVendas = Nothing
  Set rsSaidas = Nothing
  
  g_blnRelVendasPorVendedor = True
  
  Exit Function
  
TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Set rsRelVendas = Nothing
  Set rsSaidas = Nothing
  g_blnRelVendasPorVendedor = False

End Function

Public Function blnPermissaoAlterarPrecos(ByVal CodFunc As String) As Boolean
'09-07-2015 Jean Ricardo Zanella - Fun��o para verificar se usuario tem permiss�o para alterar pre�os
  Dim rstFunc As Recordset
  
  On Error GoTo TratarErro
  
  Set rstFunc = db.OpenRecordset("SELECT * FROM Acessos WHERE Programa = '" & "ALTERA PRE�OS" & "'" & " AND Usu�rio = " & CodFunc & " AND Gravar = True", dbOpenDynaset)

  If rstFunc.RecordCount = 0 Then
    blnPermissaoAlterarPrecos = False
  Else
    blnPermissaoAlterarPrecos = True
  End If

  rstFunc.Close
  Set rstFunc = Nothing

  Exit Function

TratarErro:
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  blnPermissaoAlterarPrecos = False

End Function
