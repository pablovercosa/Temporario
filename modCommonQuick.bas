Attribute VB_Name = "modCommonQuick"
Option Explicit


'---------------------------------------------------------------
' DATA: 14/06/2022
' AUTOR: Pablo Ver�osa Silva
' MUDAN�AS:
'    1) Incluir par�metros de recebimento de parcelas e cheques
'    2) Ampliar o limite de parcelas e cheques para 3 d�gitos
'---------------------------------------------------------------
Public pab_VR_Qtde_Parcela As Integer
Public pab_VR_Qtde_Cheques As Integer
'---------------------------------------------------------------


'-------------------------------------------------------------------------------------
'Fun��o gstrGetCliForNumeroDocumento
'
'Obt�m o n�mero do documento (CPF/CNPJ) do cliente/fornecedor na tabela Cli_For
'
'29/04/2008 - mpdea
'-------------------------------------------------------------------------------------

Public Function gstrGetCliForNumeroDocumento(ByVal lngCodigo As Long) As String
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT CGC FROM Cli_For WHERE C�digo = " & lngCodigo
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not .BOF And Not .EOF Then
      '15/05/2009 - mpdea
      'Corrigido RT-94 (Invalid use of the null)
      gstrGetCliForNumeroDocumento = .Fields("CGC").Value & ""
    End If
    .Close
  End With
  Set rs = Nothing
  
End Function

'29/04/2008 - mpdea
'Verifica exig�ncia e obt�m n�mero de documento (CPF ou CNPJ)
Public Function g_str_GetNumeroDocumento(ByVal intCodigoOperacao As Integer, ByVal lngCodigoCliente As Long, ByVal strNumeroDocumentoDefault As String) As String
  Dim rs As Recordset
  Dim bln_exibir_tela As Boolean
  Dim str_ret As String
  
  Set rs = db.OpenRecordset("SELECT ExibirTelaNumeroDocumento FROM [Opera��es Sa�da] WHERE C�digo = " & intCodigoOperacao, dbOpenDynaset, dbReadOnly)
  With rs
    If .RecordCount > 0 Then
      bln_exibir_tela = .Fields("ExibirTelaNumeroDocumento").Value
    End If
    .Close
  End With
  Set rs = Nothing
  
  If bln_exibir_tela Then
    str_ret = frmNumeroDocumento.Start(lngCodigoCliente, strNumeroDocumentoDefault)
  End If
  
  'Retorno da fun��o
  g_str_GetNumeroDocumento = str_ret
End Function

'29/04/2008 - mpdea
'Retorna somente n�meros de um texto
Public Function g_str_SomenteNumero(ByVal Texto As String) As String
  Dim X As Integer
  Dim str_ret As String
  Dim int_c As Integer
  
  If Len(Texto) = 0 Then Exit Function
  
  For X = 1 To Len(Texto)
    int_c = Asc(Mid(Texto, X, 1))
    If int_c >= 48 And int_c <= 57 Then
      str_ret = str_ret & Mid(Texto, X, 1)
    End If
  Next X
  
  g_str_SomenteNumero = str_ret
End Function

'11/06/2008 - mpdea
'Obt�m o valor base para c�lculo de impostos de servi�os
Public Function g_dbl_ValorBaseCalculoImpostosServicos(ByVal Filial As Integer, ByVal Cliente As Long, ByVal ValorIsencaoPisCofinsCsll As Double, ByVal TotalServicosVenda As Double) As Double
  Dim rs As Recordset
  Dim str_sql As String
  Dim dbl_total_mes As Double
  Dim dbl_base_calculo As Double
  
  str_sql = "SELECT Sum(Servi�os) as TotalServicosMes FROM Sa�das" 'Total de servi�os
  'str_sql = str_sql & " INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo"
  str_sql = str_sql & " WHERE Filial = " & Filial 'Tipo = 'V' And
  str_sql = str_sql & " And Cliente = " & Cliente
  str_sql = str_sql & " And Data Between #" & Month(CDate(Data_Atual)) & "/1/" & Year(CDate(Data_Atual)) & "#" 'Primeiro dia do m�s
  str_sql = str_sql & " And #" & Format(CDate(Data_Atual), "MM/dd/yyyy") & "#" 'Dia atual
  str_sql = str_sql & " And Efetivada And Not [Movimenta��o Desfeita]"
  
  Set rs = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  With rs
    If Not (.BOF And .EOF) Then
      Call IsDataType(dtDouble, .Fields("TotalServicosMes").Value, dbl_total_mes)
    End If
    .Close
  End With
  Set rs = Nothing
  
  'An�lise de condi��es
  '
  '1) Calcular sobre o valor total de servi�os da nota se o total do m�s iguala
  'ou ultrapassa o valor de isen��o (indica que j� houve c�lculo anterior)
  If dbl_total_mes >= ValorIsencaoPisCofinsCsll Then
    dbl_base_calculo = TotalServicosVenda
  '2) Calcular sobre o valor total do m�s mais o valor total de servi�os da nota
  'se o valor total de servi�os do m�s n�o iguala ou ultrapassa o valor de isen��o,
  'mas com a soma do valor total de servi�os da nota iguala ou ultrapassa o valor de isen��o
  ElseIf dbl_total_mes < ValorIsencaoPisCofinsCsll And (dbl_total_mes + TotalServicosVenda >= ValorIsencaoPisCofinsCsll) Then
    dbl_base_calculo = dbl_total_mes + TotalServicosVenda
  '3) Isento se o valor total do m�s mais o valor total de servi�os da nota n�o atingir
  'o valor de isen��o
  ElseIf (dbl_total_mes + TotalServicosVenda) < ValorIsencaoPisCofinsCsll Then
    dbl_base_calculo = 0 'Isento
  End If
  
  'Retorna base de c�lculo
  g_dbl_ValorBaseCalculoImpostosServicos = dbl_base_calculo
  
End Function
