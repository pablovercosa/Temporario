Attribute VB_Name = "modFunctions"
'Fun��es adicionadas, alteradas e implementadas por mpdea
Option Explicit

Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const FORMAT_VALUE As String = "#,###,###,##0.00" 'Formata��o de valores para exibi��o

Public Const SQL_DATE_MASK As String = "MM/DD/YYYY"

Enum enFieldType
  ftNumero = 1
  ftTexto = 2
  ftData = 3
End Enum

Enum TipoMargem
  tmSuperior = 1
  tmEsquerda = 2
End Enum

Public Enum enTableMovimentType
  tmEntradas = 1
  tmEntradasProdutos = 2
  tmSaidas = 3
  tmSaidasProdutos = 4
  tmSaidasServicos = 5
  tmMovimentoCheques = 6
  tmMovimentoParcelas = 7
  '10/12/2009 - Andrea
  tmMovimentoCartoes = 8
End Enum

'Tipos para navega��o de registros em um recordset
Public Enum enNavigate
  navFirst = 1
  navNext
  navPrevious
  navLast
End Enum

'Tipos de dados utilizados na fun��o IsDataType
Public Enum enDataType
  dtByte = 1
  dtInteger = 2
  dtLong = 3
  dtSingle = 4
  dtDouble = 5
  dtCurrency = 6
  dtDecimal = 7
  dtDate = 8
  dtBoolean = 9
  dtString = 10
End Enum

'08/10/2002 - mpdea
'Type para verifica��o de estoque
Public Type CheckStock
  strCode As String
  dblQuantity As Double
  dblStock As Double
  blnStockInsufficient As Boolean
End Type
'
Public typCheckStock() As CheckStock


'------------------------------------------------------------------------------
'31/01/2006 - mpdea
'
'Mensagem para Nota Fiscal
'
'Tipo de dados para armazenar os dados de produto da movimenta��o
Private Type MsgNF_Produto
  strCodigo As String
  intClasse As Integer
  intSubClasse As Integer
  intGrupoFiscal As Integer
End Type
'
'Tipo de dados para armazenar os dados da movimenta��o
Private Type MsgNF_Movimentacao
  strUF As String
  intCodigoOpSaida As Integer
  intGrupoFiscalOpSaida As Integer
  clcProdutos() As MsgNF_Produto
End Type
'------------------------------------------------------------------------------

'Verifica se o valor passado � do tipo esperado, retornando o valor correto
Public Function IsDataType(ByVal DataType As enDataType, ByVal varValue As Variant, _
  Optional ByRef varRet As Variant = "") As Boolean
  
  On Error Resume Next
  
  Select Case DataType
    Case dtBoolean
      varRet = False
      varRet = CBool(varValue)
    Case dtByte
      varRet = 0
      varRet = CByte(varValue)
    Case dtCurrency
      varRet = 0@
      varRet = CCur(varValue)
    Case dtDate
      varRet = 0
      varRet = CDate(varValue)
    Case dtDouble
      varRet = 0#
      varRet = CDbl(varValue)
    Case dtDecimal
      varRet = 0
      varRet = CDec(varValue)
    Case dtInteger
      varRet = 0
      varRet = CInt(varValue)
    Case dtLong
      varRet = 0&
      varRet = CLng(varValue)
    Case dtSingle
      varRet = 0!
      varRet = CSng(varValue)
    Case dtString
      varRet = ""
      varRet = CStr(varValue)
  End Select
  
  IsDataType = (Err.Number = 0)
  
End Function

'Obt�m o nome atrav�s do c�digo
Public Function gsGetNameFilial(ByVal nCodigo As Integer) As String
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Nome FROM [Par�metros Filial] WHERE Filial = " & nCodigo, dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gsGetNameFilial = IIf(IsNull(!Nome), "", !Nome)
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Obt�m o nome atrav�s do c�digo
Public Function gsGetNameProduto(ByVal sCodigo As String) As String
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Nome FROM Produtos WHERE C�digo = '" & sCodigo & "'", dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gsGetNameProduto = IIf(IsNull(!Nome), "", !Nome)
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Obt�m o nome atrav�s do c�digo
Public Function gsGetNameClasse(ByVal nCodigo As Integer) As String
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Nome FROM Classes WHERE C�digo = " & nCodigo, dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gsGetNameClasse = IIf(IsNull(!Nome), "", !Nome)
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Obt�m o nome atrav�s do c�digo
Public Function gsGetNameSubClasse(ByVal nCodigo As Integer) As String
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Nome FROM [Sub Classes] WHERE C�digo = " & nCodigo, dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gsGetNameSubClasse = IIf(IsNull(!Nome), "", !Nome)
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Obt�m o nome atrav�s do c�digo
Public Function gsGetNameCor(ByVal nCodigo As Integer) As String
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Nome FROM Cores WHERE C�digo = " & nCodigo, dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gsGetNameCor = IIf(IsNull(!Nome), "", !Nome)
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Obt�m o nome atrav�s do c�digo
Public Function gsGetNameTamanho(ByVal nCodigo As Integer) As String
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Nome FROM Tamanhos WHERE C�digo = " & nCodigo, dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gsGetNameTamanho = IIf(IsNull(!Nome), "", !Nome)
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Verifica se existe em determinado campo da tabela o valor passado
Public Function gbCheckValueInTable(ByVal sTable As String, ByVal sField As String, ByVal nFieldType As enFieldType, ByVal sValue As String) As Boolean
  Dim rsCheck As Recordset
  
  Select Case nFieldType
    Case ftData
      sValue = "#" & Format(CDate(sValue), "mm/dd/yyyy") & "#"
    Case ftNumero
      sValue = Val(sValue)
    Case ftTexto
      sValue = "'" & sValue & "'"
  End Select
  
  Set rsCheck = db.OpenRecordset("SELECT " & sField & " FROM " & sTable & " WHERE " & sField & " = " & sValue, dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      gbCheckValueInTable = True
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Obt�m determinado valor de campo da tabela selecionada
Public Function gvGetValueInTable(ByVal sTable As String, ByVal sGetField As String, _
  ByVal nGetFieldType As enFieldType, ByVal sSearchField As String, _
  ByVal nSearchFieldType As enFieldType, ByVal sSearchValue As String) As Variant
  
  Dim rsCheck As Recordset
  
  Select Case nGetFieldType
    Case ftData
      gvGetValueInTable = Empty
    Case ftNumero
      gvGetValueInTable = 0
    Case ftTexto
      gvGetValueInTable = ""
  End Select
  
  Select Case nSearchFieldType
    Case ftData
      sSearchValue = "#" & Format(CDate(sSearchValue), "mm/dd/yyyy") & "#"
    Case ftNumero
      sSearchValue = Val(sSearchValue)
    Case ftTexto
      sSearchValue = "'" & sSearchValue & "'"
  End Select
  
  Set rsCheck = db.OpenRecordset("SELECT " & sGetField & " FROM " & sTable & " WHERE " & sSearchField & " = " & sSearchValue, dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      gvGetValueInTable = .Fields(sGetField).Value
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Verifica se o produto possui grade
Public Function gbHasGrade(ByVal sCodigo As String) As Boolean
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE C�digo = '" & sCodigo & "'", dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      If UCase(IIf(IsNull(!Tipo), "", !Tipo)) = "G" Then
        gbHasGrade = True
      End If
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Verifica se o produto � fracion�rio
Public Function gbIsFrac(ByVal sCodigo As String, ByRef nQtdeCasaDec As Integer) As Boolean
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Fracionado, QtdeCasasDecimais FROM Produtos WHERE C�digo = '" & sCodigo & "'", dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      If !Fracionado Then
        gbIsFrac = True
        nQtdeCasaDec = IIf(IsNull(!QtdeCasasDecimais), 0, !QtdeCasasDecimais)
      End If
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Exibe mensagem na barra de status
Public Sub StatusMsg(ByVal sMsg As String)
  '27/01/2009 - mpdea
  'Atualizado objeto StatusBar
  If Not frmMain Is Nothing Then
    frmMain.CommandBars.StatusBar.Pane(0).Text = sMsg
  End If
End Sub

'Permite somente a digita��o de n�meros
Public Function gnSomenteNumero(ByVal nKeyAscii As Integer) As Integer
  If (nKeyAscii > vbKey9 Or nKeyAscii < vbKey0) And nKeyAscii <> vbKeyBack Then
    gnSomenteNumero = 0
  Else
    gnSomenteNumero = nKeyAscii
  End If
End Function

'Permite somente a digita��o de valores (n�mero, ponto, v�rgula e tra�o)
Public Function gnSomenteValor(ByVal nChr As Integer) As Integer
  If nChr <> vbKeyBack Then
    Select Case Chr(nChr)
      Case "0" To "9", ".", ",", "-"
        'OK
      Case Else
        nChr = 0
    End Select
  End If
  gnSomenteValor = nChr
End Function

'N�o permite a digita��o de caracteres espec�ficos
Public Function gnTypeValidKey(ByVal nKeyAscii As Integer) As Integer
  If InStr("'|", Chr(nKeyAscii)) > 0 Then
    Beep
    nKeyAscii = 0
  End If
  gnTypeValidKey = nKeyAscii
End Function

'Limita a quantidade de caracteres digitados em um controle. Opcional somente n�mero
Public Function gnLimitKeyPress(ByRef oText As Control, ByVal nLimit As Integer, _
  ByVal nKeyAscii As Integer, Optional ByVal bOnlyNumber As Boolean = False) As Integer
  
  If Len(oText.Text) >= nLimit Then
    If oText.SelLength = 0 And nKeyAscii <> vbKeyBack Then
      Beep
      nKeyAscii = 0
    End If
  End If
  If bOnlyNumber And nKeyAscii <> 0 Then
    nKeyAscii = gnSomenteNumero(nKeyAscii)
  End If
  gnLimitKeyPress = nKeyAscii
End Function

'Formata a cor do valor para Label, TextBox e MaskEditBox sem m�scara
Public Sub FormataValorCor(ByRef oLabelOrMaskTextBox As Control, Optional ByVal bBold As Boolean = True)
  With oLabelOrMaskTextBox
    .ForeColor = vbWindowText
    .Font.Bold = False
    If TypeOf oLabelOrMaskTextBox Is TextBox Or _
      TypeOf oLabelOrMaskTextBox Is MaskEdBox Then
      If .Text = "" Then
        .Text = Format(0, FORMAT_VALUE)
      ElseIf Not IsNumeric(.Text) Then
        .Text = Format(0, FORMAT_VALUE)
      ElseIf CDbl(.Text) = 0 Then
        .Text = Format(0, FORMAT_VALUE)
      Else
        If CDbl(.Text) < 0 Then
          .ForeColor = vbRed
        ElseIf CDbl(.Text) > 0 Then
          .ForeColor = vbBlue
        End If
        .Text = Format(CDbl(.Text), FORMAT_VALUE)
        .Font.Bold = bBold
      End If
    ElseIf TypeOf oLabelOrMaskTextBox Is Label Then
      If IsNull(.Caption) Then
        .Caption = Format(0, FORMAT_VALUE)
      ElseIf .Caption = "" Then
        .Caption = Format(0, FORMAT_VALUE)
      ElseIf Not IsNumeric(.Caption) Then
        .Caption = Format(0, FORMAT_VALUE)
      ElseIf CDbl(.Caption) = 0 Then
        .Caption = Format(0, FORMAT_VALUE)
      Else
        If CDbl(.Caption) < 0 Then
          .ForeColor = vbRed
        ElseIf CDbl(.Caption) > 0 Then
          .ForeColor = vbBlue
        End If
        .Caption = Format(CDbl(.Caption), FORMAT_VALUE)
        .Font.Bold = bBold
      End If
    End If
  End With
End Sub

'Verifica se h� movimento do caixa no dia
Public Function gbHasMovimentCaixa(ByVal nCaixa As Byte) As Boolean
  Dim rsCaixa As Recordset
  Dim sSql As String
  
  sSql = "SELECT * FROM Caixa WHERE Filial = " & gnCodFilial & " AND Caixa = " & _
    nCaixa & " AND Data = #" & Format(Data_Atual, "mm/dd/yyyy") & "#;"
  Set rsCaixa = db.OpenRecordset(sSql, dbOpenDynaset)
  With rsCaixa
    If .RecordCount > 0 Then
      'H� movimento do Caixa na data atual
      gbHasMovimentCaixa = True
    End If
    .Close
  End With
  Set rsCaixa = Nothing
End Function

'Verifica e cria, se necess�rio, a configura��o da Tabela de Pre�os
Public Sub CheckConfigTablePrice(ByVal sNomeTabela As String)
  Dim rsTabelaPreco As Recordset
  Dim sSql As String
  
  On Error GoTo ErrCheck
  
  ws.BeginTrans
  
  sSql = "SELECT * FROM [Tabela de Pre�os] WHERE Tabela = '" & sNomeTabela & "';"
  Set rsTabelaPreco = db.OpenRecordset(sSql, dbOpenDynaset)
  
  With rsTabelaPreco
    If .RecordCount = 0 Then
      .AddNew
      !Tabela = sNomeTabela
      ![Aceita Pr�] = True
      ![Prazo Pr�] = 9999
      ![Aceita Parcelamento] = True
      ![Prazo Parcelamento] = 9999
      ![Aceita Cart�o] = True
      ![Aceita Vale] = True
      ![Multiplicador Comiss�o] = 1
      ![Data Altera��o] = Format(Date, "dd/mm/yyyy")
      .Update
    End If
  End With
  
  ws.CommitTrans
  Exit Sub
  
ErrCheck:
  ws.Rollback
  '09/07/2002 - mpdea
  'Repassa o erro para a fun��o de origem
  Err.Raise Err.Number, "Configura��o da Tabela de Pre�os", Err.Description
  
End Sub

'Atualiza o novo valor do produto para a tabela selecionada na conta do cliente
Public Sub UpdateContaClientes(ByVal sTabela As String, ByVal sCodProd As String, ByVal nNewValue As Double)
  Dim sCriteria As String
  Dim rsConta_Cli As Recordset
  
  On Error GoTo ErrHandler
  
  Screen.MousePointer = vbHourglass
  
  Set rsConta_Cli = db.OpenRecordset("SELECT * FROM [Conta Cliente]", dbOpenDynaset)
  sCriteria = "TabPrecos = '" & sTabela & "' And Produto = '" & sCodProd & "'"
  
  Call ws.BeginTrans
  
  rsConta_Cli.FindFirst sCriteria
  Do While Not rsConta_Cli.NoMatch
  If rsConta_Cli("Valor") <> rsConta_Cli("Valor Pago") Then
      rsConta_Cli.Edit
      rsConta_Cli("Valor") = Round(rsConta_Cli("Qtde") * nNewValue, 2)
      rsConta_Cli("Data Altera��o") = Format(Date, "dd/mm/yyyy")
      rsConta_Cli.Update
      rsConta_Cli.FindNext sCriteria
  End If
  rsConta_Cli.FindNext sCriteria
  Loop
  Call ws.CommitTrans
  
  Screen.MousePointer = vbDefault
  rsConta_Cli.Close
  Set rsConta_Cli = Nothing
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao atualizar valores no Conta Clientes."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

'Verifica se o Form (atrav�s do seu caption) j� est� sendo exibido
Public Function gbShowWindow(ByVal sTitulo As String) As Boolean
  Dim frmWindow As Form
  
  For Each frmWindow In Forms
    With frmWindow
      If .Caption = sTitulo Then
        If .WindowState = vbMinimized Then
          .WindowState = vbNormal
        End If
        .Show
        .ZOrder
        gbShowWindow = True
        Exit For
      End If
    End With
  Next frmWindow
End Function

'Seleciona todo o texto do controle
Public Sub SelectAllText(ByRef oText As Control, _
  Optional ByVal bWithFocus As Boolean = False)
  
  On Error Resume Next
  With oText
    If bWithFocus Then .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  On Error GoTo 0
End Sub

'Formata o valor
Public Sub FormatCurrencyValue(ByRef oText As Control)
  Dim sAUX As String
  
  sAUX = "0"
  On Error Resume Next
  With oText
    If Not IsNull(.Text) Then
      If .Text <> "" Then
        If IsNumeric(.Text) Then
          sAUX = .Text
        End If
      End If
    End If
    .Text = Format(sAUX, FORMAT_VALUE)
  End With
  On Error GoTo 0
End Sub

'Converte a cor em inteiro longo no formato separado RGB
Public Sub ConvertRGB(ByVal nLongColor As Long, ByRef nRed As Byte, ByRef nGreen As Byte, ByRef nBlue As Byte)
  Dim nAuxBlue As Double
  Dim nAuxGreen As Double
  
  On Error GoTo ErrHandler
  
  nBlue = Fix((nLongColor / 256) / 256)
  nAuxBlue = CDbl(nBlue) * 256 * 256
  nGreen = Fix((nLongColor - nAuxBlue) / 256)
  nAuxGreen = CDbl(nGreen) * 256
  nRed = Fix(nLongColor - nAuxBlue - nAuxGreen)
  
  Exit Sub
  
ErrHandler:
  nRed = 255
  nGreen = 255
  nBlue = 174
  
End Sub

'Formata a data
Public Function gsFormatDate(ByVal sData As String)
  If IsDate(sData) Then
    sData = Format(sData, "dd/mm/yyyy")
    sData = Replace(sData, Mid(sData, 3, 1), "/")
  Else
    sData = "  /  /    "
  End If
  gsFormatDate = sData
End Function

'Verifica se existe a tabela de pre�os
Public Function gbCheckTabPreco(ByRef sTabela As String) As Boolean
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT Tabela FROM [Tabela de Pre�os] WHERE Tabela = '" & sTabela & "'", dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      gbCheckTabPreco = True
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Converte o path da imagem que utiliza a vari�vel de path QuickImagens
Public Function gsConvertImagePath(ByVal sPath As String) As String
  If UCase(Left(sPath, 15)) = UCase("[QuickImagens]\") Then
    sPath = gsImagePath & Right(sPath, Len(sPath) - 15)
  ElseIf UCase(Left(sPath, 14)) = UCase("[QuickImagens]") Then
    sPath = gsImagePath & Right(sPath, Len(sPath) - 14)
  End If
  gsConvertImagePath = sPath
End Function

'05/05/2004 - mpdea
'Grava e obt�m o pr�ximo n�mero de nota fiscal
Public Function g_lngNextNotaFiscal(ByVal intFilial As Integer) As Long
  Dim rstFilial As Recordset
  Dim strSQL As String
  Dim lngNotaFiscal As Long
  
  
  strSQL = "SELECT [�ltima Nota] AS NF FROM [Par�metros Filial] WHERE Filial = " & intFilial
  Set rstFilial = db.OpenRecordset(strSQL, dbOpenDynaset, 0, dbPessimistic)
  With rstFilial
    .Edit
    lngNotaFiscal = .Fields("NF").Value + 1
    .Fields("NF").Value = lngNotaFiscal
    .Update
    .Close
  End With
  Set rstFilial = Nothing
  
  g_lngNextNotaFiscal = lngNotaFiscal
  
End Function

'30/09/2009 - Andrea
'Grava e obt�m o pr�ximo n�mero de nota fiscal eletr�nica
Public Function g_lngNextNotaFiscal_e(ByVal intFilial As Integer) As Long
  Dim rstFilial As Recordset
  Dim strSQL As String
  Dim lngNotaFiscal As Long
  
  
  strSQL = "SELECT [UltimaNFe] AS NF FROM [Par�metros Filial] WHERE Filial = " & intFilial
  Set rstFilial = db.OpenRecordset(strSQL, dbOpenDynaset, 0, dbPessimistic)
  With rstFilial
    .Edit

    If Not IsNull(.Fields("NF").Value) Then
        lngNotaFiscal = .Fields("NF").Value + 1
    Else
        lngNotaFiscal = 1
    End If
    
    .Fields("NF").Value = lngNotaFiscal
    .Update
    .Close
  End With
  Set rstFilial = Nothing

  g_lngNextNotaFiscal_e = lngNotaFiscal

End Function

Public Function g_longNextNFCe(ByVal intFilial As Integer) As Long
  Dim rstFilial As Recordset
  Dim strSQL As String
  Dim lngNumNFCe As Long
  
  strSQL = "SELECT [UltimaNFCe] AS NFCe FROM [Par�metros Filial] WHERE Filial = " & intFilial
  Set rstFilial = db.OpenRecordset(strSQL, dbOpenDynaset, 0, dbPessimistic)
  With rstFilial
    .Edit
    lngNumNFCe = .Fields("NFCe").Value + 1
    .Fields("NFCe").Value = lngNumNFCe
    .Update
    .Close
  End With
  Set rstFilial = Nothing
  
  g_longNextNFCe = lngNumNFCe
  
End Function

'03-04/03/2004 - mpdea
'Otimizado a busca da nova sequ�ncia
'
'31/08/2000 - mpdea
'Atualizado
'Obt�m a �ltima sequ�ncia analisando as tabelas de Entradas e Sa�das
Public Function gnGetNextSequencia(ByVal intFilial As Integer) As Long
  Dim rstCheck As Recordset
  Dim lngSeqMaxP As Long
  Dim lngSeqMaxE As Long
  Dim lngSeqMaxS As Long
  Dim lngSeqMax As Long
  Dim strSQL As String


'  Dim lngStart As Long
'  lngStart = GetTickCount()


'''''''  'Tabela Par�metros da Filial
'''''''  strSQL = "SELECT [�ltima Movimenta��o] AS lngSeqMax FROM [Par�metros Filial] WHERE Filial = " & intFilial
'''''''  Set rstCheck = db.OpenRecordset(strSQL, dbOpenDynaset)
'''''''  With rstCheck
'''''''    .LockEdits = True
'''''''    If .RecordCount > 0 Then
'''''''      Call IsDataType(dtLong, .Fields("lngSeqMax").Value, lngSeqMaxP)
'''''''    End If
'''''''
'''''''    .Edit
'''''''    lngSeqMaxP = lngSeqMaxP + 1
'''''''    .Fields(0).Value = lngSeqMaxP
'''''''    .Update
'''''''
'''''''    .Close
'''''''  End With
'''''''  Set rstCheck = Nothing
'''''''  gnGetNextSequencia = lngSeqMaxP
'''''''  Exit Function
  
  
  'Tabela Par�metros da Filial
  strSQL = "SELECT [�ltima Movimenta��o] AS lngSeqMax FROM [Par�metros Filial] WHERE Filial = " & intFilial
  Set rstCheck = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rstCheck
    If .RecordCount > 0 Then
      Call IsDataType(dtLong, .Fields("lngSeqMax").Value, lngSeqMaxP)
    End If

    .Close
  End With
  Set rstCheck = Nothing

  'Tabela de Entradas
'  strSQL = "SELECT Max(Sequ�ncia) AS lngSeqMax FROM Entradas WHERE Filial = " & intFilial
  strSQL = "SELECT Sequ�ncia AS lngSeqMax FROM Entradas WHERE Filial = " & intFilial & " ORDER BY Sequ�ncia"
  Set rstCheck = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rstCheck
    If .RecordCount > 0 Then
      .MoveLast
      Call IsDataType(dtLong, .Fields("lngSeqMax").Value, lngSeqMaxE)
    End If
    .Close
  End With
  Set rstCheck = Nothing

  'Tabela de Sa�das
  'strSQL = "SELECT Max(Sequ�ncia) AS lngSeqMax FROM Sa�das WHERE Filial = " & intFilial
  strSQL = "SELECT Sequ�ncia AS lngSeqMax FROM Sa�das WHERE Filial = " & intFilial & " ORDER BY Sequ�ncia"
  Set rstCheck = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly) 'dbOpenSnapshot)
  With rstCheck
    If .RecordCount > 0 Then
      .MoveLast
      Call IsDataType(dtLong, .Fields("lngSeqMax").Value, lngSeqMaxS)
    End If
    .Close
  End With
  Set rstCheck = Nothing


  'Verifica a maior sequ�ncia
  lngSeqMax = IIf((lngSeqMaxE > lngSeqMaxS), lngSeqMaxE, lngSeqMaxS) 'Entradas e Sa�das
  lngSeqMax = IIf((lngSeqMaxP > lngSeqMax), lngSeqMaxP, lngSeqMax) 'Par�metros e Final
  gnGetNextSequencia = IIf((lngSeqMax = 0), 1, lngSeqMax + 1)


'  MsgBox "Tempo decorrido: " & (GetTickCount - lngStart) / 1000 & " segundo(s).", vbInformation, "TESTE"

End Function

''Obt�m a �ltima sequ�ncia analisando as tabelas de Entradas e Sa�das
''Atualizado em 31/08/2000 (inclu�do an�lise para a tebela de Par�metros da Filial) - mpdea
'Public Function gnGetNextSequencia(ByVal nFilial As Integer) As Long
'  Dim rsCheck As Recordset
'  Dim nSeqMaxP As Long
'  Dim nSeqMaxE As Long
'  Dim nSeqMaxS As Long
'  Dim nSeqMax As Long
'
'  'Tabela Par�metros da Filial
'  Set rsCheck = db.OpenRecordset("SELECT [�ltima Movimenta��o] AS nSeqMax FROM [Par�metros Filial] WHERE Filial = " & nFilial, dbOpenSnapshot)
'  With rsCheck
'    If .RecordCount > 0 Then
'      nSeqMaxP = IIf(IsNull(!nSeqMax), 1, !nSeqMax)
'    End If
'    .Close
'  End With
'  Set rsCheck = Nothing
'
'  'Tabela de Entradas
'  Set rsCheck = db.OpenRecordset("SELECT Max(Sequ�ncia) AS nSeqMax FROM Entradas WHERE Filial = " & nFilial, dbOpenSnapshot)
'  With rsCheck
'    nSeqMaxE = IIf(IsNull(!nSeqMax), 1, !nSeqMax)
'    .Close
'  End With
'  Set rsCheck = Nothing
'
'  'Tabela de Sa�das
'  Set rsCheck = db.OpenRecordset("SELECT Max(Sequ�ncia) AS nSeqMax FROM Sa�das WHERE Filial = " & nFilial, dbOpenSnapshot)
'  With rsCheck
'    nSeqMaxS = IIf(IsNull(!nSeqMax), 1, !nSeqMax)
'    .Close
'  End With
'  Set rsCheck = Nothing
'
'  'Verifica a maior sequ�ncia
'  nSeqMax = IIf((nSeqMaxE > nSeqMaxS), nSeqMaxE, nSeqMaxS) 'Entradas e Sa�das
'  nSeqMax = IIf((nSeqMaxP > nSeqMax), nSeqMaxP, nSeqMax) 'Par�metros e Final
'  gnGetNextSequencia = IIf((nSeqMax = 0), 1, nSeqMax + 1)
'
'End Function

'Apaga as informa��es de movimenta��o referente a tabela desejada
'Atualizado em 04/09/2000 - mpdea
Public Function EraseTypeMoviment(ByVal nTypeMov As enTableMovimentType, ByVal nFilial As Integer, ByVal nMovimento As Long)
On Error GoTo Erro

  Dim sTable As String
  
  Select Case nTypeMov
    Case tmEntradas
      sTable = "Entradas"
    Case tmEntradasProdutos
      sTable = "[Entradas - Produtos]"
    Case tmSaidas
      sTable = "Sa�das"
    Case tmSaidasProdutos
      sTable = "[Sa�das - Produtos]"
    Case tmSaidasServicos
      sTable = "[Sa�das - Servi�os]"
    Case tmMovimentoCheques
      sTable = "[Movimento - Cheques]"
    Case tmMovimentoParcelas
      sTable = "[Movimento - Parcelas]"
    Case tmMovimentoCartoes
      sTable = "[Movimento - Cartoes]"
  End Select
  
  Call db.Execute("DELETE * FROM " & sTable & " WHERE Filial = " & _
    nFilial & " AND Sequ�ncia = " & nMovimento, dbFailOnError)
  
  Exit Function
Erro:
  MsgBox "Erro na fun��o EraseTypeMoviment " + Err.Number + " " + Err.Description, vbInformation, "Aten��o"
End Function

'Verifica se h� Produtos na movimenta��o solicitada
Public Function gbHasSaidasProdutos(ByVal nFilial As Integer, ByVal nMovimento As Long) As Boolean
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT * FROM [Sa�das - Produtos] WHERE Filial = " & nFilial & " AND Sequ�ncia = " & nMovimento, dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      gbHasSaidasProdutos = True
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'Verifica se h� Servi�os na movimenta��o solicitada
Public Function gbHasSaidasServicos(ByVal nFilial As Integer, ByVal nMovimento As Long) As Boolean
  Dim rsCheck As Recordset
  
  Set rsCheck = db.OpenRecordset("SELECT * FROM [Sa�das - Servi�os] WHERE Filial = " & nFilial & " AND Sequ�ncia = " & nMovimento, dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      gbHasSaidasServicos = True
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'-------------------------------------------------------------------------------------
'Fun��o gstrGetCliForName
'
'Obt�m o nome do cliente/fornecedor na tabela Cli_For
'
'29/04/2002 - mpdea
'-------------------------------------------------------------------------------------

Public Function gstrGetCliForName(ByVal lngCodigo As Long) As String
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Nome FROM Cli_For WHERE C�digo = " & lngCodigo
  Set rs = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rs
    If Not .BOF And Not .EOF Then
      gstrGetCliForName = .Fields("Nome").Value
    End If
    .Close
  End With
  Set rs = Nothing
  
End Function

'---------------------------------------------------------------------------------
'07/05/2002 - mpdea
'
'Obt�m o primeiro caixa dispon�vel no cadastro de caixas
'-------------------------------------------------------------------------------
Public Function gbytFirstCaixa() As Byte
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Min(Caixa) AS FirstCaixa FROM [Caixas em Uso]"
  Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rs
    If Not IsNull(.Fields("FirstCaixa").Value) Then
      gbytFirstCaixa = .Fields("FirstCaixa").Value
    End If
    .Close
  End With
  Set rs = Nothing
  
End Function

'18/07/2002 - mpdea
'Obt�m o nome do computador
Public Function gstrGetComputerName() As String
  Dim strRet As String * 255
  
  If GetComputerName(strRet, 255) <> 0 Then
    gstrGetComputerName = gstrStripNulls(strRet)
  End If
End Function

'02/03/2023 - pablo
'Computador est� no servidor RDP da A3?
Public Function gIsRDP() As Boolean
  gIsRDP = False
  
  If Trim(gstrGetComputerName) = "WIN2003VB" Then gIsRDP = True 'Desenvolvimento
  If Trim(gstrGetComputerName) = "AMAZONA-F74E4RM" Then gIsRDP = True 'Produ��o

End Function

'18/07/2002 - mpdea
'Remove nulos de strings
Private Function gstrStripNulls(ByVal strBuffer As String) As String
  Dim lngPos As Long
  
  lngPos = InStr(strBuffer, vbNullChar)
  If lngPos > 0 Then
    strBuffer = Left$(strBuffer, lngPos - 1)
  End If
  gstrStripNulls = strBuffer
End Function

'08/08/2002 - mpdea
'Obt�m o nr. do pr�ximo or�amento livre da filial
'***Esta fun��o n�o deve ser chamada em transa��es:***
'A atualiza��o do pr�ximo nr. de or�amento deve ser imediata em rede
Public Function glngNextNrOrcamento(ByVal bytFilial As Byte) As Long
  Dim rsCheck As Recordset
  Dim lngNrOrcamento As Long
  Dim bytAttempts As Integer
  
TryAgain:
  
  On Error GoTo ErrHandler
  
  'Tabela Par�metros da Filial
  Set rsCheck = db.OpenRecordset("SELECT NrOrcamento FROM [Par�metros Filial] WHERE Filial = " & bytFilial, dbOpenDynaset)
  With rsCheck
    If .RecordCount > 0 Then
      lngNrOrcamento = CLng("0" & .Fields("NrOrcamento").Value)
      'Verifica��o
      lngNrOrcamento = lngNrOrcamento + 1
      If lngNrOrcamento > CLng(999999) Then
        lngNrOrcamento = 1
      End If
    End If
    .Edit
    .Fields("NrOrcamento").Value = lngNrOrcamento
    .Update
    .Close
  End With
  Set rsCheck = Nothing
  
  glngNextNrOrcamento = lngNrOrcamento
  
  Exit Function
  
ErrHandler:
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
  glngNextNrOrcamento = -1
  
  Select Case Err.Number
    Case 3186, 3197, 3218, 3260 'Registro bloqueado
      'Fecha o recordset e desassocia o objeto
      rsCheck.Close
      Set rsCheck = Nothing
      If bytAttempts < 30 Then
        Call StatusMsg("Aguardando registro bloqueado (" & Err.Number & ")...")
        bytAttempts = bytAttempts + 1
        Call WaitSeconds(1, False) 'Aguarda um segundo
        'Nova tentativa
        GoTo TryAgain
      Else
        If MsgBox("H� no momento registros sendo atualizados no sistema por outra esta��o." & _
          " � necess�rio aguardar por um instante e continuar. Clique em 'OK' para " & _
          "uma nova tentativa.", vbExclamation + vbOKCancel, "Obter N�mero do Or�amento") = vbOK Then
          'Nova tentativa
          bytAttempts = 0
          GoTo TryAgain
        Else
          Exit Function
        End If
      End If
    Case Else
      MsgBox "Erro ao obter o n�mero do or�amento: " & _
        Err.Number & "-" & Err.Description, vbCritical, "Erro"
  End Select
  
End Function

'11/10/2002 - mpdea
'Obt�m o nome da opera��o
Public Function gstrGetNameOper(ByVal enuTipo As enTableMovimentType, ByVal intCodigo As Integer) As String
  Dim rsCheck As Recordset
  Dim strTable As String
  
  Select Case enuTipo
    Case tmEntradas
      strTable = "Opera��es Entrada"
    Case tmSaidas
      strTable = "Opera��es Sa�da"
  End Select
  
  Set rsCheck = db.OpenRecordset("SELECT Nome FROM [" & strTable & "] WHERE C�digo = " & intCodigo, dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gstrGetNameOper = .Fields("Nome").Value & ""
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'17/09/2009 - mpdea
'Obt�m o Modelo de Documento Fiscal da opera��o
Public Function gstrGetModeloDocumentoFiscalOperacao(ByVal enuTipo As enTableMovimentType, ByVal intCodigo As Integer) As String
  Dim rsCheck As Recordset
  Dim strTable As String
  
  Select Case enuTipo
    Case tmEntradas
      strTable = "Opera��es Entrada"
    Case tmSaidas
      strTable = "Opera��es Sa�da"
  End Select
  
  Set rsCheck = db.OpenRecordset("SELECT ModeloDocumentoFiscal FROM [" & strTable & "] WHERE C�digo = " & intCodigo, dbOpenDynaset, dbReadOnly)
  With rsCheck
    If .RecordCount > 0 Then
      gstrGetModeloDocumentoFiscalOperacao = .Fields("ModeloDocumentoFiscal").Value & ""
    End If
    .Close
  End With
  Set rsCheck = Nothing
End Function

'12/11/2002 - mpdea
'Obt�m o Estado cadastrado em Par�metros da Filial
Public Function gstrGetEstadoFilial(ByVal intFilial As Integer) As String
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("SELECT Estado FROM [Par�metros Filial] WHERE Filial = " & intFilial, dbOpenDynaset, dbReadOnly)
  With rs
    If .RecordCount > 0 Then
      gstrGetEstadoFilial = .Fields("Estado").Value & ""
    End If
    .Close
  End With
  Set rs = Nothing
End Function

'31/12/2002 - mpdea
'Fun��o gcGetPrecoProduto traz cota��o agora
'
'Obt�m o pre�o do produto
Public Function gcGetPrecoProduto(ByVal sCodigo As String, ByVal sTabelaPreco As String) As Currency
  Dim rsCheck As Recordset
  Dim curValue As Currency
  Dim bytMoeda As Byte
  Dim dblCotacao As Double
  
  'Obt�m pre�o da tabela
  Set rsCheck = db.OpenRecordset("SELECT Pre�o FROM Pre�os WHERE Produto = '" & sCodigo & "' AND Tabela = '" & sTabelaPreco & "'", dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      Call IsDataType(dtCurrency, .Fields("Pre�o").Value, curValue)
    End If
    .Close
  End With
  Set rsCheck = Nothing
  
  'Verifica moeda e cota��o
  Set rsCheck = db.OpenRecordset("SELECT Moeda FROM Produtos WHERE C�digo = '" & sCodigo & "'", dbOpenSnapshot)
  With rsCheck
    If .RecordCount > 0 Then
      Call IsDataType(dtByte, .Fields("Moeda").Value, bytMoeda)
    End If
    .Close
  End With
  Set rsCheck = Nothing
  
  If bytMoeda <> 1 Then
    
    Set rsCheck = db.OpenRecordset("SELECT Cota��o FROM Cota��es WHERE Moeda = " & bytMoeda & " AND Data <= #" & Format(Data_Atual, "mm/dd/yyyy") & "#", dbOpenSnapshot)
    With rsCheck
      If .RecordCount > 0 Then
        Call IsDataType(dtDouble, .Fields("Cota��o").Value, dblCotacao)
      End If
      .Close
    End With
    Set rsCheck = Nothing
    
    If dblCotacao > 0 Then
      curValue = Format(curValue * dblCotacao, FORMAT_VALUE)
    End If
    
  End If
  
  gcGetPrecoProduto = curValue
  
End Function

'27/12/2002 - mpdea
'Fun��o que trunca o n�mero com a quantidade de decimais desejada
Public Function Truncate(ByVal Number, Optional ByVal NumDigitsAfterDecimal As Long = 0)
  Dim strNumber As String
  Dim lngX As Long
  Dim strSep As String
  
  'Remove express�es cient�ficas (limitando a 50 decimais)
  Number = Format(Number, "0." & String(50, "#"))
  
  'Acha o s�mbolo para separador decimal
  strSep = Mid(0.1, 2, 1)
  
  'Acha a posi��o do separador decimal
  lngX = InStr(Number, strSep)
  
  'Trunca o n�mero decimal
  If lngX > 0 Then
    Number = Left(Number, NumDigitsAfterDecimal + lngX)
  End If
  
  Truncate = Number
    
End Function

'17/04/2003 - mpdea
'Verifica e retorna flag indicando se existe a altera��o personalizada
'
Public Function g_blnCheckChangePersonalized(ByVal strCode As String) As Boolean
  Const strfile As String = "qsperson.cfg"
  Dim intFreeFile As Integer
  Dim strLinha As String
  
  On Error GoTo ErrHandler
  
  'Abre arquivo de configura��o personalizada
  intFreeFile = FreeFile
  If UCase(Dir(gsDefaultPath & strfile)) = UCase(strfile) Then
    Open gsDefaultPath & strfile For Input As #intFreeFile
    Do Until EOF(intFreeFile)
      Line Input #intFreeFile, strLinha
      
      'Analisa linha
      If strLinha = strCode Then
        Close #intFreeFile
        g_blnCheckChangePersonalized = True
        Exit Function
      End If
      
    Loop
    Close #intFreeFile
    
  End If
  
  Exit Function
  
ErrHandler:
  MsgBox "Erro [" & Err.Number & " - " & Err.Description & _
    "] ao ler configura��o personalizada.", vbCritical, "Erro"
  
End Function

Public Function gnGetNextConsignacao(ByVal bytFilial As Byte) As Long
  Dim rstFiliais        As Recordset
  Dim blnInTransaction  As Boolean
  Dim lngNovaSequencia  As Long
  
  On Error GoTo Erro:
  
  ws.BeginTrans
  blnInTransaction = True
  
  Set rstFiliais = db.OpenRecordset(" SELECT Filial, UltimaConsignacao FROM [Par�metros Filial] " & _
                                    " WHERE Filial = " & bytFilial, dbOpenDynaset)
  With rstFiliais
    If Not (.BOF And .EOF) Then
      .Edit
      If IsNull(.Fields("UltimaConsignacao")) Then .Fields("UltimaConsignacao") = 0
      
      lngNovaSequencia = .Fields("UltimaConsignacao") + 1
      .Fields("UltimaConsignacao") = lngNovaSequencia
      .Update
    End If
  End With
  
  ws.CommitTrans
  blnInTransaction = False
  
  gnGetNextConsignacao = lngNovaSequencia
  
  Exit Function
  
Erro:
  If MsgBox("Erro ao gerar a nova sequ�ncia de consigna��o ! " & _
            Err.Number & vbCrLf & vbCrLf & _
            Err.Description, _
            vbCritical + vbRetryCancel, "Quick Store") = vbRetry Then
    Resume
  End If
  
  If blnInTransaction Then
    ws.Rollback
  End If
  
End Function

'05/05/2004 - mpdea
'Fun��o para ler arquivos de configura��o .ini
Public Function gstrReadIniFile(ByVal strFilename As String, ByVal strSection As String, ByVal strKey As String) As String
  Dim strBuffer As String
  Const BUFFER_SIZE As Long = 255
  
  strBuffer = Space$(BUFFER_SIZE)
  If GetPrivateProfileString(strSection, strKey, "", strBuffer, BUFFER_SIZE, strFilename) Then
    gstrReadIniFile = StringFromBuffer(strBuffer)
  End If
End Function

'17/09/2009 - mpdea
'Comentado tratamento de erro, pois a fun��o � utilizada dentro de outras fun��es que j� possuem
'tratamento de erro. Da forma anterior se ocorresse erro seria exibido o mesmo, mas a fun��o origem
'continuaria a ser executada.
Public Function gbNotaManual(ByVal CodOper As Integer, ByVal strMovimentacao As String) As Boolean
  '19/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '
  'O sistema dever� julgar se a nota fiscal ser� criada
  'automaticamente ou manualmente a partir da opera��o escolhida
  Dim rstOperacao As Recordset
  Dim strSQL      As String
  
'  On Error GoTo TratarErro
  
  strSQL = "SELECT EmitirNFManualmente "

  If strMovimentacao = "ENTRADA" Then
    strSQL = strSQL & "FROM [Opera��es Entrada] WHERE C�digo = " & CodOper
  Else
    strSQL = strSQL & "FROM [Opera��es Sa�da] WHERE C�digo = " & CodOper
  End If

  Set rstOperacao = db.OpenRecordset(strSQL, dbOpenDynaset)

  With rstOperacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
      gbNotaManual = .Fields("EmitirNFManualmente").Value
    End If
    .Close
  End With

  Set rstOperacao = Nothing

'  Exit Function
'
'TratarErro:
'  MsgBox "Fun��o gbNotaManual - Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function

'15/09/2005 - mpdea
'Limpa o texto do controle do tipo MaskEdit preservando a m�scara
Public Function ClearMaskEditControl(ByRef mskControl As MaskEdBox)
  Dim strMask As String
  
  With mskControl
    strMask = .Mask
    .Mask = ""
    .Text = ""
    .Mask = strMask
  End With
End Function

'18/01/2006 - mpdea
'Verifica se utiliza a tela de Venda R�pida estilo CheckOut
Public Function g_bln_VendaRapidaCheckOut(ByVal intFilial As Integer) As Boolean
  Dim rs As Recordset
  
  Set rs = db.OpenRecordset("SELECT VR_Tela_CheckOut FROM [Par�metros Filial] WHERE Filial = " & intFilial, dbOpenDynaset, dbReadOnly)
  With rs
    If .RecordCount > 0 Then
      g_bln_VendaRapidaCheckOut = .Fields("VR_Tela_CheckOut").Value
    End If
    .Close
  End With
  Set rs = Nothing
End Function

'31/01/2006 - mpdea
'Obt�m as mensagens para Nota Fiscal a serem utilizadas
'para a movimenta��o informada
'
'Par�metros:
'
'intFilial.....: Filial
'lngSequencia..: N�mero da sequ�ncia de movimenta��o
'strMensagens(): Retorna as mensagens para a movimenta��o informada
Public Sub GetMensagemNotaFiscal(ByVal intFilial As Integer, _
  ByVal lngSequencia As Long, ByRef clcMensagens As Collection)
  
  Dim rstX As Recordset
  Dim strSQL As String
  
  Dim intGrupoFiscalOpSaida As Integer
  Dim strUF As String
  Dim intClasse As Integer
  Dim intSubClasse As Integer
  Dim intGrupoFiscal As Integer
  Dim objMsgNF As MsgNF_Movimentacao
  
  Dim objMensagensNotaFiscal As New clsMensagensNotaFiscal
  Dim objMensagemNotaFiscal As New clsMensagemNotaFiscal
  Dim varRet As Variant
  Dim intX As Integer
  Dim blnRegraOK As Boolean
  
  '07/04/2008 - mpdea
  'Flag para verifica��o de produtos, caso tenha somente servi�os
  Dim blnHasProduto As Boolean
  
  
  On Error GoTo ErrHandler
  
  
  '----------------------------------------------------------------------------
  'Carrega as Mensagens para Nota Fiscal com suas regras
  '
  Call objMensagensNotaFiscal.Load
  '----------------------------------------------------------------------------
  
  
  '----------------------------------------------------------------------------
  'Obt�m os dados da movimenta��o de sa�da para an�lise
  '
  strSQL = "SELECT Opera��o, Cliente "
  strSQL = strSQL & "FROM Sa�das "
  strSQL = strSQL & "WHERE Filial = " & intFilial
  strSQL = strSQL & " AND Sequ�ncia = " & lngSequencia
  
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    If Not (.BOF And .EOF) Then
      'Obt�m
      Call GetDetailsForMsgNFMovimentacao( _
        rstX.Fields("Opera��o").Value, rstX.Fields("Cliente").Value, _
        intGrupoFiscalOpSaida, strUF)
      'Preenche objeto
      With objMsgNF
        .intCodigoOpSaida = CInt(rstX.Fields("Opera��o").Value)
        .intGrupoFiscalOpSaida = intGrupoFiscalOpSaida
        .strUF = strUF
      End With
    End If
    .Close
  End With
  Set rstX = Nothing
  '----------------------------------------------------------------------------
  
  
  '----------------------------------------------------------------------------
  'Obt�m os dados de produtos da movimenta��o para an�lise
  '
  strSQL = "SELECT [C�digo sem Grade] "
  strSQL = strSQL & "FROM [Sa�das - Produtos] "
  strSQL = strSQL & "WHERE Filial = " & intFilial
  strSQL = strSQL & " AND Sequ�ncia = " & lngSequencia
  strSQL = strSQL & " ORDER BY Linha;"
  
  blnHasProduto = False
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    '12/04/2006 - mpdea
    'Inclu�do verifica��o de recordset vazio
    If Not (.BOF And .EOF) Then
      blnHasProduto = True
      .MoveLast
      .MoveFirst
      
      'Redimensiona a cole��o de produtos
      ReDim objMsgNF.clcProdutos(.RecordCount - 1)
      
      Do Until .EOF
        'Obt�m a classe, sub classe e o grupo fiscal do produto
        Call GetDetailsForMsgNFProduct( _
          rstX.Fields("C�digo sem Grade").Value, _
          intClasse, intSubClasse, intGrupoFiscal)
        'Preenche objeto
        With objMsgNF.clcProdutos(.AbsolutePosition)
          .strCodigo = rstX.Fields("C�digo sem Grade").Value
          .intClasse = intClasse
          .intSubClasse = intSubClasse
          .intGrupoFiscal = intGrupoFiscal
        End With
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rstX = Nothing
  '----------------------------------------------------------------------------
  
  '07/04/2008 - mpdea
  'Verifica se h� produtos, caso n�o tenha (ex. nota de servi�os)
  'sai da fun��o
  If Not blnHasProduto Then
    Exit Sub
  End If
  
  '----------------------------------------------------------------------------
  'Analisa as regras das Mensagens
  '
  For Each objMensagemNotaFiscal In objMensagensNotaFiscal
    'Padr�o
    blnRegraOK = False
    
    '1) Filtro para Produtos
    Select Case objMensagemNotaFiscal.TipoFiltroProduto
      Case tfpTodos
        blnRegraOK = True
        
      Case tfpGrupoFiscal
        Call IsDataType(dtInteger, objMensagemNotaFiscal.FiltroProduto, varRet)
        For intX = LBound(objMsgNF.clcProdutos) To UBound(objMsgNF.clcProdutos)
          With objMsgNF.clcProdutos(intX)
            'Grupo Fiscal
            If CInt(varRet) = .intGrupoFiscal Then
              blnRegraOK = True
              Exit For
            End If
          End With
        Next intX
        
      Case tfpClasseSubClasse
        varRet = Split(objMensagemNotaFiscal.FiltroProduto, "|")
        For intX = LBound(objMsgNF.clcProdutos) To UBound(objMsgNF.clcProdutos)
          With objMsgNF.clcProdutos(intX)
            'Classe e Sub Classe
            If varRet(0) <> "" And varRet(1) <> "" Then
              If CInt(varRet(0)) = .intClasse And CInt(varRet(1)) = .intSubClasse Then
                blnRegraOK = True
                Exit For
              End If
            End If
            'Classe
            If varRet(0) <> "" Then
              If CInt(varRet(0)) = .intClasse Then
                blnRegraOK = True
                Exit For
              End If
            End If
            'Sub Classe
            If varRet(1) <> "" Then
              If CInt(varRet(1)) = .intSubClasse Then
                blnRegraOK = True
                Exit For
              End If
            End If
          End With
        Next intX
        
      Case tfpEspecifico
        Call IsDataType(dtString, objMensagemNotaFiscal.FiltroProduto, varRet)
        For intX = LBound(objMsgNF.clcProdutos) To UBound(objMsgNF.clcProdutos)
          With objMsgNF.clcProdutos(intX)
            'C�digo do Produto
            If CStr(varRet) = .strCodigo Then
              blnRegraOK = True
              Exit For
            End If
          End With
        Next intX
        
    End Select
    
    '2) Filtro para Opera��es de Sa�da
    'Somente executa se atendeu o passo anterior
    If blnRegraOK Then
      Select Case objMensagemNotaFiscal.TipoFiltroOpSaida
        Case tfoTodas
          blnRegraOK = True
          
        Case tfoGrupoFiscal
          Call IsDataType(dtInteger, objMensagemNotaFiscal.FiltroOpSaida, varRet)
          blnRegraOK = (objMsgNF.intGrupoFiscalOpSaida = varRet)
          
        Case tfoEspecifica
          Call IsDataType(dtInteger, objMensagemNotaFiscal.FiltroOpSaida, varRet)
          blnRegraOK = (objMsgNF.intCodigoOpSaida = varRet)
          
      End Select
    End If
    
    '3) Filtro para Estado (UF) do Cliente
    'Somente executa se atendeu o passo anterior
    If blnRegraOK Then
      Select Case objMensagemNotaFiscal.TipoFiltroUF
        Case tfuTodos
          blnRegraOK = True
          
        Case tfuEspecifico
          Call IsDataType(dtString, objMensagemNotaFiscal.FiltroUF, varRet)
          blnRegraOK = (objMsgNF.strUF = varRet)
          
      End Select
    End If
    
    'Adiciona Mensagem se a atende a regra
    If blnRegraOK Then
      clcMensagens.Add objMensagemNotaFiscal.Mensagem
    End If
  Next objMensagemNotaFiscal
  '----------------------------------------------------------------------------
  
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstX Is Nothing Then
    rstX.Close
    Set rstX = Nothing
  End If
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

'31/01/2006 - mpdea
'Obt�m dados do cadastro de produtos para a an�lise das mensagens para nota fiscal
Private Sub GetDetailsForMsgNFProduct(ByVal strCodigoProduto As String, _
  ByRef intClasse As Integer, ByRef intSubClasse As Integer, _
  ByRef intGrupoFiscal As Integer)
  
  Dim rstX As Recordset
  Dim strSQL As String
  
  
  On Error GoTo ErrHandler
  
  
  strSQL = "SELECT Classe, [Sub Classe], GrupoFiscal "
  strSQL = strSQL & "FROM Produtos "
  strSQL = strSQL & "WHERE C�digo = '" & strCodigoProduto & "';"
  
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    If Not (.BOF And .EOF) Then
      Call IsDataType(dtInteger, .Fields("Classe").Value, intClasse)
      Call IsDataType(dtInteger, .Fields("Sub Classe").Value, intSubClasse)
      Call IsDataType(dtInteger, .Fields("GrupoFiscal").Value, intGrupoFiscal)
    End If
    .Close
  End With
  Set rstX = Nothing
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstX Is Nothing Then
    rstX.Close
    Set rstX = Nothing
  End If
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

'31/01/2006 - mpdea
'Obt�m dados da movimenta��o de sa�da para a an�lise das mensagens para nota fiscal
Private Sub GetDetailsForMsgNFMovimentacao(ByVal intCodigoOpSaida As Integer, _
  ByVal lngCodigoCliente As Long, ByRef intOpSaidaGrupoFiscal As Integer, _
  ByRef strUF As String)
  
  Dim rstX As Recordset
  Dim strSQL As String
  
  
  On Error GoTo ErrHandler
  
  
  'Opera��o de Sa�da
  strSQL = "SELECT GrupoFiscal "
  strSQL = strSQL & "FROM [Opera��es Sa�da] "
  strSQL = strSQL & "WHERE C�digo = " & intCodigoOpSaida
  
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    If Not (.BOF And .EOF) Then
      Call IsDataType(dtInteger, .Fields("GrupoFiscal").Value, intOpSaidaGrupoFiscal)
    End If
    .Close
  End With
  Set rstX = Nothing
  
  'Cliente
  strSQL = "SELECT Estado "
  strSQL = strSQL & "FROM Cli_For "
  strSQL = strSQL & "WHERE C�digo = " & lngCodigoCliente
  
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    If Not (.BOF And .EOF) Then
      Call IsDataType(dtString, .Fields("Estado").Value, strUF)
    End If
    .Close
  End With
  Set rstX = Nothing
  
  Exit Sub
  
ErrHandler:
  'Fecha tabela
  If Not rstX Is Nothing Then
    rstX.Close
    Set rstX = Nothing
  End If
  'Repassa erro
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

'-------------------------------------------------------------------------------------
'Obt�m o nome do cliente/fornecedor e seu email na tabela Cli_For
'
'30/01/2009 - mpdea
'-------------------------------------------------------------------------------------
Public Sub GetEmailDetailsCliFor(ByVal lngCodigo As Long, ByRef strNome As String, ByRef strEmail As String)
  Dim rs As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT Nome, email FROM Cli_For WHERE C�digo = " & lngCodigo
  Set rs = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rs
    If Not (.BOF And .EOF) Then
      strNome = .Fields("Nome").Value & ""
      strEmail = .Fields("email").Value & ""
    End If
    .Close
  End With
  Set rs = Nothing
  
End Sub

