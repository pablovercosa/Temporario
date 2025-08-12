Attribute VB_Name = "modNfe"
Option Explicit

'17/09/2009 - mpdea
'Habilitar uso de Nota Fiscal Eletr�nica
Public gblnNFe As Boolean

'20/08/2009 - mpdea
Public Sub AlteraDBNFe()
  Dim intStep As Integer
  Dim strErro As String
  
  
  On Error GoTo ErrHandler
  
  '1. Tabela para codifica��o de UF
  intStep = intStep + 1
  If Not gbGetTable("TerritorioUf") Then
    CreateTableTerritorioUf
  End If
  
  '2. Tabela para codifica��o de Munic�pios
  intStep = intStep + 1
  If Not gbGetTable("TerritorioMunicipio") Then
    CreateTableTerritorioMunicipio
  End If
  
  '3. Tabela para codifica��o de Pa�ses
  intStep = intStep + 1
  If Not gbGetTable("TerritorioPais") Then
    CreateTableTerritorioPais
  End If
  
  '25/08/2009 - mpdea
  '4. Inscri��o Municipal para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "InscricaoMunicipal") Then
    gbCreateField "Par�metros Filial", "InscricaoMunicipal", dbText, 20, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET InscricaoMunicipal = ''", dbFailOnError
  End If
  
  '25/08/2009 - mpdea
  '5. Inscri��o Municipal para clientes/fornecedores
  intStep = intStep + 1
  If Not gbGetField("Cli_For", "InscricaoMunicipal") Then
    gbCreateField "Cli_For", "InscricaoMunicipal", dbText, 20, True, False, False
    'Valor padr�o
    db.Execute "UPDATE Cli_For SET InscricaoMunicipal = ''", dbFailOnError
  End If
  
  '25/08/2009 - mpdea
  '6. CNAE para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "CNAE") Then
    gbCreateField "Par�metros Filial", "CNAE", dbText, 10, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET CNAE = ''", dbFailOnError
  End If
  
  '25/08/2009 - mpdea
  '7. CNAE para clientes/fornecedores
  intStep = intStep + 1
  If Not gbGetField("Cli_For", "CNAE") Then
    gbCreateField "Cli_For", "CNAE", dbText, 10, True, False, False
    'Valor padr�o
    db.Execute "UPDATE Cli_For SET CNAE = ''", dbFailOnError
  End If
    
  '25/08/2009 - mpdea
  '8. N�mero (Endere�o) para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "EnderecoNumero") Then
    gbCreateField "Par�metros Filial", "EnderecoNumero", dbText, 10, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET EnderecoNumero = ''", dbFailOnError
  End If
  
  '25/08/2009 - mpdea
  '9. Complemento (Endere�o) para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "EnderecoComplemento") Then
    gbCreateField "Par�metros Filial", "EnderecoComplemento", dbText, 30, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET EnderecoComplemento = ''", dbFailOnError
  End If
  
  '25/08/2009 - mpdea
  '10. CEP para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "CEP") Then
    gbCreateField "Par�metros Filial", "CEP", dbText, 8, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET CEP = ''", dbFailOnError
  End If
  
  '25/08/2009 - mpdea
  '11. Pa�s para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "Pais") Then
    gbCreateField "Par�metros Filial", "Pais", dbText, 60, True, False, False
    'Valor padr�o (Brasil)
    SetPaisBrasil "Par�metros Filial"
  End If
  
  '25/08/2009 - mpdea
  '12. Pa�s para clientes/fornecedores
  intStep = intStep + 1
  If Not gbGetField("Cli_For", "Pais") Then
    gbCreateField "Cli_For", "Pais", dbText, 60, True, False, False
    'Valor padr�o (Brasil)
    SetPaisBrasil "Cli_For"
  End If
  
  '25/08/2009 - mpdea
  '13. SUFRAMA para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "InscricaoSuframa") Then
    gbCreateField "Par�metros Filial", "InscricaoSuframa", dbText, 9, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET InscricaoSuframa = ''", dbFailOnError
  End If
  
  '25/08/2009 - mpdea
  '14. SUFRAMA para clientes/fornecedores
  intStep = intStep + 1
  If Not gbGetField("Cli_For", "InscricaoSuframa") Then
    gbCreateField "Cli_For", "InscricaoSuframa", dbText, 9, True, False, False
    'Valor padr�o
    db.Execute "UPDATE Cli_For SET InscricaoSuframa = ''", dbFailOnError
  End If
  
  '01/09/2009 - mpdea
  '15. Identifica��o do Ambiente NFe para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "AmbienteNfe") Then
    gbCreateField "Par�metros Filial", "AmbienteNFe", dbByte, , , False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET AmbienteNFe = 2", dbFailOnError
  End If
  
  '01/09/2009 - mpdea
  '16. Formato de Impress�o do DANFE NFe para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "FormatoImpressaoDanfeNfe") Then
    gbCreateField "Par�metros Filial", "FormatoImpressaoDanfeNfe", dbByte, , , False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET FormatoImpressaoDanfeNfe = 1", dbFailOnError
  End If
  
  '01/09/2009 - mpdea
  '17. Modalidade de determina��o da Base de C�lculo do ICMS para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "ModDetBaseCalculoIcms") Then
    gbCreateField "Par�metros Filial", "ModDetBaseCalculoIcms", dbByte, , , False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET ModDetBaseCalculoIcms = 0", dbFailOnError
  End If
  
  '01/09/2009 - mpdea
  '18. Modalidade de determina��o da Base de C�lculo do ICMS ST para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "ModDetBaseCalculoIcmsSt") Then
    gbCreateField "Par�metros Filial", "ModDetBaseCalculoIcmsSt", dbByte, , , False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET ModDetBaseCalculoIcmsSt = 0", dbFailOnError
  End If
  
  '01/09/2009 - mpdea
  '19. Pasta de envio NFe para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "PastaEnvioNfe") Then
    gbCreateField "Par�metros Filial", "PastaEnvioNfe", dbText, 255, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET PastaEnvioNfe = ''", dbFailOnError
  End If
  
  '01/09/2009 - mpdea
  '20. Pasta de retorno NFe para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "PastaRetornoNfe") Then
    gbCreateField "Par�metros Filial", "PastaRetornoNfe", dbText, 255, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET PastaRetornoNfe = ''", dbFailOnError
  End If
  
  '14/09/2009 - mpdea
  '21. Adiciona permiss�o para Envio e Retorno de Nota Fiscal Eletr�nica
  intStep = intStep + 1
  Call AddUserPermission("NOTA FISCAL ELETR�NICA", "Nota Fiscal Eletr�nica - Envio e Retorno", 182, ID_ITEM_MOVIMENTO_NOTA_FISCAL_ELETRONICA)
  
  '17/09/2009 - mpdea
  '22. Habilitar NFe para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "HabilitarNotaFiscalEletronica") Then
    gbCreateField "Par�metros Filial", "HabilitarNotaFiscalEletronica", dbBoolean, , False, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET HabilitarNotaFiscalEletronica = False", dbFailOnError
  End If
  
  '17/09/2009 - mpdea
  '23. Modelo de documento fiscal para opera��es de entrada
  intStep = intStep + 1
  If Not gbGetField("Opera��es Entrada", "ModeloDocumentoFiscal") Then
    gbCreateField "Opera��es Entrada", "ModeloDocumentoFiscal", dbText, 2, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Opera��es Entrada] SET ModeloDocumentoFiscal = '1'", dbFailOnError
  End If
  
  '17/09/2009 - mpdea
  '24. Modelo de documento fiscal para opera��es de sa�da
  intStep = intStep + 1
  If Not gbGetField("Opera��es Sa�da", "ModeloDocumentoFiscal") Then
    gbCreateField "Opera��es Sa�da", "ModeloDocumentoFiscal", dbText, 2, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Opera��es Sa�da] SET ModeloDocumentoFiscal = '1'", dbFailOnError
  End If
  
  '17/09/2009 - mpdea
  '25. Modelo de documento fiscal para entradas
  intStep = intStep + 1
  If Not gbGetField("Entradas", "ModeloDocumentoFiscal") Then
    gbCreateField "Entradas", "ModeloDocumentoFiscal", dbText, 2, True, False, False
    'Valor padr�o
    db.Execute "UPDATE Entradas SET ModeloDocumentoFiscal = '1'", dbFailOnError
  End If
  
  '17/09/2009 - mpdea
  '26. Modelo de documento fiscal para sa�das
  intStep = intStep + 1
  If Not gbGetField("Sa�das", "ModeloDocumentoFiscal") Then
    gbCreateField "Sa�das", "ModeloDocumentoFiscal", dbText, 2, True, False, False
    'Valor padr�o
    db.Execute "UPDATE Sa�das SET ModeloDocumentoFiscal = '1'", dbFailOnError
  End If
  
  '18/09/2009 - mpdea
  '27. Tabela para informa��es sobre as NFe enviadas
  intStep = intStep + 1
  If Not gbGetTable("NFe") Then
    CreateTableNFe
  End If
  
  '18/09/2009 - mpdea
  '28. Tabela para detalhes de retorno das NFe enviadas
  intStep = intStep + 1
  If Not gbGetTable("NFeRetorno") Then
    CreateTableNFeRetorno
  End If
  
  '30/09/2009 - Andrea
  '29. Numero da NFe para empresa/filial
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "UltimaNFe") Then
    gbCreateField "Par�metros Filial", "UltimaNFe", dbLong
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET UltimaNFe = 0", dbFailOnError
  End If
  
  '25/01/2010 - mpdea
  '30. Altera��o do tamanho do campo CEP para empresa/filial
  intStep = intStep + 1
  Call gbAlteraTamanhoCampo("Par�metros Filial", "CEP", dbText, 10, False)
  
  '17/11/2010 - Andrea
  '31. Vers�o do layout de envio da NFe
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "VersaoLayoutEnvio") Then
    gbCreateField "Par�metros Filial", "VersaoLayoutEnvio", dbText, 6, False, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET VersaoLayoutEnvio = '1.10'", dbFailOnError
  End If

  '24/11/2010 - Andrea
  '32. C�digo do Regime Tribut�rio da empresa/filial para NFe
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "CodigoRegimeTributario") Then
    gbCreateField "Par�metros Filial", "CodigoRegimeTributario", dbByte, , , False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET CodigoRegimeTributario = 1", dbFailOnError
  End If

  '11/03/2011 - Andrea
  '33. Percentual Simples Nacional da empresa/filial para NFe
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "PercentualSimplesNacional") Then
    gbCreateField "Par�metros Filial", "PercentualSimplesNacional", dbDouble, , , False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET PercentualSimplesNacional = 0", dbFailOnError
  End If
  
  '11/03/2011 - Andrea
  '34. Inclus�o na Tabela Produtos
  '    Inclu�do campo CSO (C�digo da Situacao da Operacao - Simples Nacional
  intStep = intStep + 1
  If Not gbGetField("Produtos", "CSO") Then
    gbCreateField "Produtos", "CSO", dbText, 3, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Produtos] SET CSO = 000", dbFailOnError
  End If

  '30/03/2011 - Andrea
  '35. Percentual de Redu��o da Base de C�lculo do Simples Nacional da empresa/filial para NFe
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "PercentualReducaoBCSimplesNacional") Then
    gbCreateField "Par�metros Filial", "PercentualReducaoBCSimplesNacional", dbDouble, , , False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET PercentualReducaoBCSimplesNacional = 0", dbFailOnError
  End If
  
  '36. C�digo de Situa��o da Opera��o - Simples Nacional - Opera��es de Sa�da
  intStep = intStep + 1
  If Not gbGetField("Opera��es Sa�da", "CSO") Then
    gbCreateField "Opera��es Sa�da", "CSO", dbText, 3, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Opera��es Sa�da] SET CSO = '000'", dbFailOnError
  End If
  
  '37. C�digo de Situa��o da Opera��o - Simples Nacional - Opera��es de Entrada
  intStep = intStep + 1
  If Not gbGetField("Opera��es Entrada", "CSO") Then
    gbCreateField "Opera��es Entrada", "CSO", dbText, 3, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [Opera��es Entrada] SET CSO = '000'", dbFailOnError
  End If
  
  '38. C�digo de Situa��o da Opera��o - Simples Nacional - ProdutosCFOP
  intStep = intStep + 1
  If Not gbGetField("ProdutoCFOP", "CSO") Then
    gbCreateField "ProdutoCFOP", "CSO", dbText, 3, True, False, False
    'Valor padr�o
    db.Execute "UPDATE [ProdutoCFOP] SET CSO = '000'", dbFailOnError
  End If
  
  '03/05/2011 - Andrea
  '39. Inclus�o na Tabela Produtos
  '    Inclu�do campo IPI_Reduzido
  intStep = intStep + 1
  If Not gbGetField("Produtos", "IPI_Reduzido") Then
    gbCreateField "Produtos", "IPI_Reduzido", dbBoolean, , False, False, False
    'Valor padr�o
    db.Execute "UPDATE [Produtos] SET IPI_Reduzido = False", dbFailOnError
  End If
  
  '16/11/2011 - Andrea
  '40. Padr�o do arquivo de integra��o (TXT ou XML)
  intStep = intStep + 1
  If Not gbGetField("Par�metros Filial", "PadraoArquivoIntegracao") Then
    gbCreateField "Par�metros Filial", "PadraoArquivoIntegracao", dbText, 6, False, False, False
    'Valor padr�o
    db.Execute "UPDATE [Par�metros Filial] SET PadraoArquivoIntegracao = 'TXT'", dbFailOnError
  End If

  


  Exit Sub
  
ErrHandler:
  strErro = "Erro ao atualizar informa��es para NF-e, fase " & intStep & ". "
  Err.Raise Err.Number, Err.Source, strErro & Err.Description, Err.HelpFile, Err.HelpContext

End Sub

'17/08/2009 - mpdea
'Tabela para codifica��o de UF
Private Sub CreateTableTerritorioUf()
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  Set td = db.CreateTableDef("TerritorioUf")
  
  Set fd = td.CreateField("Nome", dbText, 64)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Sigla", dbText, 2)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoIbge", dbByte)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Sigla")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing
  
  'Insere os itens
  InsertItensTerritorio "TerritorioUf"
    
End Sub

'17/08/2009 - mpdea
'Tabela para codifica��o de Munic�pios
Private Sub CreateTableTerritorioMunicipio()
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  Set td = db.CreateTableDef("TerritorioMunicipio")
  
  Set fd = td.CreateField("Uf", dbText, 2)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Nome", dbText, 64)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Nome2", dbText, 64)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoIbge", dbLong)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Uf")
  iX.Fields.Append iX.CreateField("Nome")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing

  'Insere os itens
  InsertItensTerritorio "TerritorioMunicipio"

End Sub

'17/08/2009 - mpdea
'Tabela para codifica��o de Pa�ses
Private Sub CreateTableTerritorioPais()
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  Set td = db.CreateTableDef("TerritorioPais")
  
  Set fd = td.CreateField("Nome", dbText, 64)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("CodigoBacen", dbInteger)
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Nome")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing

  'Insere os itens
  InsertItensTerritorio "TerritorioPais"

End Sub

'17/08/2009 - mpdea
'Insere os registros nas tabelas de territ�rios
Private Sub InsertItensTerritorio(ByVal strItem As String)
  Dim intFeeFile As Integer
  Dim strLinha As String
  Dim strCampos() As String
  Dim strSQL As String
  
  On Error GoTo ErrHandler
  
  intFeeFile = FreeFile
  Open gsDefaultPath & "Resources\" & strItem & ".txt" For Input As #intFeeFile
  Do Until EOF(intFeeFile)
    'L� a linha
    Line Input #intFeeFile, strLinha
    If strLinha <> "" Then
      'Separa em campos
      strCampos = Split(strLinha, vbTab)
      'Tipo de item
      Select Case strItem
        Case "TerritorioUf"
          strSQL = "INSERT INTO " & strItem & " VALUES ('" & strCampos(0) & "', '" & strCampos(1) & "', " & strCampos(2) & ")"
        Case "TerritorioMunicipio"
          strSQL = "INSERT INTO " & strItem & " VALUES ('" & strCampos(0) & "', '" & strCampos(1) & "', '" & RetiraAcento(strCampos(1)) & "', " & strCampos(2) & ")"
        Case "TerritorioPais"
          strSQL = "INSERT INTO TerritorioPais VALUES ('" & strCampos(0) & "', " & strCampos(1) & ")"
      End Select
      'Insere o registro
      db.Execute strSQL, dbFailOnError
    End If
  Loop

  Close intFeeFile
  
  Exit Sub
ErrHandler:
  Close
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Sub

'17/08/2009 - mpdea
'Obt�m o c�digo IBGE para a UF informada
Public Function GetTerritorioUfCodigoIbge(ByVal strUF As String) As Byte
  Dim rstX As Recordset
  Dim strSQL As String
  Dim bytReturn As Byte
  
  strSQL = "SELECT CodigoIbge FROM TerritorioUf WHERE Sigla = '" & strUF & "'"
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    If Not (.BOF And .EOF) Then
      bytReturn = .Fields("CodigoIbge").Value
    End If
    .Close
  End With
  Set rstX = Nothing
  
  GetTerritorioUfCodigoIbge = bytReturn
End Function

'17/08/2009 - mpdea
'Obt�m o c�digo IBGE para o Munic�pio da UF informada
Public Function GetTerritorioMunicipioCodigoIbge(ByVal strUF As String, ByVal strMunicipio As String) As Long
  Dim rstX As Recordset
  Dim strSQL As String
  Dim lngReturn As Long
  
  ' Tratar APOSTROFE
  Dim iIndex As Integer
  iIndex = -1
  iIndex = InStr(1, strMunicipio, "'")
  If iIndex > 0 Then
    strMunicipio = Replace(strMunicipio, "'", "''")
  End If
  '
  
  strSQL = "SELECT CodigoIbge FROM TerritorioMunicipio WHERE Uf = '" & strUF & "' AND Nome2 = '" & RetiraAcento(strMunicipio) & "'"
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    If Not (.BOF And .EOF) Then
      lngReturn = .Fields("CodigoIbge").Value
    End If
    .Close
  End With
  Set rstX = Nothing
  
  GetTerritorioMunicipioCodigoIbge = lngReturn
End Function

'17/08/2009 - mpdea
'Obt�m o c�digo BACEN para o Pa�s informado
Public Function GetTerritorioPaisCodigoBacen(ByVal strPais As String) As Integer
  Dim rstX As Recordset
  Dim strSQL As String
  Dim intReturn As Integer
  
  strSQL = "SELECT CodigoBacen FROM TerritorioPais WHERE Nome = '" & strPais & "'"
  Set rstX = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstX
    If Not (.BOF And .EOF) Then
      intReturn = .Fields("CodigoBacen").Value
    End If
    .Close
  End With
  Set rstX = Nothing
  
  GetTerritorioPaisCodigoBacen = intReturn
End Function

'17/08/2009 - mpdea
'Retira acentos do texto informado
Public Function RetiraAcento(ByVal strTexto As String) As String
  Const ESPECIAL As String = "������������������������������������������������"
  Const NORMAL   As String = "cCaaaaaAAAAAeeeeEEEEiiiiIIIIoooooOOOOOuuuuUUUUnN"
  Dim intX As Integer
  
  For intX = 1 To Len(ESPECIAL)
    strTexto = Replace(strTexto, Mid(ESPECIAL, intX, 1), Mid(NORMAL, intX, 1))
  Next

  RetiraAcento = strTexto
End Function

Public Function RetiraAcento2(ByVal strTexto As String) As String
  Const ESPECIAL As String = "������������������������������������������������"
  Const NORMAL   As String = "cCaaaaaAAAAAeeeeEEEEiiiiIIIIoooooOOOOOuuuuUUUUnN"
  Dim intX As Integer
  
  For intX = 1 To Len(ESPECIAL)
    strTexto = Replace(strTexto, Mid(ESPECIAL, intX, 1), Mid(NORMAL, intX, 1))
  Next
  
  strTexto = Replace(strTexto, "'", "")
  
  RetiraAcento2 = strTexto
End Function

'26/08/2009 - mpdea
'Atualiza a tabela com o nome do Pa�s igual a Brasil caso perten�a a algum estado brasileiro
Private Sub SetPaisBrasil(ByVal strTabela As String)
  Dim strSQL As String
  
  strSQL = "UPDATE [" & strTabela & "] "
  strSQL = strSQL & "SET Pais = 'Brasil' "
  strSQL = strSQL & "WHERE Estado IN ('RO', 'AC', 'AM', 'RR', 'PA', 'AP', 'TO', "
  strSQL = strSQL & "'MA', 'PI', 'CE', 'RN', 'PB', 'PE', 'AL', 'SE', 'BA', 'MG', "
  strSQL = strSQL & "'ES', 'RJ', 'SP', 'PR', 'SC', 'RS', 'MS', 'MT', 'GO', 'DF')"
  
  db.Execute strSQL, dbFailOnError
End Sub

'18/09/2009 - mpdea
'Tabela para informa��es sobre as NFe enviadas
Private Sub CreateTableNFe()
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  Set td = db.CreateTableDef("NFe")
  
  Set fd = td.CreateField("Filial", dbInteger)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Sequencia", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("TipoMovimento", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("DataHoraEnvio", dbDate)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Status", dbInteger)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Ambiente", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("FormaEmissao", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Numero", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Serie", dbText, 3)
  fd.AllowZeroLength = False
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Modelo", dbText, 2)
  fd.AllowZeroLength = False
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ChaveAcesso", dbText, 44)
  fd.AllowZeroLength = False
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("ProtocoloAutorizacao", dbText, 15)
  fd.AllowZeroLength = True
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("DataHoraAutorizacao", dbDate)
  fd.Required = False
  td.Fields.Append fd
  Set fd = td.CreateField("ProtocoloCancelamento", dbText, 15)
  fd.AllowZeroLength = True
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("DataHoraCancelamento", dbDate)
  fd.Required = False
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Filial")
  iX.Fields.Append iX.CreateField("Sequencia")
  iX.Fields.Append iX.CreateField("TipoMovimento")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  Set iX = td.CreateIndex("ChaveAcesso")
  iX.Fields.Append iX.CreateField("ChaveAcesso")
  iX.Primary = False
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing

End Sub

'18/09/2009 - mpdea
'Tabela para detalhes de retorno das NFe enviadas
Private Sub CreateTableNFeRetorno()
  Dim td As TableDef
  Dim fd As Field
  Dim iX As Index
  
  Set td = db.CreateTableDef("NFeRetorno")
  
  Set fd = td.CreateField("Filial", dbInteger)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Sequencia", dbLong)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("TipoMovimento", dbByte)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("DataHora", dbDate)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Protocolo", dbText, 15)
  fd.AllowZeroLength = True
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("DigestValue", dbText, 28)
  fd.AllowZeroLength = True
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("Status", dbInteger)
  fd.Required = True
  td.Fields.Append fd
  Set fd = td.CreateField("StatusDescricao", dbText, 255)
  fd.AllowZeroLength = False
  fd.Required = True
  td.Fields.Append fd
    
  Set iX = td.CreateIndex("PrimaryKey")
  iX.Fields.Append iX.CreateField("Filial")
  iX.Fields.Append iX.CreateField("Sequencia")
  iX.Fields.Append iX.CreateField("TipoMovimento")
  iX.Primary = True
  iX.Unique = True
  td.Indexes.Append iX
    
  db.TableDefs.Append td
  
  Set td = Nothing

End Sub

