Attribute VB_Name = "modXML_Regras"
Option Explicit

Public BancoPDV As Database

'Variaveis globais da tem Gerenciamento NFe Aba Preta
Public nfeDevolucao_impostoDevol As Boolean
Public nfeInfAdProd As Boolean

Public nfe_xPed_nItemPed As Boolean
Public gSequenciaSaidas As Long
Public gProdutosXPed_nItemPed(100, 6) As String
Public gProdutosXPed_nItemPedContador As Integer
Public gtotal_vBCST_Parametrizacao As String
Public gtotal_vICMSST_Parametrizacao As String
'Fim aba preta


Public gRetOID As Long


Public gCdGuidResgate As String
Public gSaldoCdGuidResgate As Double
Public gCdClienteCdGuidResgate As Long
Public gNmClienteCdGuidResgate As String
Public gClienteEntregouResgatePontos As Boolean
Public gSaldoCdGuidResgate_clicou_ok_telaDesconto As Boolean

Public sMENSAGEM_LOG_TESTE_GERAL As String
Public gSetPrinterName_jaChamou_REL As Integer    '0-Não; 1-JáChamou
Public gObjReport_Global_REL As Variant
Public gSetPrinterName_jaChamou_NOTA As Integer    '0-Não; 1-JáChamou
Public gObjReport_Global_NOTA As Variant
Public gSetPrinterName_jaChamou_TICKET As Integer    '0-Não; 1-JáChamou
Public gObjReport_Global_TICKET As Variant
Public gSetPrinterName_jaChamou_CHEQUE As Integer    '0-Não; 1-JáChamou
Public gObjReport_Global_CHEQUE As Variant
Public gSetPrinterName_jaChamou_BOLETO As Integer    '0-Não; 1-JáChamou
Public gObjReport_Global_BOLETO As Variant
Public gSetPrinterName_jaChamou_CARNE As Integer    '0-Não; 1-JáChamou
Public gObjReport_Global_CARNE As Variant 'CrystalReport


Public CodigoProdutoCestaPesq As String
Public NomeProdutoCestaPesq As String

Public gCNPJ_CPFControleDeLicencaWebApi As String
Public gNomeEmpresaFilial As String
Public gAcessoContingencia_ControleDeLicencaWebApi As String
Public gParticipaProgramaFidelidade As Integer      '1-SIM PARTICIPA;   0-NÃO PARTICIPA Empresa/filial;
Public gIndicadorProgramaFidelidadeCNPJPrincipal As Integer  '1-CNPJ PRINCIPAL;  2-CNPJ VINCULADO;   3-NADA
Public gCNPJProgramaFidelidadeCNPJPrincipal As String


Public gAbreModuloXML As Integer                    '1-ABRE; 0-NÃO ABRE;
Public gTrataControleDeLicencaWebApi As Integer     '1-Trata; 0-Não Trata

Public gESTRATEGICO_Relatorios As Integer           '1-Tem acesso; 0-Não tem

Public gCodClienteA3CadResult As String
Public gINTEGRACAO_APP_ERR_QUICK As Boolean

Public origemTelaSaidasParaTelaNFe As String
Public bHorarioDeVerao As Boolean

Public sVersaoDoSistema As String

Public bSoapClient_MSSoapInit As Boolean
Public soapclient As SoapClient30

Public bSoapClient_MSSoapInit_NFCe As Boolean
Public soapclient_NFCe As SoapClient30

Public sEnderecoValidaLicencaQuickStoreWEBAPI As String
Public sSoapClient_MSSoapInit As String
Public sSoapClient_ConnectorProperty_EndPointURL As String
Public sCaminhoDanfe_Benefix As String

Public sSoapClient_MSSoapInit_NFCe As String
Public sSoapClient_ConnectorProperty_EndPointURL_NFCe As String

Public gStringRetorno As String

' gArrayNrSerieNF
' até 10 CNPJs
' ex:
'     NrCnpj1 , SerieNFe1, SerieNFCe1
'     NrCnpj2 , SerieNFe2, SerieNFCe2
'     ...
Public gArrayNrSerieNF(10, 3) As String
Public gNrSerieNF As Integer

Public gManterAtivo As String

Public gTipoSituacaoTributariaPIS As Integer

Public gbMostrarTelaPesquisaProdutoTipoFoto As Boolean
Public gbUsuarioAcessoApenasTelaVendaRapida As Boolean      'Tela de venda rapida com funcionalidades mínimas (no caso somente os botões basicos)
                                                            'Importante: Isto não impede o usuario de acessar as outras telas.

Public gPastaRetornoNfe As String

Public gTabelaPrecoAcatadaTelaPesquisaProduto As String

Public gSenhaUsuarioLogado As String

Public gOrigemTelaSaidasChamadorDaTelaAcharVendaHoje As Boolean

Public giQuick_viaRDP As Integer
Public giQuick_viaRDP_ticket As Integer
Public gsEstadoOrigemEmpresaLogado As String
Public gblnSimplesNacional As Boolean
Public gbPingAtivado As Boolean

Public sEnviaDataRetroativa As String
Public bolEnviaDataRetroativa As Boolean

Public sRetornoEnvioNFCe As String

Public Declare Sub openport Lib "c:\windows\system\tsclib.dll" (ByVal PrinterName As String)
Public Declare Sub closeport Lib "c:\windows\system\tsclib.dll" ()
Public Declare Sub sendcommand Lib "c:\windows\system\tsclib.dll" (ByVal command As String)
Public Declare Sub setup Lib "c:\windows\system\tsclib.dll" (ByVal LabelWidth As String, ByVal LabelHeight As String, _
ByVal Speed As String, _
ByVal Density As String, _
ByVal Sensor As String, _
ByVal Vertical As String, _
ByVal Offset As String)
Public Declare Sub downloadpcx Lib "c:\windows\system\tsclib.dll" (ByVal Filename As String, ByVal ImageName As String)
Public Declare Sub barcode Lib "c:\windows\system\tsclib.dll" (ByVal X As String, ByVal y As String, ByVal CodeType As String, _
ByVal Height As String, _
ByVal Readable As String, _
ByVal rotation As String, _
ByVal Narrow As String, _
ByVal Wide As String, _
ByVal Code As String)
Public Declare Sub printerfont Lib "c:\windows\system\tsclib.dll" (ByVal X As String, ByVal y As String, ByVal FontName As String, _
ByVal rotation As String, _
ByVal Xmul As String, _
ByVal Ymul As String, _
ByVal Content As String)
Public Declare Sub clearbuffer Lib "c:\windows\system\tsclib.dll" ()
Public Declare Sub printlabel Lib "c:\windows\system\tsclib.dll" (ByVal NumberOfSet As String, ByVal NumberOfCopy As String)
Public Declare Sub formfeed Lib "c:\windows\system\tsclib.dll" ()
Public Declare Sub nobackfeed Lib "c:\windows\system\tsclib.dll" ()
Public Declare Sub windowsfont Lib "c:\windows\system\tsclib.dll" (ByVal X As Integer, ByVal y As Integer, ByVal fontheight As Integer, _
ByVal rotation As Integer, _
ByVal fontstyle As Integer, _
ByVal fontunderline As Integer, _
ByVal FaceName As String, _
ByVal TextContent As String)

Public gstr_posicaoNFeErro_telaAcelerador As String

Public Function TratarCaracteresEspeciais0001(sDado As String) As String
On Error GoTo Erro:

  sDado = Replace(sDado, "#225;", "á")
  sDado = Replace(sDado, "#227;", "ã")
  sDado = Replace(sDado, "#231;", "ç")
  sDado = Replace(sDado, "#233;", "é")
  
  TratarCaracteresEspeciais0001 = sDado
  
  Exit Function
Erro:
  MsgBox "Erro no metodo de TratarCaracteresEspeciais0001. Cod: " & Err.Number & " Desc:" & Err.Description, vbCritical, "Erro"
End Function

Public Function ProgramaFidelidadeEmpresaGrupoValida()
On Error GoTo Erro:

  Dim rsProgramaEmpGrupo As ADODB.Recordset
  Dim sSql As String
  
  If Not IsNull(gIndicadorProgramaFidelidadeCNPJPrincipal) And (gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Or gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Or gIndicadorProgramaFidelidadeCNPJPrincipal = 3) Then
    Exit Function
  End If

  'Verifica se o CNPJ logado é o CNPJ principal, ou seja, o criador do programa fidelidade
  sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
  sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"

  Set rsProgramaEmpGrupo = New ADODB.Recordset
  rsProgramaEmpGrupo.Open sSql, gDB_SQLSERVER

  If rsProgramaEmpGrupo.EOF And rsProgramaEmpGrupo.BOF Then
    gIndicadorProgramaFidelidadeCNPJPrincipal = 3  ' NADA
  Else
    gIndicadorProgramaFidelidadeCNPJPrincipal = 1  ' CNPJ PRINCIPAL
    gCNPJProgramaFidelidadeCNPJPrincipal = rsProgramaEmpGrupo.Fields("CNPJ").Value
    rsProgramaEmpGrupo.Close
    Set rsProgramaEmpGrupo = Nothing
    Exit Function
  End If
  rsProgramaEmpGrupo.Close
  Set rsProgramaEmpGrupo = Nothing
  
  'Entao verifica se o CNPJ logado é um CNPJ vinculado a um CNPJ Principal (pode fazer parte do grupo)
  sSql = "SELECT * FROM [ProgramaFidelidade_empresaGrupo] "
  sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "'"

  Set rsProgramaEmpGrupo = New ADODB.Recordset
  rsProgramaEmpGrupo.Open sSql, gDB_SQLSERVER

  If rsProgramaEmpGrupo.EOF And rsProgramaEmpGrupo.BOF Then
    gIndicadorProgramaFidelidadeCNPJPrincipal = 3  ' NADA
  Else
    gIndicadorProgramaFidelidadeCNPJPrincipal = 2  ' CNPJ VINCULADO APENAS
    gCNPJProgramaFidelidadeCNPJPrincipal = rsProgramaEmpGrupo.Fields("CNPJ_principal").Value
    rsProgramaEmpGrupo.Close
    Set rsProgramaEmpGrupo = Nothing
    Exit Function
  End If
  rsProgramaEmpGrupo.Close
  Set rsProgramaEmpGrupo = Nothing
  
  
  
  Exit Function
Erro:
  MsgBox "Erro no metodo de ProgdramaFidelidadeEmpresaGrupo. Cod: " & Err.Number & " Desc:" & Err.Description, vbCritical, "Erro"

End Function

Public Function ProgramaFidelidadeCriarLancamento(ByVal codOperacao As Integer, ByVal vlCompraCliente As Double, ByVal CodCliente As Long, ByVal cpf_cgcCliente As String, ByVal CodOperador As Long, ByVal CodSequenciaVenda As Long, ByVal p_nmCliente As String)
On Error GoTo Erro:
  Dim sSql As String
  Dim rsProgramaOpSaida As ADODB.Recordset
  Dim rsPrograma As ADODB.Recordset
  Dim rsClienteNaoPart As ADODB.Recordset
  Dim lCodPrograma As Long
  Dim dVl_ProgFidelidade As Double
  Dim dVl_ProgFidelidadeParaCadaPonto As Double
  Dim l_Nm_PontosAdquiridos As Long
  Dim d_Vl_SaldoEmReais As Double
  Dim dtValidadeInicio As Date
  Dim dtValidadeFim As Date

  ' Abrir conexão com o banco de dados SQL SERVER
  gnOpenDB_SQLSERVER
  
  ProgramaFidelidadeEmpresaGrupoValida

  sSql = "SELECT count(*) FROM ProgramaFidelidade_OperacaoSaida "
  sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and Cd_OperacaoSaida=" & codOperacao

  Set rsProgramaOpSaida = New ADODB.Recordset
  rsProgramaOpSaida.Open sSql, gDB_SQLSERVER
  'Set rsProgramaOpSaida = db_SQLSERVER.OpenRecordset(sSql, dbOpenDynaset, dbSeeChanges)

  If rsProgramaOpSaida.Fields(0).Value = 0 Then
      ' Operacao de Saída informada não esta vinculada ao programa de fidelidade da empresa
      rsProgramaOpSaida.Close
      Set rsProgramaOpSaida = Nothing
      
      ' Fechar conexão com o banco de dados SQL SERVER
      'gnCloseDB_SQLSERVER
      
      Exit Function
  End If
  rsProgramaOpSaida.Close
  Set rsProgramaOpSaida = Nothing
  
  If gIndicadorProgramaFidelidadeCNPJPrincipal = 1 Then  '1-CNPJ PRINCIPAL;  2-CNPJ VINCULADO;   3-NADA
    ' Buscar o programa de fidelidade da empresa que esta ATIVO (só pode haver um programa ATIVO)
    sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
    sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_status]= 1 "
  ElseIf gIndicadorProgramaFidelidadeCNPJPrincipal = 2 Then
    ' Buscar o programa de fidelidade da empresa que esta ATIVO (só pode haver um programa ATIVO)
    sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
    sSql = sSql & " WHERE CNPJ = '" & gCNPJProgramaFidelidadeCNPJPrincipal & "' and [Cd_status]= 1 "
  End If

  Set rsPrograma = New ADODB.Recordset
  rsPrograma.Open sSql, gDB_SQLSERVER
  'Set rsPrograma = db_SQLSERVER.OpenRecordset(sSql, dbOpenDynaset, dbSeeChanges)
  
  If rsPrograma.BOF And rsPrograma.EOF Then
      ' Não existe um programa de fidelidade da empresa ATIVO
      rsPrograma.Close
      Set rsPrograma = Nothing
      
      ' Fechar conexão com o banco de dados SQL SERVER
      'gnCloseDB_SQLSERVER

      Exit Function
  End If
  lCodPrograma = rsPrograma.Fields("Cd_programa").Value
  dVl_ProgFidelidade = rsPrograma.Fields("Vl_ProgFidelidade").Value
  dVl_ProgFidelidadeParaCadaPonto = rsPrograma.Fields("Vl_ProgFidelidadeParaCadaPonto").Value
  dtValidadeInicio = rsPrograma.Fields("Dt_IniPrograma").Value
  dtValidadeFim = rsPrograma.Fields("Dt_FimPrograma").Value
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  If Now < dtValidadeInicio Or Now > dtValidadeFim Then
      'Lançamento não será realizado pois mesmo havendo um programa de fidelidade...ele esta fora do prazo

      ' Fechar conexão com o banco de dados SQL SERVER
      'gnCloseDB_SQLSERVER

      Exit Function
  End If
    

  'Calcular PONTOS
  l_Nm_PontosAdquiridos = Int(vlCompraCliente / dVl_ProgFidelidade)
  d_Vl_SaldoEmReais = l_Nm_PontosAdquiridos * dVl_ProgFidelidadeParaCadaPonto


  'Verificar se é um cliente identificado ou não
  'caso não, então não realizar o lançamento no programa de fidelidade...
  sSql = "SELECT * FROM [ProgramaFidelidade_ClienteNaoParticipa] "
  sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_Cliente]= " & CodCliente

  Set rsClienteNaoPart = New ADODB.Recordset
  rsClienteNaoPart.Open sSql, gDB_SQLSERVER
  'Set rsClienteNaoPart = db_SQLSERVER.OpenRecordset(sSql, dbOpenDynaset, dbSeeChanges)
  
  If Not (rsClienteNaoPart.BOF And rsClienteNaoPart.EOF) Then
      ' Encontrou o cliente na tabela de NÃO PODE PARTICIPAR DO PROGRAMA
      rsClienteNaoPart.Close
      Set rsClienteNaoPart = Nothing
      
      ' Fechar conexão com o banco de dados SQL SERVER
      'gnCloseDB_SQLSERVER

      Exit Function
  End If
  rsClienteNaoPart.Close
  Set rsClienteNaoPart = Nothing

    
  cpf_cgcCliente = Replace(cpf_cgcCliente, "/", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, "-", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, ".", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, ",", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, "\", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, ";", "")
  
  'Gerar lançamento no programa de fidelidade
  sSql = "Insert into [ProgramaFidelidade_lancamentos] ([CNPJ],[CPF_CGC_CLIENTE],[Cd_programa],"
  sSql = sSql & "[Cd_cliente],[Dt_criacao],[Vl_CompraCliente],[Nm_PontosAdquiridos],"
  sSql = sSql & "[Vl_SaldoEmReais],[Tp_lancamento],[Cd_operador],[Cd_guid_resgate],[Cd_SequenciaVenda], Nm_cliente) "
  sSql = sSql & " VALUES ('" & gCNPJ_CPFControleDeLicencaWebApi & "','" & cpf_cgcCliente & "'," & lCodPrograma & ", "
  sSql = sSql & CodCliente & ",convert(datetime,'" & Now & "', 103)," & Replace(vlCompraCliente, ",", ".") & "," & l_Nm_PontosAdquiridos & ","
  sSql = sSql & Replace(d_Vl_SaldoEmReais, ",", ".") & ",1," & CodOperador & ",'" & gCdGuidResgate & "'," & CodSequenciaVenda & ",'" & p_nmCliente & "' )"

  'db_SQLSERVER.Execute sSql
  Dim cmd As New ADODB.command
  cmd.ActiveConnection = gDB_SQLSERVER
  cmd.CommandText = sSql
  cmd.CommandType = adCmdText
  cmd.Execute
  Set cmd = Nothing

  ' Fechar conexão com o banco de dados SQL SERVER
  'gnCloseDB_SQLSERVER

  Exit Function
Erro:
  MsgBox "Erro no metodo de ProgramaFidelidadeCriarLancamento. Cod: " & Err.Number & " Desc:" & Err.Description, vbCritical, "Erro"
End Function

Public Function ProgramaFidelidadeCriarLancamentoRESGATE(ByVal cpf_cgcCliente As String, ByVal CodProgramaFidelidade As Long, ByVal numPontosResgate As Long, ByVal vlResgateCliente As Double, ByVal p_nmCliente As String) As Integer
On Error GoTo Erro:
  Dim sSql As String
  Dim rsPrograma As ADODB.Recordset
  Dim dtValidadeTroca As Date
  Dim cdGuidResgate As String
  Dim boDBSQLSERVER_jaESTAVAABERTO As Boolean

  boDBSQLSERVER_jaESTAVAABERTO = False

  If gDB_SQLSERVER.State = 1 Then
    boDBSQLSERVER_jaESTAVAABERTO = True
  Else
    ' Abrir conexão com o banco de dados SQL SERVER
    gnOpenDB_SQLSERVER
  End If
  
  ' Verificar se é possível fazer o RESGATE...Se esta dentro do prazo
  sSql = "SELECT * FROM [ProgramaFidelidade_empresa] "
  'sSql = sSql & " WHERE CNPJ = '" & gCNPJ_CPFControleDeLicencaWebApi & "' and [Cd_programa]=" & CodProgramaFidelidade
  sSql = sSql & " WHERE [Cd_programa]=" & CodProgramaFidelidade

  Set rsPrograma = New ADODB.Recordset
  rsPrograma.Open sSql, gDB_SQLSERVER
  
  If rsPrograma.BOF And rsPrograma.EOF Then
      ' Não existe um programa de fidelidade da empresa ATIVO
      rsPrograma.Close
      Set rsPrograma = Nothing
      
      If boDBSQLSERVER_jaESTAVAABERTO = False Then
        ' Fechar conexão com o banco de dados SQL SERVER
        'gnCloseDB_SQLSERVER
      End If

      Exit Function
  End If
  dtValidadeTroca = rsPrograma.Fields("Dt_PrazoLimiteTrocaPontos").Value
  rsPrograma.Close
  Set rsPrograma = Nothing
  
  If Now > dtValidadeTroca Then
      MsgBox "Prazo para RESGATE esta expirado!", vbInformation, "Atenção"
      ProgramaFidelidadeCriarLancamentoRESGATE = -1
      Exit Function
  End If
  
  cpf_cgcCliente = Replace(cpf_cgcCliente, "/", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, "-", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, ".", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, ",", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, "\", "")
  cpf_cgcCliente = Replace(cpf_cgcCliente, ";", "")
  
  'Gerar um GUID (cod unico = chave do RESGATE)
  cdGuidResgate = Replace(Now, " ", "")
  cdGuidResgate = Mid(cpf_cgcCliente, 1, 8) & cdGuidResgate
  cdGuidResgate = Replace(cdGuidResgate, "/", "")
  cdGuidResgate = Replace(cdGuidResgate, ":", "")
  cdGuidResgate = Replace(cdGuidResgate, "-", "")
  cdGuidResgate = Replace(cdGuidResgate, ",", "")
  cdGuidResgate = Replace(cdGuidResgate, ";", "")
  cdGuidResgate = Replace(cdGuidResgate, ".", "")
  cdGuidResgate = Replace(cdGuidResgate, "_", "")
  
  'Gerar lançamento no programa de fidelidade
  sSql = "Insert into [ProgramaFidelidade_lancamentos] ([CNPJ],[CPF_CGC_CLIENTE],[Cd_programa],"
  sSql = sSql & "[Cd_cliente],[Dt_criacao],[Vl_CompraCliente],[Nm_PontosAdquiridos],"
  sSql = sSql & "[Vl_SaldoEmReais],[Tp_lancamento],[Cd_operador],[Cd_guid_resgate], Nm_cliente) "
  sSql = sSql & " VALUES ('" & gCNPJ_CPFControleDeLicencaWebApi & "','" & cpf_cgcCliente & "'," & CodProgramaFidelidade & ", "
  sSql = sSql & "0, convert(datetime,'" & Now & "', 103),0," & numPontosResgate & ","
  sSql = sSql & Replace(vlResgateCliente, ",", ".") & ",2," & gnUserCode & ",'" & cdGuidResgate & "','" & p_nmCliente & "')"

  'db_SQLSERVER.Execute sSql
  Dim cmd As New ADODB.command
  cmd.ActiveConnection = gDB_SQLSERVER
  cmd.CommandText = sSql
  cmd.CommandType = adCmdText
  cmd.Execute
  Set cmd = Nothing

  If boDBSQLSERVER_jaESTAVAABERTO = False Then
    ' Fechar conexão com o banco de dados SQL SERVER
    'gnCloseDB_SQLSERVER
  End If

  gCdGuidResgate = cdGuidResgate
  
  ProgramaFidelidadeCriarLancamentoRESGATE = 0

  Exit Function
Erro:
  MsgBox "Erro no metodo de ProgramaFidelidadeCriarLancamentoRESGATE. Cod: " & Err.Number & " Desc:" & Err.Description, vbCritical, "Erro"
  ProgramaFidelidadeCriarLancamentoRESGATE = -1
End Function

Public Function TrataCaracteresEspeciaisASCII_Traducao(ByVal pDado As String) As String

  pDado = Replace(pDado, "&amp;", " ")
  pDado = Replace(pDado, "#39;", " ")
  pDado = Replace(pDado, "'", "*")


  'Primeiro traduzir frases prontas...
  pDado = Replace(pDado, "start tag on line 1 position", "abriu a TAG XML na posição ")
  pDado = Replace(pDado, "does not match the end tag of", "e não fechou a TAG XML de ")
  pDado = Replace(pDado, "element is invalid - The value", "esta incorreto - O valor")
  pDado = Replace(pDado, "is invalid according to its datatype", "esta inválido de acordo com o TipoDeDado")
  pDado = Replace(pDado, "The Pattern constraint failed.", "Regra/Restrição falhou.")
  pDado = Replace(pDado, "http://www.portalfiscal.inf.br/nfe:", "")
  
  'No final (agora)...então traduzir palavras soltas
  pDado = Replace(pDado, "position", "posição")
  pDado = Replace(pDado, "Line", "Linha")
  pDado = Replace(pDado, "The   ", "")
  pDado = Replace(pDado, "The  ", "")
  pDado = Replace(pDado, "  ", " ")
  pDado = Replace(pDado, "  ", " ")
  pDado = Replace(pDado, "  ", " ")

  TrataCaracteresEspeciaisASCII_Traducao = pDado
End Function

Public Function leArquivoERP_APP_QUICKTORE()
  On Error GoTo ErroGeral
  
  Dim sINTEGRACAO As String
  
  ' Abrir arquivo .txt para recuperar o CODIGO DO CLIENTE DA A3
  Dim ff As Integer
  ff = FreeFile
  Open App.Path + "\ERP_APP_QUICKSTORE_Config.txt" For Input As #ff
    Line Input #ff, sINTEGRACAO
    Line Input #ff, gCodClienteA3CadResult
  Close #ff
  
  'INTEGRACAO_APP_ERP_QUICKSTORE=SIM
  sINTEGRACAO = Mid(sINTEGRACAO, 31, 3)
  
  If sINTEGRACAO = "SIM" Then
    gINTEGRACAO_APP_ERR_QUICK = True
  Else
    gINTEGRACAO_APP_ERR_QUICK = False
  End If
  
  'CODIGO_CLIENTE=100
  gCodClienteA3CadResult = Mid(gCodClienteA3CadResult, 16, Len(gCodClienteA3CadResult) - 15)

  

  Exit Function
  
ErroGeral:
  MsgBox "Erro na leitura do arquivo ERP_APP_QUICKSTORE_Config.txt. Verifique se o mesmo existe na pasta.", vbCritical
End Function

' Exemplo de saída de um XML de uma venda simples...
'<NFe xmlns="http://www.portalfiscal.inf.br/nfe"><infNFe Id="NFe15171004333691000104550010000000021362235676" versao="3.10"><ide><cUF>15</cUF><cNF>36223567</cNF><natOp>Venda Estadual</natOp><indPag>2</indPag><mod>55</mod><serie>1</serie><nNF>2</nNF><dhEmi>2017-10-02T21:49:58-03:00</dhEmi><dhSaiEnt>2017-10-02T21:49:58-03:00</dhSaiEnt><tpNF>1</tpNF><idDest>2</idDest><cMunFG>0</cMunFG><tpImp>1</tpImp><tpEmis>1</tpEmis><cDV>6</cDV><tpAmb>2</tpAmb><finNFe>1</finNFe><indFinal>1</indFinal><indPres>1</indPres><procEmi>1</procEmi><verProc>3.10</verProc></ide>
'<emit><CNPJ>04333691000104</CNPJ><xNome>PESCARE DISTRIBUIDORA</xNome><xFant>PESCARE DISTRIBUIDORA</xFant><enderEmit><xLgr>Av. Marechal Floriano Peixoto 1389</xLgr><nro></nro><xBairro>Centro.</xBairro><cMun>0</cMun><xMun>Bragança.</xMun><UF>PA</UF><CEP></CEP><fone>34253178</fone></enderEmit><IE>152171223</IE><CRT>1</CRT></emit>
'<dest><CNPJ>96734850000192</CNPJ><xNome>SÃO LEOPOLDO</xNome><enderDest><xLgr>Av. São Borja, 372 </xLgr><nro></nro><xBairro>Rio Branco </xBairro><cMun>0</cMun><xMun>Rio Grande do Sul </xMun><UF>RS</UF><CEP>93032000</CEP><fone>0513588441</fone></enderDest><indIEDest>1</indIEDest><IE>1240007288</IE><email>csl@csL.com.br</email></dest>
'<det nItem="1"><prod><cProd>557</cProd><cEAN /><xProd>Aco Mole Nº 14</xProd><NCM></NCM><CFOP></CFOP><uCom>RL</uCom><qCom>1.0000</qCom><vUnCom>18.0200000000</vUnCom><vProd>18.02</vProd><cEANTrib /><uTrib>RL</uTrib><qTrib>1.0000</qTrib><vUnTrib>18.0200000000</vUnTrib><indTot>1</indTot></prod>
'<imposto><ICMS><ICMSSN102><orig> </orig><CSOSN>103</CSOSN></ICMSSN102>
'</ICMS><PIS><PISNT><CST>07</CST></PISNT></PIS><COFINS><COFINSNT><CST>07</CST></COFINSNT></COFINS>
'<total><ICMSTot><vBC>18.02</vBC><vICMS>2.16</vICMS><vICMSDeson>0.00</vICMSDeson><vBCST>0.00</vBCST><vST>0.00</vST><vProd>18.02</vProd><vFrete>0.00</vFrete><vSeg>0.00</vSeg><vDesc>0.00</vDesc><vII>0.00</vII><vIPI>0.00</vIPI><vPIS>0.00</vPIS><vCOFINS>0.00</vCOFINS><vOutro>0.00</vOutro><vNF>18.02</vNF></ICMSTot></total>
'<transp><modFrete>1</modFrete>
'<cobr><fat><nFat>2</nFat><vOrig>18.02</vOrig><vLiq>18.02</vLiq></fat><dup><nDup>0</nDup><dVenc>2017-10-02</dVenc><vDup>18.02</vDup></dup>
'<infAdic><infCpl>Val Aprox dos Tributos R$0(0%) Fonte IBPT.</infCpl></infAdic></infNFe></NFe>

Public Function func_horarioFuso(ByVal bHorarioDeVerao As Boolean, ByVal Estado As String, ByRef sFuso As String)

  If UCase(Estado) = "RS" Or UCase(Estado) = "SC" Or UCase(Estado) = "PR" Or _
      UCase(Estado) = "SP" Or UCase(Estado) = "RJ" Or UCase(Estado) = "ES" Or _
      UCase(Estado) = "MG" Or UCase(Estado) = "GO" Or UCase(Estado) = "DF" Then
    If bHorarioDeVerao = True Then
      sFuso = "-02:00"
    Else
      sFuso = "-03:00"
    End If
  ElseIf UCase(Estado) = "MS" Or UCase(Estado) = "MT" Then
    If bHorarioDeVerao = True Then
      sFuso = "-03:00"
    Else
      sFuso = "-04:00"
    End If
  ElseIf UCase(Estado) = "AP" Or UCase(Estado) = "PA" Or UCase(Estado) = "MA" Or _
  UCase(Estado) = "PI" Or UCase(Estado) = "CE" Or UCase(Estado) = "RN" Or _
  UCase(Estado) = "PB" Or UCase(Estado) = "PE" Or UCase(Estado) = "AL" Or _
  UCase(Estado) = "SE" Or UCase(Estado) = "BA" Or UCase(Estado) = "TO" Then
    If bHorarioDeVerao = True Then
      sFuso = "-03:00"
    Else
      sFuso = "-03:00"
    End If
  ElseIf UCase(Estado) = "AC" Then
    If bHorarioDeVerao = True Then
      sFuso = "-05:00"
    Else
      sFuso = "-05:00"
    End If
  ElseIf UCase(Estado) = "RO" Or UCase(Estado) = "AM" Or UCase(Estado) = "RR" Then
    If bHorarioDeVerao = True Then
      sFuso = "-04:00"
    Else
      sFuso = "-04:00"
    End If
  End If
  
End Function

Public Function func_validaCamposObrigatoriosNFE_3_10(ByVal clcArquivo As String, ByRef sRetorno As String)
  Dim varLinha As Variant
  Dim strSaida As String
  Dim sValida As String
  Dim iIndice1 As Integer
  Dim iIndice2 As Integer
  Dim iIndice3 As Integer
  Dim sEmitSaida As String
  Dim sDestSaida As String
  Dim sDetItemSaida As String
  Dim sTotalSaida As String
  
  '  For Each varLinha In clcArquivo
  '    strSaida = strSaida + CStr(varLinha)
  '  Next
  strSaida = clcArquivo
  
  ' ******************************************************************
  ' Trata TAG <emit> - GRUPO OBRIGATORIO
  iIndice1 = InStr(1, strSaida, "<emit>")
  If iIndice1 > 0 Then
      iIndice2 = InStr(iIndice1 + 3, strSaida, "</emit>")
  End If
  sValida = Mid(strSaida, iIndice1 + Len("<emit>"), iIndice2 - (iIndice1 + Len("<emit>")))
  
  func_validaCamposObrigatoriosNFE_3_10_emit sValida, sEmitSaida
  
  If sEmitSaida <> "" Then
    sRetorno = sRetorno + "     -> Emitente: " + sEmitSaida & Chr(13)
  End If
  
  
  ' ******************************************************************
  ' Trata TAG <dest> - GRUPO OPCIONAL
  iIndice1 = InStr(1, strSaida, "<dest>")
  If iIndice1 > 0 Then
      iIndice2 = InStr(iIndice1 + 3, strSaida, "</dest>")
  
      sValida = Mid(strSaida, iIndice1 + Len("<dest>"), iIndice2 - (iIndice1 + Len("<dest>")))
      
      func_validaCamposObrigatoriosNFE_3_10_dest sValida, sDestSaida
      
      If sDestSaida <> "" Then
        sRetorno = sRetorno + "     -> Destinatário: " + sDestSaida & Chr(13)
      End If
  End If
  
  
  ' ******************************************************************
  ' Trata TAG <dest> - GRUPO OBRIGATORIO
  iIndice1 = InStr(1, strSaida, "<det nItem")
  While iIndice1 > 0
    If iIndice1 > 0 Then
        iIndice3 = InStr(iIndice1 + 3, strSaida, ">")
        iIndice2 = InStr(iIndice1 + 3, strSaida, "</det>")
    
        sValida = Mid(strSaida, iIndice3, iIndice2 - iIndice3)
        
        func_validaCamposObrigatoriosNFE_3_10_detItem sValida, sDetItemSaida
        
        If sDetItemSaida <> "" Then
          sRetorno = sRetorno + "     -> Produtos: " + sDetItemSaida & Chr(13)
        End If
    End If
    
    iIndice1 = InStr(iIndice2, strSaida, "<det nItem")
  Wend

  ' ******************************************************************
  ' Trata TAG <total> - GRUPO OBRIGATORIO
  iIndice1 = InStr(1, strSaida, "<total>")
  If iIndice1 > 0 Then
    iIndice3 = InStr(iIndice1 + 3, strSaida, "<ICMSTot>")
    iIndice2 = InStr(iIndice1 + 3, strSaida, "</ICMSTot>")

    sValida = Mid(strSaida, iIndice3 + Len("<ICMSTot>"), iIndice2 - (iIndice3 + Len("<ICMSTot>")))

    func_validaCamposObrigatoriosNFE_3_10_total sValida, sTotalSaida
    
    If sTotalSaida <> "" Then
      sRetorno = sRetorno + "     -> Total: " + sTotalSaida & Chr(13)
    End If
  End If

End Function

Public Function func_validaCamposObrigatoriosNFE_3_10_emit(ByVal sEmit As String, ByRef sEmitSaida As String)
  Dim i As Integer
  Dim strArray(6, 2) As String
  Dim strArrayAgrup1(11, 2) As String
  Dim strArrayAgrup2(2, 2) As String
  
  ' SIM   - obrigatorio
  ' NAO   - opcional
    
  ' Tratamento TAG principal <EMIT>
  ' <CNPJ> ou <CPF>         - SIM
  ' <xNome>                 - SIM
  ' <xFant>                 - NAO
  ' <enderEmit>             - SIM             ....ESTE NÃO COLOCA NO ARRAY
  '    <xLgr>                   - SIM
  '    <nro>                    - SIM
  '    <xCpl>                   - NAO
  '    <xBairro>                - SIM
  '    <cMun>                   - SIM
  '    <xMun>                   - SIM
  '    <UF>                     - SIM
  '    <CEP>                    - SIM
  '    <cPais>                  - NAO
  '    <xPais>                  - NAO
  '    <fone>                   - NAO
  ' <IE>                    - SIM
  ' <IEST>                  - NAO
  ' Agrupamento (inicio)    - NAO       ....ESTE NÃO COLOCA NO ARRAY
  '     <IM>                    - SIM
  '     <CNAE>                  - NAO
  ' Agrupamento (fim)
  ' <CRT>                   - SIM
  
  strArray(0, 0) = "CNPJ"
  strArray(0, 1) = "S"
  strArray(1, 0) = "xNome"
  strArray(1, 1) = "S"
  strArray(2, 0) = "xFant"
  strArray(2, 1) = "N"
  strArrayAgrup1(0, 0) = "xLgr"
  strArrayAgrup1(0, 1) = "S"
  strArrayAgrup1(1, 0) = "nro"
  strArrayAgrup1(1, 1) = "S"
  strArrayAgrup1(2, 0) = "xCpl"
  strArrayAgrup1(2, 1) = "N"
  strArrayAgrup1(3, 0) = "xBairro"
  strArrayAgrup1(3, 1) = "S"
  strArrayAgrup1(4, 0) = "cMun"
  strArrayAgrup1(4, 1) = "S"
  strArrayAgrup1(5, 0) = "xMun"
  strArrayAgrup1(5, 1) = "S"
  strArrayAgrup1(6, 0) = "UF"
  strArrayAgrup1(6, 1) = "S"
  strArrayAgrup1(7, 0) = "CEP"
  strArrayAgrup1(7, 1) = "S"
  strArrayAgrup1(8, 0) = "cPais"
  strArrayAgrup1(8, 1) = "N"
  strArrayAgrup1(9, 0) = "xPais"
  strArrayAgrup1(9, 1) = "N"
  strArrayAgrup1(10, 0) = "fone"
  strArrayAgrup1(10, 1) = "N"
  strArray(3, 0) = "IE"
  strArray(3, 1) = "S"
  strArray(4, 0) = "IEST"
  strArray(4, 1) = "N"
  strArrayAgrup2(0, 0) = "IM"
  strArrayAgrup2(0, 1) = "S"
  strArrayAgrup2(1, 0) = "CNAE"
  strArrayAgrup2(1, 1) = "N"
  strArray(5, 0) = "CRT"
  strArray(5, 1) = "S"
  
  Dim iIndice_1 As Integer
  Dim iIndice_2 As Integer
  Dim sConteudo As String
  
  For i = 0 To 5
    If i = 0 Then
        ' Tratamento para escolha CNPJ ou CPF
        If strArray(i, 1) = "S" Then
          iIndice_1 = InStr(1, sEmit, "<" + strArray(i, 0) + ">")
          If iIndice_1 < 0 Then
            strArray(i, 0) = "CPF"
            iIndice_1 = InStr(1, sEmit, "<" + strArray(i, 0) + ">")
          End If
          
          If iIndice_1 > 0 Then
            iIndice_2 = InStr(1, sEmit, "</" + strArray(i, 0) + ">")
            
            If iIndice_2 > 0 Then
              sConteudo = ""
              sConteudo = Mid(sEmit, iIndice_1 + Len(strArray(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArray(i, 0)) + 2))
              If Trim(sConteudo) = "" Then
                sEmitSaida = sEmitSaida + " *" + strArray(i, 0)
              End If
            Else
              sEmitSaida = sEmitSaida + " *" + strArray(i, 0)
            End If
          Else
            sEmitSaida = sEmitSaida + " *" + strArray(i, 0)
          End If
        End If
    Else
      If strArray(i, 1) = "S" Then
        iIndice_1 = InStr(1, sEmit, "<" + strArray(i, 0) + ">")
        If iIndice_1 > 0 Then
          iIndice_2 = InStr(1, sEmit, "</" + strArray(i, 0) + ">")
          
          If iIndice_2 > 0 Then
            sConteudo = ""
            sConteudo = Mid(sEmit, iIndice_1 + Len(strArray(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArray(i, 0)) + 2))
            If Trim(sConteudo) = "" Then
              sEmitSaida = sEmitSaida + " *" + strArray(i, 0)
            End If
          Else
            sEmitSaida = sEmitSaida + " *" + strArray(i, 0)
          End If
        Else
          sEmitSaida = sEmitSaida + " *" + strArray(i, 0)
        End If
      End If
    End If
  Next
  
  i = 0
  For i = 0 To 10
    If strArrayAgrup1(i, 1) = "S" Then
        iIndice_1 = InStr(1, sEmit, "<" + strArrayAgrup1(i, 0) + ">")
        If iIndice_1 > 0 Then
          iIndice_2 = InStr(1, sEmit, "</" + strArrayAgrup1(i, 0) + ">")
          
          If iIndice_2 > 0 Then
            sConteudo = ""
            sConteudo = Mid(sEmit, iIndice_1 + Len(strArrayAgrup1(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArrayAgrup1(i, 0)) + 2))
            If Trim(sConteudo) = "" Then
              sEmitSaida = sEmitSaida + " *" + strArrayAgrup1(i, 0)
            End If
          Else
            sEmitSaida = sEmitSaida + " *" + strArrayAgrup1(i, 0)
          End If
        Else
          sEmitSaida = sEmitSaida + " *" + strArrayAgrup1(i, 0)
        End If
    End If
  Next
  
'  i = 0
'  For i = 0 To 1
'    If strArrayAgrup2(i, 1) = "S" Then
'        iIndice_1 = InStr(1, sEmit, "<" + strArrayAgrup2(i, 0) + ">")
'        If iIndice_1 > 0 Then
'          iIndice_2 = InStr(1, sEmit, "</" + strArrayAgrup2(i, 0) + ">")
'
'          If iIndice_2 > 0 Then
'            sConteudo = ""
'            sConteudo = Mid(sEmit, iIndice_1 + Len(strArrayAgrup2(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArrayAgrup2(i, 0)) + 2))
'            If Trim(sConteudo) = "" Then
'              sEmitSaida = sEmitSaida + " *" + strArrayAgrup2(i, 0)
'            End If
'          Else
'            sEmitSaida = sEmitSaida + " *" + strArrayAgrup2(i, 0)
'          End If
'        Else
'          sEmitSaida = sEmitSaida + " *" + strArrayAgrup2(i, 0)
'        End If
'    End If
'  Next
  
End Function

Public Function func_validaCamposObrigatoriosNFE_3_10_dest(ByVal sDest As String, ByRef sDestSaida As String)
  Dim i As Integer
  Dim strArray(7, 2) As String
  Dim strArrayAgrup1(11, 2) As String
  
  ' SIM   - obrigatorio
  ' NAO   - opcional
    
  ' Tratamento TAG principal <DEST>
  ' <CNPJ> ou <CPF> ou <idEstrangeiro> - SIM
  ' <xNome>                            - NAO
  ' <enderDest>                        - NAO             ....ESTE NÃO COLOCA NO ARRAY
  '    <xLgr>                   - SIM
  '    <nro>                    - SIM
  '    <xCpl>                   - NAO
  '    <xBairro>                - SIM
  '    <cMun>                   - SIM
  '    <xMun>                   - SIM
  '    <UF>                     - SIM
  '    <CEP>                    - SIM
  '    <cPais>                  - NAO
  '    <xPais>                  - NAO
  '    <fone>                   - NAO
  ' <indIEDest>                       - SIM
  ' <IE>                              - NAO
  ' <ISUF>                            - NAO
  ' <IM>                              - NAO
  ' <email>                           - NAO
  
  strArray(0, 0) = "CNPJ"
  strArray(0, 1) = "S"
  strArray(1, 0) = "xNome"
  strArray(1, 1) = "N"
  strArrayAgrup1(0, 0) = "xLgr"
  strArrayAgrup1(0, 1) = "S"
  strArrayAgrup1(1, 0) = "nro"
  strArrayAgrup1(1, 1) = "S"
  strArrayAgrup1(2, 0) = "xCpl"
  strArrayAgrup1(2, 1) = "N"
  strArrayAgrup1(3, 0) = "xBairro"
  strArrayAgrup1(3, 1) = "S"
  strArrayAgrup1(4, 0) = "cMun"
  strArrayAgrup1(4, 1) = "S"
  strArrayAgrup1(5, 0) = "xMun"
  strArrayAgrup1(5, 1) = "S"
  strArrayAgrup1(6, 0) = "UF"
  strArrayAgrup1(6, 1) = "S"
  strArrayAgrup1(7, 0) = "CEP"
  strArrayAgrup1(7, 1) = "S"
  strArrayAgrup1(8, 0) = "cPais"
  strArrayAgrup1(8, 1) = "N"
  strArrayAgrup1(9, 0) = "xPais"
  strArrayAgrup1(9, 1) = "N"
  strArrayAgrup1(10, 0) = "fone"
  strArrayAgrup1(10, 1) = "N"
  strArray(2, 0) = "indIEDest"
  strArray(2, 1) = "S"
  strArray(3, 0) = "IE"
  strArray(3, 1) = "N"
  strArray(4, 0) = "ISUF"
  strArray(4, 1) = "N"
  strArray(5, 0) = "IM"
  strArray(5, 1) = "N"
  strArray(6, 0) = "email"
  strArray(6, 1) = "N"

  
  Dim iIndice_1 As Integer
  Dim iIndice_2 As Integer
  Dim sConteudo As String
  
  For i = 0 To 6
    If i = 0 Then
        ' Tratamento para escolha CNPJ ou CPF ou idEstrangeiro
        If strArray(i, 1) = "S" Then
          iIndice_1 = InStr(1, sDest, "<" + strArray(i, 0) + ">")
          If iIndice_1 < 0 Then
            strArray(i, 0) = "CPF"
            iIndice_1 = InStr(1, sDest, "<" + strArray(i, 0) + ">")
          
            If iIndice_1 < 0 Then
              strArray(i, 0) = "idEstrangeiro"
              iIndice_1 = InStr(1, sDest, "<" + strArray(i, 0) + ">")
            End If
          End If
          
          If iIndice_1 > 0 Then
            iIndice_2 = InStr(1, sDest, "</" + strArray(i, 0) + ">")
            
            If iIndice_2 > 0 Then
              sConteudo = ""
              sConteudo = Mid(sDest, iIndice_1 + Len(strArray(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArray(i, 0)) + 2))
              If Trim(sConteudo) = "" Then
                sDestSaida = sDestSaida + " *" + strArray(i, 0)
              End If
            Else
              sDestSaida = sDestSaida + " *" + strArray(i, 0)
            End If
          Else
            sDestSaida = sDestSaida + " *" + strArray(i, 0)
          End If
        End If
    Else
      If strArray(i, 1) = "S" Then
        iIndice_1 = InStr(1, sDest, "<" + strArray(i, 0) + ">")
        If iIndice_1 > 0 Then
          iIndice_2 = InStr(1, sDest, "</" + strArray(i, 0) + ">")
          
          If iIndice_2 > 0 Then
            sConteudo = ""
            sConteudo = Mid(sDest, iIndice_1 + Len(strArray(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArray(i, 0)) + 2))
            If Trim(sConteudo) = "" Then
              sDestSaida = sDestSaida + " *" + strArray(i, 0)
            End If
          Else
            sDestSaida = sDestSaida + " *" + strArray(i, 0)
          End If
        Else
          sDestSaida = sDestSaida + " *" + strArray(i, 0)
        End If
      End If
    End If
  Next
  
  i = 0
  For i = 0 To 10
    If strArrayAgrup1(i, 1) = "S" Then
        iIndice_1 = InStr(1, sDest, "<" + strArrayAgrup1(i, 0) + ">")
        If iIndice_1 > 0 Then
          iIndice_2 = InStr(1, sDest, "</" + strArrayAgrup1(i, 0) + ">")
          
          If iIndice_2 > 0 Then
            sConteudo = ""
            sConteudo = Mid(sDest, iIndice_1 + Len(strArrayAgrup1(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArrayAgrup1(i, 0)) + 2))
            If Trim(sConteudo) = "" Then
              sDestSaida = sDestSaida + " *" + strArrayAgrup1(i, 0)
            End If
          Else
            sDestSaida = sDestSaida + " *" + strArrayAgrup1(i, 0)
          End If
        Else
          sDestSaida = sDestSaida + " *" + strArrayAgrup1(i, 0)
        End If
    End If
  Next
    
End Function


Public Function func_validaCamposObrigatoriosNFE_3_10_detItem(ByVal sDetItem As String, ByRef sDetItemSaida As String)
  Dim i As Integer
  Dim strArray(11, 2) As String
  Dim sCodigoProduto As String
  Dim sSaida As String
  
  sCodigoProduto = ""
  sSaida = ""

  ' SIM   - obrigatorio
  ' NAO   - opcional
  
  ' *** Informado as tags obrigatorios e algumas opcionais
  ' Tratamento TAG principal <det nItem>
  ' <cProd>
  ' <cEAN>                      - NAO
  ' <NCM>
  ' <NVE>                       - NAO
  ' <Agrupamento opcional aqui> - NAO
  ' <cBenef>                    - NAO
  ' <EXTIPI>                    - NAO
  ' <CFOP>
  ' <uCom>
  ' <qCom>
  ' <vUnCom>
  ' <vProd>
  ' <cEANTrib>                  - NAO
  ' <uTrib>
  ' <qTrib>
  ' <vUnTrib>
  ' <indTot>
  
  strArray(0, 0) = "cProd"
  strArray(0, 1) = "S"
  strArray(1, 0) = "NCM"
  strArray(1, 1) = "S"
  strArray(2, 0) = "CFOP"
  strArray(2, 1) = "S"
  strArray(3, 0) = "uCom"
  strArray(3, 1) = "S"
  strArray(4, 0) = "qCom"
  strArray(4, 1) = "S"
  strArray(5, 0) = "vUnCom"
  strArray(5, 1) = "S"
  strArray(6, 0) = "vProd"
  strArray(6, 1) = "S"
  strArray(7, 0) = "uTrib"
  strArray(7, 1) = "S"
  strArray(8, 0) = "qTrib"
  strArray(8, 1) = "S"
  strArray(9, 0) = "vUnTrib"
  strArray(9, 1) = "S"
  strArray(10, 0) = "indTot"
  strArray(10, 1) = "S"
  
  Dim iIndice_1 As Integer
  Dim iIndice_2 As Integer
  Dim sConteudo As String
  
  For i = 0 To 10
    If strArray(i, 1) = "S" Then
      iIndice_1 = InStr(1, sDetItem, "<" + strArray(i, 0) + ">")
      If iIndice_1 > 0 Then
        iIndice_2 = InStr(1, sDetItem, "</" + strArray(i, 0) + ">")
        
        If iIndice_2 > 0 Then
          sConteudo = ""
          sConteudo = Mid(sDetItem, iIndice_1 + Len(strArray(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArray(i, 0)) + 2))
          
          If i = 0 Then
            sCodigoProduto = sConteudo
          End If
          
          If Trim(sConteudo) = "" Then
            sSaida = sSaida + " *" + strArray(i, 0)
          End If
        Else
          sSaida = sSaida + " *" + strArray(i, 0)
        End If
      Else
        sSaida = sSaida + " *" + strArray(i, 0)
      End If
    End If
  Next
  
  If Len(sSaida) > 0 Then
    sDetItemSaida = Chr(13) & "          -> Cód.Produto " + sCodigoProduto + ":" + sSaida
  End If
End Function


Public Function func_validaCamposObrigatoriosNFE_3_10_total(ByVal sTOTAL As String, ByRef sTotalSaida As String)
  Dim i As Integer
  Dim strArray(23, 2) As String
  
  ' SIM   - obrigatorio
  ' NAO   - opcional
    
  ' Tratamento TAG principal <TOTAL>
  '<total>
  '  <ICMSTot>
  '    <vBC>
  '    <vICMS>
  '    <vICMSDeson>
  '    <vFCPUFDest>  - nao
  '    <vICMSUFDest> - nao
  '    <vICMSUFRemet>  - nao
  '    <vFCP>
  '    <vBCST>
  '    <vST>
  '    <vFCPST>
  '    <vFCPSTRet>
  '    <vProd>
  '    <vFrete>
  '    <vSeg>
  '    <vDesc>
  '    <vII>
  '    <vIPI>
  '    <vIPIDevol>
  '    <vPIS>
  '    <vCOFINS>
  '    <vOutro>
  '    <vNF>
  '    <vTotTrib>  - nao
  '  </ICMSTot>
  
  strArray(0, 0) = "vBC"
  strArray(0, 1) = "S"
  strArray(1, 0) = "vICMS"
  strArray(1, 1) = "S"
  strArray(2, 0) = "vICMSDeson"
  strArray(2, 1) = "S"
  strArray(3, 0) = "vFCPUFDest"
  strArray(3, 1) = "N"
  strArray(4, 0) = "vICMSUFDest"
  strArray(4, 1) = "N"
  strArray(5, 0) = "vICMSUFRemet"
  strArray(5, 1) = "N"
  strArray(6, 0) = "vFCP"
  strArray(6, 1) = "N"
  strArray(7, 0) = "vBCST"
  strArray(7, 1) = "S"
  strArray(8, 0) = "vST"
  strArray(8, 1) = "S"
  strArray(9, 0) = "vFCPST"
  strArray(9, 1) = "N"
  strArray(10, 0) = "vFCPSTRet"
  strArray(10, 1) = "N"
  strArray(11, 0) = "vProd"
  strArray(11, 1) = "S"
  strArray(12, 0) = "vFrete"
  strArray(12, 1) = "S"
  strArray(13, 0) = "vSeg"
  strArray(13, 1) = "S"
  strArray(14, 0) = "vDesc"
  strArray(14, 1) = "S"
  strArray(15, 0) = "vII"
  strArray(15, 1) = "S"
  strArray(16, 0) = "vIPI"
  strArray(16, 1) = "S"
  strArray(17, 0) = "vIPIDevol"
  strArray(17, 1) = "N"
  strArray(18, 0) = "vPIS"
  strArray(18, 1) = "S"
  strArray(19, 0) = "vCOFINS"
  strArray(19, 1) = "S"
  strArray(20, 0) = "vOutro"
  strArray(20, 1) = "S"
  strArray(21, 0) = "vNF"
  strArray(21, 1) = "S"
  strArray(22, 0) = "vTotTrib"
  strArray(22, 1) = "N"
  
  Dim iIndice_1 As Integer
  Dim iIndice_2 As Integer
  Dim sConteudo As String
  
  For i = 0 To 22
    If strArray(i, 1) = "S" Then
      iIndice_1 = InStr(1, sTOTAL, "<" + strArray(i, 0) + ">")
      If iIndice_1 > 0 Then
        iIndice_2 = InStr(1, sTOTAL, "</" + strArray(i, 0) + ">")
        
        If iIndice_2 > 0 Then
          sConteudo = ""
          sConteudo = Mid(sTOTAL, iIndice_1 + Len(strArray(i, 0)) + 2, iIndice_2 - (iIndice_1 + Len(strArray(i, 0)) + 2))
          If Trim(sConteudo) = "" Then
            sTotalSaida = sTotalSaida + " *" + strArray(i, 0)
          End If
        Else
          sTotalSaida = sTotalSaida + " *" + strArray(i, 0)
        End If
      Else
        sTotalSaida = sTotalSaida + " *" + strArray(i, 0)
      End If
    End If
  Next
        
End Function

Public Function RemoveCaracteresEspeciaisParaNFE(ByVal strDado As String) As String
  
    strDado = Replace(strDado, "ã", "a")
    strDado = Replace(strDado, "à", "a")
    strDado = Replace(strDado, "á", "a")
    strDado = Replace(strDado, "â", "a")
    strDado = Replace(strDado, "é", "e")
    strDado = Replace(strDado, "è", "e")
    strDado = Replace(strDado, "ê", "e")
    strDado = Replace(strDado, "í", "i")
    strDado = Replace(strDado, "ì", "i")
    strDado = Replace(strDado, "ó", "o")
    strDado = Replace(strDado, "ò", "o")
    strDado = Replace(strDado, "ô", "o")
    strDado = Replace(strDado, "ú", "u")
    strDado = Replace(strDado, "ù", "u")
    strDado = Replace(strDado, "ü", "u")
   
    strDado = Replace(strDado, "Ã", "A")
    strDado = Replace(strDado, "À", "A")
    strDado = Replace(strDado, "Á", "A")
    strDado = Replace(strDado, "Â", "A")
    strDado = Replace(strDado, "É", "E")
    strDado = Replace(strDado, "È", "E")
    strDado = Replace(strDado, "Ê", "E")
    strDado = Replace(strDado, "Í", "I")
    strDado = Replace(strDado, "Ì", "I")
    strDado = Replace(strDado, "Ó", "O")
    strDado = Replace(strDado, "Ò", "O")
    strDado = Replace(strDado, "Ô", "O")
    strDado = Replace(strDado, "Ú", "U")
    strDado = Replace(strDado, "Ù", "U")
    strDado = Replace(strDado, "Ü", "U")
    
    strDado = Replace(strDado, "ç", "c")
    strDado = Replace(strDado, "Ç", "C")
   
    strDado = Replace(strDado, "&", "")
    strDado = Replace(strDado, "ª", "")
    strDado = Replace(strDado, "º", "")
   
    strDado = RTrim(LTrim(strDado))
    
    RemoveCaracteresEspeciaisParaNFE = strDado
 
End Function
