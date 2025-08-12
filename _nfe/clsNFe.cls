VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iIndex As Integer

Private sCNPJ_EmitenteXML As String

'20/08/2009 - mpdea
'Classe para geração da NF-e (Nota Fiscal Eletrônica)

Private Const VERSAO_LAYOUT_ENVIO As String = "4.00"
Private Const VERSAO_LAYOUT_INUTILIZACAO As String = "4.00"
Private Const VERSAO_LAYOUT_CANCELAMENTO As String = "1.07"
Private Const VERSAO_PROCESSO_EMISSAO As String = "4.00"

Private m_versao_layout_envio As String

'03/01/2011 - Andrea
Private m_clcMensagens As Collection

'26/04/2011 - Andrea
Private m_bln_tem_desconto As Boolean

Private isInterestadual As Boolean
Private isConsumidorFinal As Boolean
Private isIEIsento As Boolean

Private TotalVFCPUFDest As Double
Private TotalVICMSUFDest As Double
Private TotalVICMSUFRemet As Double
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private sEstadoFilial As String
Private UF_destino As String

Private str_FCP_UF_emitente As String
Private str_FCP_INS_EST_emitente As String
Private str_FCP_CNPJ_CPF_destinatario As String
Private str_FCP_RG_INS_EST_destinatario As String
Private str_FCP_UF_destinatario As String
Private str_IndicadorIE_destinatario As String

Dim AliquotaAdesaoUFDestFCP As Double
Dim AliquotaFCP_EMIT_DEST As Double
Dim AliquotaFCP_Emitente As Double
Dim AliquotaFCP_Destinatario As Double
Dim CalculoFCPEstadosDiferentes As Double
Dim CalculoFCPEstadosIguais As Double
Dim CalculoFCP_diferenca As Double
Dim CalculoFCP_participacaoDest As Double
Dim CalculoFCP_participacaoEmitente As Double

Dim iAnoFCP As Integer
Dim vBC_fcp As Double
Dim vFCPUFDest_registroTotal As Double
Dim vICMSUFDest_registroTotal As Double
Dim vFCP_registroTotal As Double
Dim vFCPST_registroTotal As Double
Dim vFCPSTRet_registroTotal As Double
Dim vResultado As Double
Dim blnSimplesNacionalGeral As Boolean
Dim rsSaidas As Recordset
Dim str_sql As String
Dim bFINALIDADE_NFE_DEVOLUCAO As Boolean
Dim iFINALIDADE_NFE As Integer
Dim sTIPO_OPERACAO As String
Dim bTEM_ICMS_3179a As Boolean
Dim vlSOMA_DO_ICMS As Double

Dim vDespAcessorias As Double
Public sXML_40 As String
Dim sPadraoArquivoIntegracao As String

Dim vTotal_Desconto_ICMS_vNF As Double

Dim vlSOMA_DO_IPI As Double

Dim bTemDescontoNaVenda As Boolean
Dim vTemDescontoNaVendaPercentual As Double
Dim vAuxTotalSaida As Double
Dim vSOMA_TOTAL_DESCONTO As Double

Dim s_vlSOMA_DO_ICMS_BaseDeCalculo As String
Dim v_vlSOMA_DO_ICMS_BaseDeCalculo As Double

Dim s_vlSOMA_DO_ICMS_ST_BaseDeCalculo As String
Dim v_vlSOMA_DO_ICMS_ST_BaseDeCalculo As Double
Dim v_vlSOMA_DO_ICMS_ST As Double

Dim localTipoSituacaoTributariaPIS As Integer

Dim s_vTotTrib As String
Dim vl_vTotTrib As Double

Private iDetNumItemXML As Long

Private inst_dblPercentualPIS As Double
Private inst_dblPercentualCofins As Double

Private aliquotaICMS_tab_ICMS_PERC_ESTADOS As String
Private bo_AliquotaICMS_interestadual As Boolean

Private bo_operacaoTransferenciaEntreEmpresas As Boolean

Private bo_CalcularIPI      As Boolean                    ' Parametros da área Impostos, da tela de Operações de Saídas
Private bo_SomarIPItotal    As Boolean                    ' Parametros da área Impostos, da tela de Operações de Saídas
Private bo_SomarIPIBC       As Boolean                    ' Parametros da área Impostos, da tela de Operações de Saídas
Private bo_CalcularICMS_ST  As Boolean                    ' Parametros da área Impostos, da tela de Operações de Saídas
Private bo_CalcularICMS     As Boolean                    ' Parametros da área Impostos, da tela de Operações de Saídas
Private iObterTributosProduto_EntradaOuSaida As Integer   ' Parametros da área Impostos, da tela de Operações de Saídas

Dim IndicadorIE As String


'17/10/2010 - Andrea
'Envia a NFe 2.0
Public Function Enviar_4_0(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal TipoMovimento As Integer, ByVal ConcatenarDescricaoAdicional As Boolean, ByVal TipoEmissao As Integer, ByVal DataInicioContingencia As Date, ByVal HoraInicioContingencia As String, ByVal MotivoContingencia As String, ByVal PadraoArquivoIntegracao As String) As Boolean
  
  If TipoMovimento = 0 Then 'Entrada
    Enviar_4_0 = EnviarEntrada_4_0(Filial, Sequencia, TipoEmissao, DataInicioContingencia, HoraInicioContingencia, MotivoContingencia, PadraoArquivoIntegracao)
  Else 'Saída
    Enviar_4_0 = EnviarSaida_4_0_REMODELADA(Filial, Sequencia, ConcatenarDescricaoAdicional, TipoEmissao, DataInicioContingencia, HoraInicioContingencia, MotivoContingencia, PadraoArquivoIntegracao)
  End If
  
End Function

'02/10/2009 - Andrea
'Cancelar a NFe
Public Function Cancelar(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal TipoMovimento As Integer, ByVal MotivoCancelamento As String) As Boolean
  
  If TipoMovimento = 0 Then 'Entrada
    Cancelar = CancelarEntrada(Filial, Sequencia)
  Else 'Saída
    Cancelar = CancelarSaida(Filial, Sequencia, MotivoCancelamento)
  End If
  
End Function

'06/10/2009 - Andrea
'Inutilizar NFe
Public Function Inutilizar(ByVal Filial As Integer, ByVal TipoMovimento As Integer, ByVal MotivoInutilizacao As String, ByVal Serie As String, ByVal NumeroNFeInicial As String, ByVal NumeroNFeFinal As String) As Boolean
  
  If TipoMovimento = 0 Then 'Entrada
    Inutilizar = InutilizarEntrada(Filial)
  Else 'Saída
    Inutilizar = InutilizarSaida(Filial, MotivoInutilizacao, Serie, NumeroNFeInicial, NumeroNFeFinal)
  End If
  
End Function


Private Function CancelarSaida(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal MotivoCancelamento As String) As Boolean
  Dim clcArquivo As New Collection
  Dim strLinha As String
  On Error GoTo ErrHandler
  Dim strSQL As String
  Dim strPastaEnvio As String
  
  Dim rsParametros As Recordset
  Dim rsSaidas As Recordset
  Dim rsNFe As Recordset
  
  Dim str_numero_NFe_sem_zeros As String
  Dim str_numero_protocolo_autorizacao_NFe As String
  
  '--------------------------------------------------------------------------------------------------
  'Dados
  '--------------------------------------------------------------------------------------------------
  '
  'Parâmetros da Filial
  strSQL = "SELECT *  "
  strSQL = strSQL & "FROM [Parâmetros Filial] "
  strSQL = strSQL & "WHERE Filial = " & Filial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsParametros
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & Filial, vbExclamation, "Atenção"
      .Close
      Set rsParametros = Nothing
      Exit Function
    Else
      'Validações
      strPastaEnvio = Trim(.Fields("PastaEnvioNfe").Value)
      If strPastaEnvio = "" Then
        MsgBox "Pasta de envio não configurada para a Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsParametros = Nothing
        Exit Function
      Else
        If Right(strPastaEnvio, 1) <> "\" Then
          strPastaEnvio = strPastaEnvio & "\"
        End If
      End If
    End If
  End With
  '
  'Saída
  strSQL = "SELECT [Nota Impressa] "
  strSQL = strSQL & "FROM Saídas "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequência = " & Sequencia
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rsSaidas
    If (.BOF And .EOF) Then
      MsgBox "Saída não localizada, sequência: " & Sequencia, vbExclamation, "Atenção"
      .Close
      Set rsSaidas = Nothing
      Exit Function
    Else
      'Validações
      str_numero_NFe_sem_zeros = Trim(.Fields("Nota Impressa").Value)
      If str_numero_NFe_sem_zeros = "" Then
        MsgBox "Nota Fiscal não enviada, não é possível cancelar. Favor verificar, sequência: " & Sequencia, vbExclamation, "Atenção"
        .Close
        Set rsSaidas = Nothing
        Exit Function
      End If
    End If
  End With
  '------------------------------------------------------------------------------
  
  'NFe
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM NFe "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequencia = " & Sequencia & " AND TipoMovimento = 1"
  Set rsNFe = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rsNFe
    If (.BOF And .EOF) Then
      MsgBox "NFe não localizada, sequência: " & Sequencia, vbExclamation, "Atenção"
      .Close
      Set rsNFe = Nothing
      Exit Function
    Else
      'Validações
      str_numero_protocolo_autorizacao_NFe = Trim(.Fields("ProtocoloAutorizacao").Value)
      If str_numero_protocolo_autorizacao_NFe = "" Then
        MsgBox "Retorno de envio da Nota Fiscal não recebido, não existe ainda número do Protocolo de Autorização para poder pedir o cancelamento. Favor verificar, sequência: " & Sequencia, vbExclamation, "Atenção"
        .Close
        Set rsNFe = Nothing
        Exit Function
      End If
    End If
  End With
  
  'Pegar a chave de acesso da NFe na tbl NFe
  
  '--------------------------------------------------------------------------------
  'Estrutura da Nota Fical
  '-------------------------------------------------------------------------------------------------------------
  '1ª linha [00069957000194;1;cancNFe;34527]
  '--------------------------------------------------------------------------------------------------
  ' Cabeçalho
  '--------------------------------------------------------------------------------------------------
  If rsParametros.Fields("PadraoArquivoIntegracao") = "TXT" Then
   strLinha = FormataCnpjCpf(rsParametros.Fields("CGC").Value & "") & ";" 'CNPJ da Filial
   strLinha = strLinha & "1;" 'Código fixo
   strLinha = strLinha & "cancNFe;" 'Tipo de operação
   strLinha = strLinha & str_numero_NFe_sem_zeros 'Número do Documento Fiscal
   AddLinha clcArquivo, strLinha
   '--------------------------------------------------------------------------------------------------
   '2ªlinha [0000;1.07;cancNFe]
   '--------------------------------------------------------------------------------------------------
   ' A- Dados da Nota Fiscal eletrônica
   '--------------------------------------------------------------------------------------------------
   strLinha = "0000;" 'Registro
   strLinha = strLinha & VERSAO_LAYOUT_CANCELAMENTO & ";" 'Versão
   strLinha = strLinha & "cancNFe" 'Tipo de operação
   AddLinha clcArquivo, strLinha
   '--------------------------------------------------------------------------------------------------
   '3ªlinha [1000;ID41090500069957000194550010000320040000000012]
   '-------------------------------------------------------------------------------------------------------------
   strLinha = "1000;ID" 'Registro
   ' Chave Numérica 44, podendo conter zeros à esquerda, iniciando com o literal ID. Chave de Acesso da Nfe composta por Código da UF + AAMM da emissão + CNPJ do Emitente + Modelo, Série e Número da Nfe + Código Numérico + DV
   strLinha = strLinha & rsNFe.Fields("ChaveAcesso").Value 'Chave de acesso
   AddLinha clcArquivo, strLinha
   '--------------------------------------------------------------------------------------------------
  
   '-------------------------------------------------------------------------------------------------------------
   '4ªlinha [1100;1;CANCELAR;41090500069957000194550010000320040000000012;123400000987567;CANCELAMENTO DE NOTA FISCAL;]
   strLinha = "1100;" 'Registro
   'produção/homologação
   strLinha = strLinha & rsParametros.Fields("AmbienteNfe").Value 'Ambiente da NFe (Produção ou Homologação)
   strLinha = strLinha & ";CANCELAR;"
   'Chave Numérica 44,
   strLinha = strLinha & "ID"
   strLinha = strLinha & rsNFe.Fields("ChaveAcesso").Value 'Chave de acesso
   strLinha = strLinha & ";"
   strLinha = strLinha & rsNFe.Fields("ProtocoloAutorizacao").Value 'Protocolo de Autorização
   strLinha = strLinha & ";"
   'Motivo do cancelamento
   strLinha = strLinha & MotivoCancelamento
   AddLinha clcArquivo, strLinha
  
  ElseIf rsParametros.Fields("PadraoArquivoIntegracao") = "XML" Then
    
    strLinha = "<evento versao=""" & rsParametros.Fields("VersaoLayoutEnvio") & """ xmlns=""http://www.portalfiscal.inf.br/nfe"">" & vbCrLf
    strLinha = strLinha & "  <infEvento Id=""ID110111" & rsNFe.Fields("ChaveAcesso").Value & """>" & vbCrLf
    strLinha = strLinha & "    <cOrgao>" & GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") & "</cOrgao>" & vbCrLf
    strLinha = strLinha & "    <tpAmb>" & rsParametros.Fields("AmbienteNFe").Value & "</tpAmb>" & vbCrLf
    strLinha = strLinha & "    <CNPJ>" & FormataCnpjCpf(rsParametros.Fields("CGC")) & "</CNPJ>" & vbCrLf
    strLinha = strLinha & "    <chNFe>" & rsNFe.Fields("ChaveAcesso").Value & "</chNFe>" & vbCrLf
    strLinha = strLinha & "    <dhEvento>" & Now & "</dhEvento>" & vbCrLf
    strLinha = strLinha & "    <tpEvento>110111</tpEvento>" & vbCrLf
    strLinha = strLinha & "    <nSeqEvento>" & rsNFe.Fields("nrEvento") + 1 & "</nSeqEvento>" & vbCrLf
    strLinha = strLinha & "    <verEvento>" & rsParametros.Fields("VersaoLayoutEnvio") & "</verEvento>" & vbCrLf
    strLinha = strLinha & "    <detEvento versao=""" & rsParametros.Fields("VersaoLayoutEnvio") & """>" & vbCrLf
    strLinha = strLinha & "     <descEvento>Cancelamento</descEvento>" & vbCrLf
    strLinha = strLinha & "     <nProt>" & rsNFe.Fields("ProtocoloAutorizacao").Value & "</nProt>" & vbCrLf
    strLinha = strLinha & "     <xJust>" & MotivoCancelamento & "</xJust>" & vbCrLf
    strLinha = strLinha & "    </detEvento>" & vbCrLf
    strLinha = strLinha & "   </infEvento>" & vbCrLf
    strLinha = strLinha & "</evento>"
    AddLinha clcArquivo, strLinha
    rsNFe.Edit
    rsNFe!nrEvento = rsNFe.Fields("nrEvento") + 1
    rsNFe.Update
  End If
  
  '------------------------------------------------------------------------------
 
  'Grava arquivo
  If rsParametros.Fields("PadraoArquivoIntegracao") = "TXT" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeCancSaida" & str_numero_NFe_sem_zeros & ".txt"
  ElseIf rsParametros.Fields("PadraoArquivoIntegracao") = "XML" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeCancSaida" & str_numero_NFe_sem_zeros & ".xml"
  ElseIf rsParametros.Fields("PadraoArquivoIntegracao") = "TxtERP" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeCancSaida" & str_numero_NFe_sem_zeros & ".txt"
  End If
 
  'Retorno da função
  CancelarSaida = True
 
  MsgBox "Nota Fiscal Eletrônica de Saída cancelada com sucesso!", vbInformation, "Atenção"
  
  Exit Function

ErrHandler:

  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Private Function CancelarEntrada(ByVal Filial As Integer, ByVal Sequencia As Long) As Boolean
  Dim clcArquivo As New Collection
  Dim strLinha As String
  Dim strPastaEnvio As String
  On Error GoTo ErrHandler
  
  'Grava arquivo
  GravarArquivo clcArquivo, strPastaEnvio & "NFeCancEntrada" & Sequencia & ".txt"
  
  'Retorno da função
  CancelarEntrada = True
 
  MsgBox "Nota Fiscal Eletrônica de Entrada cancelada com sucesso!", vbInformation, "Atenção"
 
  Exit Function

ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function


Private Function trataCaracteresEspeciaisNFE(ByVal sDado As String) As String
    sDado = Replace(sDado, "ã", "a")
    sDado = Replace(sDado, "à", "a")
    sDado = Replace(sDado, "á", "a")
    sDado = Replace(sDado, "â", "a")
    sDado = Replace(sDado, "é", "e")
    sDado = Replace(sDado, "è", "e")
    sDado = Replace(sDado, "ê", "e")
    sDado = Replace(sDado, "í", "i")
    sDado = Replace(sDado, "ì", "i")
    sDado = Replace(sDado, "ó", "o")
    sDado = Replace(sDado, "ò", "o")
    sDado = Replace(sDado, "ô", "o")
    sDado = Replace(sDado, "ú", "u")
    sDado = Replace(sDado, "ù", "u")
    sDado = Replace(sDado, "ü", "u")
   
    sDado = Replace(sDado, "Ã", "A")
    sDado = Replace(sDado, "À", "A")
    sDado = Replace(sDado, "Á", "A")
    sDado = Replace(sDado, "Â", "A")
    sDado = Replace(sDado, "É", "E")
    sDado = Replace(sDado, "È", "E")
    sDado = Replace(sDado, "Ê", "E")
    sDado = Replace(sDado, "Í", "I")
    sDado = Replace(sDado, "Ì", "I")
    sDado = Replace(sDado, "Ó", "O")
    sDado = Replace(sDado, "Ò", "O")
    sDado = Replace(sDado, "Ô", "O")
    sDado = Replace(sDado, "Ú", "U")
    sDado = Replace(sDado, "Ù", "U")
    sDado = Replace(sDado, "Ü", "U")
    
    sDado = Replace(sDado, "ç", "c")
    sDado = Replace(sDado, "Ç", "C")
   
    sDado = Replace(sDado, "&", "")
    sDado = Replace(sDado, "ª", "")
    sDado = Replace(sDado, "º", "")
    
    sDado = Replace(sDado, ">", "")
    sDado = Replace(sDado, "<", "")
    
    sDado = Replace(sDado, vbCrLf, "")
    
    sDado = LTrim(RTrim(sDado))
    
    trataCaracteresEspeciaisNFE = sDado
  
End Function

' Pilatti Outubro/17
Private Function func_trataCaracteresEspeciaisNFE(ByVal clcArquivo As Collection, ByRef clcArquivoTratado As Collection)
  Dim varLinha As Variant
  Dim strSaida As String
  
  For Each varLinha In clcArquivo
    strSaida = CStr(varLinha)
    strSaida = Replace(strSaida, "ã", "a")
    strSaida = Replace(strSaida, "à", "a")
    strSaida = Replace(strSaida, "á", "a")
    strSaida = Replace(strSaida, "â", "a")
    strSaida = Replace(strSaida, "é", "e")
    strSaida = Replace(strSaida, "è", "e")
    strSaida = Replace(strSaida, "ê", "e")
    strSaida = Replace(strSaida, "í", "i")
    strSaida = Replace(strSaida, "ì", "i")
    strSaida = Replace(strSaida, "ó", "o")
    strSaida = Replace(strSaida, "ò", "o")
    strSaida = Replace(strSaida, "ô", "o")
    strSaida = Replace(strSaida, "ú", "u")
    strSaida = Replace(strSaida, "ù", "u")
    strSaida = Replace(strSaida, "ü", "u")
   
    strSaida = Replace(strSaida, "Ã", "A")
    strSaida = Replace(strSaida, "À", "A")
    strSaida = Replace(strSaida, "Á", "A")
    strSaida = Replace(strSaida, "Â", "A")
    strSaida = Replace(strSaida, "É", "E")
    strSaida = Replace(strSaida, "È", "E")
    strSaida = Replace(strSaida, "Ê", "E")
    strSaida = Replace(strSaida, "Í", "I")
    strSaida = Replace(strSaida, "Ì", "I")
    strSaida = Replace(strSaida, "Ó", "O")
    strSaida = Replace(strSaida, "Ò", "O")
    strSaida = Replace(strSaida, "Ô", "O")
    strSaida = Replace(strSaida, "Ú", "U")
    strSaida = Replace(strSaida, "Ù", "U")
    strSaida = Replace(strSaida, "Ü", "U")
    
    strSaida = Replace(strSaida, "ç", "c")
    strSaida = Replace(strSaida, "Ç", "C")
   
    strSaida = Replace(strSaida, "&", "")
    strSaida = Replace(strSaida, "ª", "")
    strSaida = Replace(strSaida, "º", "")
   
    clcArquivoTratado.Add strSaida
  Next
  
End Function

Private Function EnviarSaida_4_0_REMODELADA(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal ConcatenarDescricaoAdicional As Boolean, ByVal TpEmis As Integer, ByVal dte_data_inicio_contingencia As Date, ByVal str_hora_inicio_contingencia As String, ByVal str_motivo_contingencia As String, ByVal PadraoArquivoIntegracao As String) As Boolean
  
  Dim bGerarContaPaga As Boolean
  
  Dim clcArquivo As New Collection
  Dim clcArquivoTratadoCaracteres As New Collection
  
  m_versao_layout_envio = "4.00"
  m_bln_tem_desconto = False
  
  Dim strLinha As String
  Dim strChaveNFe As String
  Dim strCodigoAleatorioChaveNFe As String
  Dim intDV As Integer
  Dim intAux As Integer
  
  Dim strSQL As String
  Dim rsParametros As Recordset
  Dim rsOpSaida As Recordset
  Dim rsCliFor As Recordset
  Dim rsCadastroTransportadora As Recordset
  Dim rsTransportadora As Recordset
  Dim rsContasReceber As Recordset
  
  Dim blnShowObs As Boolean
  Dim frmX As Form
  Dim intX As Integer

  Dim strPastaEnvio As String
  Dim lngNotaFiscal As Long
  Dim intSerieNotaFiscal As Integer
  Dim dblValorTotalProdutos As Double
  Dim dblValorTotalServicos As Double
    
  'Incluído calculo de PIS e COFINS usando os percentuais da tabela Parâmetros Filial
  Dim dblPercentualPIS As Double
  Dim dblPercentualCofins As Double
    
  Dim blnSimplesNacional As Boolean
  Dim dblAliquotaSimples As Double
  Dim dblPercentualReducaoSimplesNacional As Double
  
  Dim boTEM_SERVICOS As Boolean
  Dim dbl_vSERVICOS As Double
  
  'Tratamento para peso bruto e líquido
  Dim dblPeso As Double
  
  Dim strEmail As String
  
  Dim boFreteJaSomadoNaVenda As Boolean
  boFreteJaSomadoNaVenda = False
  
  Dim IE As String
  Dim sTrataDado As String
  Dim stotal_vICMSST_ParametrizacaoAUX As String
  Dim bTEM_ALGUM_PAGAMENTO As Boolean
  bTEM_ALGUM_PAGAMENTO = False
  
  stotal_vICMSST_ParametrizacaoAUX = "0,00"
  vDespAcessorias = 0
  
  On Error GoTo ErrHandler
  
  'variaveis partilha icms
  isInterestadual = False
  isConsumidorFinal = False
  isIEIsento = False

  bTEM_ICMS_3179a = False
  
  inst_dblPercentualPIS = 0
  inst_dblPercentualCofins = 0

  v_vlSOMA_DO_ICMS_ST_BaseDeCalculo = 0
  v_vlSOMA_DO_ICMS_ST = 0

  aliquotaICMS_tab_ICMS_PERC_ESTADOS = ""
  bo_AliquotaICMS_interestadual = False
  
  sXML_40 = ""
  sPadraoArquivoIntegracao = PadraoArquivoIntegracao
  '--------------------------------------------------------------------------------------------------
  'Dados
  '--------------------------------------------------------------------------------------------------
  'Parâmetros da Filial
  strSQL = "SELECT CGC, [Razão Social], Nome, Inscrição, Endereço, EnderecoNumero, EnderecoComplemento, "
  strSQL = strSQL & "Bairro, Cidade, Estado, CEP, Pais, Fone, "
  strSQL = strSQL & "AmbienteNfe, FormatoImpressaoDanfeNfe, ModDetBaseCalculoIcms, ModDetBaseCalculoIcmsSt, "
  strSQL = strSQL & "PastaEnvioNfe, PastaRetornoNfe, PIS, COFINS, CodigoRegimeTributario, PercentualSimplesNacional, PercentualReducaoBCSimplesNacional, [Gerar Conta Paga], TipoSituacaoTributariaPIS, [Saida Descr Adicional], InscricaoMunicipal, CNAE "
  strSQL = strSQL & ", EditarNomeProduto "
  strSQL = strSQL & "FROM [Parâmetros Filial] "
  strSQL = strSQL & "WHERE Filial = " & Filial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsParametros
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & Filial, vbExclamation, "Atenção"
      .Close
      Set rsSaidas = Nothing
      Exit Function
    Else
      bGerarContaPaga = .Fields("Gerar Conta Paga").Value
      ConcatenarDescricaoAdicional = .Fields("Saida Descr Adicional").Value

      If Not IsNull(.Fields("TipoSituacaoTributariaPIS").Value) Then
          gTipoSituacaoTributariaPIS = .Fields("TipoSituacaoTributariaPIS").Value
      Else
          gTipoSituacaoTributariaPIS = 0
      End If

      'Validações
      strPastaEnvio = Trim(.Fields("PastaEnvioNfe").Value)
      If strPastaEnvio = "" Then
        MsgBox "Pasta de envio não configurada para a Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsParametros = Nothing
        Exit Function
      Else
        If Right(strPastaEnvio, 1) <> "\" Then
          strPastaEnvio = strPastaEnvio & "\"
        End If
      End If
      
      'Valida código IBGE para UF
      If GetTerritorioUfCodigoIbge(.Fields("Estado").Value & "") = 0 Then
        MsgBox "Código IBGE inválido para a UF da Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsParametros = Nothing
        Exit Function
      End If
      'Valida CNPJ
      If (Trim(.Fields("CGC").Value) & "") = "" Or Not Valid_CGC(.Fields("CGC").Value & "") Then
        MsgBox "CNPJ inválido para a Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsParametros = Nothing
        Exit Function
      End If
      
      ' Pilatti adicionei uma linha
      sEstadoFilial = .Fields("Estado").Value
      
      dblPercentualPIS = (.Fields("PIS").Value / 100)
      dblPercentualCofins = (.Fields("COFINS").Value / 100)
      
      If (.Fields("CodigoRegimeTributario").Value) = 1 Or (.Fields("CodigoRegimeTributario").Value) = 2 Then
        blnSimplesNacional = True
      Else
        blnSimplesNacional = False
      End If
      
      blnSimplesNacionalGeral = blnSimplesNacional
              
      dblAliquotaSimples = .Fields("PercentualSimplesNacional").Value
      
      dblPercentualReducaoSimplesNacional = (1 - (.Fields("PercentualReducaoBCSimplesNacional").Value / 100))
    End If
  End With
  
  ' *******************************************
  ' PILATTI Outubro/17
  ' Tratar horario de verão
  Dim sFuso As String
  sFuso = ""
  func_horarioFuso bHorarioDeVerao, sEstadoFilial, sFuso
  ' *******************************************
  
  'Saída
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM Saídas "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequência = " & Sequencia
  
  'Tirei o ReadOnly pq estava dando erro na hora de gravar o Nro da NFe na tbl de Saidas.
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)
  isConsumidorFinal = rsSaidas.Fields("Consumidor_Final").Value
  
  ' ====================================================================
  ' Além de PRODUTOS, verificar se tem valores de SERVICOS
  If Not IsNull(rsSaidas.Fields("Serviços").Value) Then
      If rsSaidas.Fields("Serviços").Value > 0 Then
          boTEM_SERVICOS = True
          dbl_vSERVICOS = rsSaidas.Fields("Serviços").Value
      Else
          boTEM_SERVICOS = False
          dbl_vSERVICOS = 0
      End If
  Else
      boTEM_SERVICOS = False
      dbl_vSERVICOS = 0
  End If
  ' ====================================================================
  
  'Cliente
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM [Cli_For] "
  strSQL = strSQL & "WHERE Código = " & rsSaidas.Fields("Cliente").Value
  Set rsCliFor = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsCliFor
    If (.BOF And .EOF) Then
      MsgBox "Cliente não localizado para o código: " & rsSaidas.Fields("Cliente").Value, vbExclamation, "Atenção"
      .Close
      Set rsCliFor = Nothing
      Exit Function
    Else
      'Valida código IBGE para UF
      If GetTerritorioUfCodigoIbge(.Fields("Estado").Value & "") = 0 Then
        MsgBox "Código IBGE inválido para a UF do Cliente: " & rsSaidas.Fields("Cliente").Value, vbExclamation, "Atenção"
        .Close
        Set rsCliFor = Nothing
        Exit Function
      End If
      UF_destino = .Fields("Estado").Value
      
      'Incluído validação de CPF
      'Valida CNPJ
      If .Fields("Física_Jurídica").Value = "J" Then
        If (Trim(.Fields("CGC").Value) & "") = "" Or Not Valid_CGC(.Fields("CGC").Value & "") Then
          MsgBox "CNPJ inválido para o Cliente: " & rsSaidas.Fields("Cliente").Value, vbExclamation, "Atenção"
          .Close
          Set rsCliFor = Nothing
          Exit Function
        End If
      Else
        If (Trim(.Fields("CGC").Value) & "") = "" Or Not Valid_CPF(.Fields("CGC").Value & "") Then
          MsgBox "CPF inválido para o Cliente: " & rsSaidas.Fields("Cliente").Value, vbExclamation, "Atenção"
          .Close
          Set rsCliFor = Nothing
          Exit Function
        End If
      End If
    End If
  End With
  '
  
  '************************
  'Tratamento para desconto na venda
  Dim vColuna_Desconto As Double          ' Coluna da tabela de Saídas
  Dim vColuna_DescontoSubTotal As Double  ' Coluna da tabela de Saídas
  Dim vAuxDesc As Double
  Dim vAuxTotalSaidaProdutos As Double
  bTemDescontoNaVenda = False
  
  vAuxTotalSaida = rsSaidas.Fields("Produtos").Value
  vColuna_DescontoSubTotal = rsSaidas.Fields("DescontoSubTotal").Value
  vColuna_Desconto = rsSaidas.Fields("Desconto").Value

  If vColuna_DescontoSubTotal > 0 Then
      bTemDescontoNaVenda = True
      vTemDescontoNaVendaPercentual = (vColuna_DescontoSubTotal * 100) / vAuxTotalSaida
      
  ElseIf vColuna_Desconto > 0 Then
      Dim rsSaidasProdAux As Recordset
      strSQL = "SELECT SUM([Preço Final]) From [Saídas - Produtos] "
      strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequência = " & Sequencia
      Set rsSaidasProdAux = db.OpenRecordset(strSQL, dbOpenDynaset)
      vAuxTotalSaidaProdutos = rsSaidasProdAux.Fields(0).Value
      rsSaidasProdAux.Close
      Set rsSaidasProdAux = Nothing
      
      vAuxDesc = vAuxTotalSaida - vAuxTotalSaidaProdutos
      Dim sValorAux As String
      sValorAux = FormataValorTexto(vAuxDesc, 2)
      vAuxDesc = Replace(sValorAux, ".", ",")
      
      If vColuna_Desconto <> vAuxDesc Then
        If vAuxDesc = 0 Then
          bTemDescontoNaVenda = True
          vTemDescontoNaVendaPercentual = (vColuna_Desconto * 100) / vAuxTotalSaida
        End If
      End If
  End If
  vSOMA_TOTAL_DESCONTO = vColuna_Desconto + vColuna_DescontoSubTotal

  'Operação de Saída
  'Incluído verificação se é para chamar a tela de observações.
  strSQL = "SELECT Nome, InTelaObsTransp, CSO, [Soma Frete],Tipo, IPI, SomaIcmsRetidoTotalNota, ICM, ObterTributosProduto_EntradaOuSaida, SomaIpiTotalNota, SomaIpiTotalBC "
  strSQL = strSQL & "FROM [Operações Saída] "
  strSQL = strSQL & "WHERE Código = " & rsSaidas.Fields("Operação").Value
  Set rsOpSaida = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsOpSaida
    If (.BOF And .EOF) Then
        MsgBox "Operação não localizado para o código: " & rsSaidas.Fields("Operação").Value, vbExclamation, "Atenção"
        .Close
        Set rsOpSaida = Nothing
        Exit Function
    Else
        blnShowObs = .Fields("InTelaObsTransp").Value
      
        If .Fields("IPI").Value = True Then
            bo_CalcularIPI = True
        Else
            bo_CalcularIPI = False
        End If
      
        '09/03/2023 - Pablo
        If .Fields("SomaIpiTotalNota").Value = True Then
            bo_SomarIPItotal = True
        Else
            bo_SomarIPItotal = False
        End If
      
        '15/03/2023 - Pablo
        If .Fields("SomaIpiTotalBC").Value = True Then
            bo_SomarIPIBC = True
        Else
            bo_SomarIPIBC = False
        End If
        
        If .Fields("ICM").Value = True Then
            bo_CalcularICMS = True
        Else
            bo_CalcularICMS = False
        End If
        
        If .Fields("SomaIcmsRetidoTotalNota").Value = -1 Then
            bo_CalcularICMS_ST = True
        Else
            bo_CalcularICMS_ST = False
        End If
        
        If Not IsNull(.Fields("ObterTributosProduto_EntradaOuSaida").Value) Then
            If .Fields("ObterTributosProduto_EntradaOuSaida").Value = 1 Then
                'Obter Tributos dos produtos de Entrada
                iObterTributosProduto_EntradaOuSaida = 1
            ElseIf .Fields("ObterTributosProduto_EntradaOuSaida").Value = 2 Then
                'Obter Tributos dos produtos de Saida
                iObterTributosProduto_EntradaOuSaida = 2
            Else
                'Obter Tributos dos produtos de Entrada
                iObterTributosProduto_EntradaOuSaida = 1
            End If
        Else
            'Obter Tributos dos produtos de Entrada
            iObterTributosProduto_EntradaOuSaida = 1
        End If
        
        
    End If
  End With
  
  If Not IsNull(rsOpSaida.Fields("Soma Frete").Value) Then
      If rsOpSaida.Fields("Soma Frete").Value = True Then
          boFreteJaSomadoNaVenda = True
      Else
          boFreteJaSomadoNaVenda = False
      End If
  End If
  
  If Not IsNull(rsOpSaida.Fields("Tipo").Value) Then
      sTIPO_OPERACAO = rsOpSaida.Fields("Tipo").Value
  End If
  
  'Chamada da tela de observações (dados dos volumes transportados).
  If blnShowObs Then
    Set frmX = New frmObsNota
    frmX.gsCliente = rsCliFor.Fields("Transportadora").Value
    frmX.lngSequencia = Sequencia
    frmX.bytTipoTabela = 1
    frmX.Show vbModal
    Set frmX = Nothing
    If gsRetornoDoc <> "OK" Then
      StatusMsg "Nota não impressa."
      Exit Function
    End If
  Else
    For intX = 0 To 1
      gsObsDoc(intX) = ""
    Next intX
    gsPlaca = ""
    gsUfrmPlaca = ""
    gsQtdeTrans = ""
    gsMarcaTrans = ""
    gsEspecieTrans = ""
    gsPesoBruto = ""
    gsPesoLiquido = ""
    gsTransportadora = ""
  End If
  
  If IsNull(rsSaidas.Fields("obs_Transportadora").Value) And (rsSaidas.Fields("obs_FretePago").Value <> 9) Then
    'Cadastro da transportadora
    strSQL = "SELECT Nome FROM Transportadoras WHERE Código = " & rsCliFor.Fields("Transportadora").Value
    Set rsCadastroTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    If rsCadastroTransportadora.BOF And rsCadastroTransportadora.EOF Then
      MsgBox "Se optou por despachar o produto pela Transportadora do Cliente, informo que não foi localizada." & vbCrLf & vbCrLf & "Senão, IGNORE ESTE AVISO.", vbInformation, "Atenção"
    End If
  End If
    
  With rsSaidas
    If (.BOF And .EOF) Then
      MsgBox "Movimento não localizado para a Filial: " & Filial & ", sequência: " & Sequencia, vbExclamation, "Atenção"
      .Close
      Set rsSaidas = Nothing
      Exit Function
    Else
      Call IsDataType(dtDouble, .Fields("Produtos").Value, dblValorTotalProdutos)
      Call IsDataType(dtDouble, .Fields("Serviços").Value, dblValorTotalServicos)
      '
      'Número da Nota Fiscal
      Call IsDataType(dtLong, .Fields("Nota Impressa").Value, lngNotaFiscal)
      If lngNotaFiscal = 0 Then
          .Edit
        
          '13/09/2022 - Pablo
          'And Not (rsCadastroTransportadora Is Nothing)
          'Incluído para evitar erro de objeto nulo
          If IsNull(.Fields("obs_Transportadora").Value) And Not (rsCadastroTransportadora Is Nothing) Then
            If rsCadastroTransportadora.RecordCount > 0 Then
              .Fields("obs_Transportadora").Value = rsCadastroTransportadora.Fields("Nome").Value
            Else
              rsCadastroTransportadora.Close
              Set rsCadastroTransportadora = Nothing
            End If
          End If
        
          'Atualizado os campos da tela de observação
          If blnShowObs Then
            .Fields("obs_Placa").Value = gsPlaca
            .Fields("obs_UF").Value = gsUfrmPlaca
            .Fields("obs_Qtde").Value = gsQtdeTrans
            .Fields("obs_Marca").Value = gsMarcaTrans
            .Fields("obs_Especie").Value = gsEspecieTrans
            .Fields("obs_PesoBruto").Value = gsPesoBruto
            .Fields("obs_PesoLiquido").Value = gsPesoLiquido
            .Fields("obs_Transportadora").Value = gsTransportadora
            .Fields("obs_infCpl1").Value = gsObsDoc(0)
            .Fields("obs_infCpl2").Value = gsObsDoc(1)
            .Fields("obs_FretePago").Value = gsFretePago
          End If
          
          lngNotaFiscal = g_lngNextNotaFiscal_e(Filial)
          .Fields("Nota Impressa").Value = lngNotaFiscal
          .Fields("DataEmissaoNota").Value = Now
          
          .Update
      
          'Atualiza a Nota Fiscal e Fatura do Contas a Receber
          strSQL = "UPDATE [Contas a Receber] SET Nota = " & lngNotaFiscal
          strSQL = strSQL & ", Fatura = '" & lngNotaFiscal & "/ ' & Parcela"
          strSQL = strSQL & " WHERE Tipo = 'R'"
          strSQL = strSQL & " AND Filial = " & Filial
          strSQL = strSQL & " AND Sequência = " & Sequencia
      
          db.Execute strSQL, dbFailOnError
      End If
      '
      'Série da Nota Fiscal
      Call IsDataType(dtInteger, .Fields("SerieNF").Value, intSerieNotaFiscal)
      If intSerieNotaFiscal = 0 Then
          'Tratamento NrSerie pelo arquivo .txt
          If gNrSerieNF > 0 Then
              Dim xConta As Integer
              Dim sCGC As String
              sCGC = rsParametros.Fields("CGC").Value
              sCGC = Replace(sCGC, " ", "")
              sCGC = Replace(sCGC, "/", "")
              sCGC = Replace(sCGC, "\", "")
              sCGC = Replace(sCGC, ".", "")
              sCGC = Replace(sCGC, ",", "")
              sCGC = Replace(sCGC, "-", "")
              sCGC = Replace(sCGC, ";", "")
              For xConta = 0 To gNrSerieNF - 1
                  If gArrayNrSerieNF(xConta, 0) = sCGC Then
                      'NrCnpj1 , SerieNFe1, SerieNFCe1
                      intSerieNotaFiscal = gArrayNrSerieNF(xConta, 1)
                      Exit For
                  End If
              Next
              
              If intSerieNotaFiscal = 0 Then
                  MsgBox "Atenção, não foi encontrado o cadastro da <Serie> da NFe...estará saindo com valor '1'", vbCritical, "Atenção"
                  intSerieNotaFiscal = 1
              End If
          Else
              intSerieNotaFiscal = 1
          End If
          'Fim tratamento NrSerie
      End If
    End If
  End With
  
  Dim intTipoEmissao As Integer
  intTipoEmissao = GetTipoEmissao(Filial, Sequencia, 1)
  
  'Chave de Acesso
  strChaveNFe = GetChaveAcesso(Filial, Sequencia, 1)
  If strChaveNFe <> "" Then    'Código Numérico (aleatório)
    'Alterado de 9 para 8 posicoes
    strCodigoAleatorioChaveNFe = Mid(strChaveNFe, 36, 8)
    'Dígito verificador
    intDV = Right(strChaveNFe, 1)
  

    If intTipoEmissao <> TpEmis Then 'Se a forma de emissao que está gravada na NFe for diferente da forma de emissao atual
      If TpEmis <> 1 Then 'Se for emissão em estado de contingência
        'Tem que refazer a chave de acesso, como o tipo de emissao de contingencia
        'Código Numérico (aleatório)
        Randomize
        strCodigoAleatorioChaveNFe = Format(Int((99999999) * Rnd + 1), String(8, "0"))
        'Monta o Identificador da TAG a ser assinada
        'Exemplo (campos separados por espaço): 41 0908 79077343000104 55 001 000027025 000000001 4
        strChaveNFe = GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") 'Código da UF
        strChaveNFe = strChaveNFe & Format(Date, "yyMM") 'AAMM da emissão
        strChaveNFe = strChaveNFe & FormataCnpjCpf(rsParametros.Fields("CGC").Value & "") 'CNPJ do Emitente
        strChaveNFe = strChaveNFe & "55" 'Modelo da NFe
        strChaveNFe = strChaveNFe & Format(intSerieNotaFiscal, String(3, "0")) 'Serie da NFe
        strChaveNFe = strChaveNFe & Format(lngNotaFiscal, String(9, "0")) 'Número da NFe
        strChaveNFe = strChaveNFe & TpEmis
        strChaveNFe = strChaveNFe & strCodigoAleatorioChaveNFe 'Código Numérico (aleatório)
        'Dígito verificador
        intDV = CalculaDV(strChaveNFe)
        strChaveNFe = strChaveNFe & intDV
      End If
    End If
    
  Else
    'Código Numérico (aleatório)
    Randomize
    'Alterado de 9 para 8 posicoes
    strCodigoAleatorioChaveNFe = Format(Int((99999999) * Rnd + 1), String(8, "0"))
    'Monta o Identificador da TAG a ser assinada
    'Exemplo (campos separados por espaço): 41 0908 79077343000104 55 001 000027025 000000001 4
    strChaveNFe = GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") 'Código da UF
    
    If bolEnviaDataRetroativa = True Then
        strChaveNFe = strChaveNFe & Mid(sEnviaDataRetroativa, 9, 2) & Mid(sEnviaDataRetroativa, 4, 2)  'AAMM da emissão
    Else
        strChaveNFe = strChaveNFe & Format(Date, "yyMM") 'AAMM da emissão
    End If
    
    strChaveNFe = strChaveNFe & FormataCnpjCpf(rsParametros.Fields("CGC").Value & "") 'CNPJ do Emitente
    strChaveNFe = strChaveNFe & "55" 'Modelo da NFe
    strChaveNFe = strChaveNFe & Format(intSerieNotaFiscal, String(3, "0")) 'Serie da NFe
    strChaveNFe = strChaveNFe & Format(lngNotaFiscal, String(9, "0")) 'Número da NFe
    
    'Outra importante alteração esta na construção da chave de acesso, que agora passa a ter o tpEmis como componente.
    'A construção da chave agora é obtida através da concatenação dos seguintes campos:
    '•  Código UF: 2 posições
    '•  AAMM (Ano e mês): 4 posições
    '•  CNPJ Emitente: 14 posições (precedidos de 0 caso o número não atinja 14 dígitos)
    '•  Modelo: 2 posições
    '•  Série: 3 posições
    '•  Número NF-e: 9 dígitos
    '•  Forma de emissão da NF-e tpEmis: 1 digito
    '1 – Normal – emissão normal;
    '2 – Contingência FS – emissão em contingência com impressão do DANFE em Formulário de Segurança;
    '3 – Contingência SCAN – emissão em contingência no Sistema de Contingência do Ambiente Nacional – SCAN;
    '4 – Contingência DPEC - emissão em contingência com envio da Declaração Prévia de Emissão em Contingência – DPEC;
    '5 – Contingência FS-DA - emissão em contingência com impressão do DANFE em
    '•  Código Numérico (controle ERP): 8 posições
    '•  Dígito Verificador: 1 posição
    'A alteração esta no campo Código Numérico que perdeu um dígito para o tpEmis.
    
    strChaveNFe = strChaveNFe & TpEmis
    
    strChaveNFe = strChaveNFe & strCodigoAleatorioChaveNFe 'Código Numérico (aleatório)
    'Dígito verificador
    intDV = CalculaDV(strChaveNFe)
    strChaveNFe = strChaveNFe & intDV
  End If
  '20/09/2009 - mpdea
  'Salva registro
  SaveEnvioNFe_4_0 Filial, Sequencia, 1, Now, 0, CByte(rsParametros.Fields("AmbienteNFe").Value), 1, _
    lngNotaFiscal, intSerieNotaFiscal, "55", strChaveNFe, TpEmis
  '--------------------------------------------------------------------------------------------------
  
  
  '--------------------------------------------------------------------------------------------------
  ' Cabeçalho
  '--------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------
  ' A- Dados da Nota Fiscal eletrônica
  '--------------------------------------------------------------------------------------------------
      
  ' -----------------------
  ' XML versão NFe 4.00
  
  Dim slngNotaFiscal As String
  slngNotaFiscal = CStr(lngNotaFiscal)
    
  If Len(slngNotaFiscal) < 15 Then
        If Len(slngNotaFiscal) = 1 Then
            slngNotaFiscal = "00000000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 2 Then
            slngNotaFiscal = "0000000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 3 Then
            slngNotaFiscal = "000000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 4 Then
            slngNotaFiscal = "00000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 5 Then
            slngNotaFiscal = "0000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 6 Then
            slngNotaFiscal = "000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 7 Then
            slngNotaFiscal = "00000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 8 Then
            slngNotaFiscal = "0000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 9 Then
            slngNotaFiscal = "000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 10 Then
            slngNotaFiscal = "00000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 11 Then
            slngNotaFiscal = "0000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 12 Then
            slngNotaFiscal = "000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 13 Then
            slngNotaFiscal = "00" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 14 Then
            slngNotaFiscal = "0" + slngNotaFiscal
        End If
  End If

  strLinha = "<?xml version=""1.0"" encoding=""UTF-8""?>"
  strLinha = strLinha & "<enviNFe versao=""" & VERSAO_PROCESSO_EMISSAO & """ xmlns=""http://www.portalfiscal.inf.br/nfe""><idLote>" + slngNotaFiscal + "</idLote><indSinc>0</indSinc>"
  strLinha = strLinha & "<NFe xmlns=""http://www.portalfiscal.inf.br/nfe"">"
  strLinha = strLinha & "<infNFe versao=""" & VERSAO_PROCESSO_EMISSAO & """ Id=""NFe" & strChaveNFe & """>"
  strLinha = strLinha & "<ide>"
  strLinha = strLinha & "<cUF>" & GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") & "</cUF>"
  strLinha = strLinha & "<cNF>" & strCodigoAleatorioChaveNFe & "</cNF>"
  strLinha = strLinha & "<natOp>" & trataCaracteresEspeciaisNFE(Left(rsOpSaida.Fields("Nome").Value, 60)) & "</natOp>"
  strLinha = strLinha & "<mod>55</mod>"
  strLinha = strLinha & "<serie>" & intSerieNotaFiscal & "</serie>"
  strLinha = strLinha & "<nNF>" & lngNotaFiscal & "</nNF>"

  If bolEnviaDataRetroativa = True Then
      sEnviaDataRetroativa = Mid(sEnviaDataRetroativa, 7, 4) & "-" & Mid(sEnviaDataRetroativa, 4, 2) & "-" & Mid(sEnviaDataRetroativa, 1, 2) & "T00:00:00-03:00"
      strLinha = strLinha & "<dhEmi>" & sEnviaDataRetroativa & "</dhEmi>"
      bolEnviaDataRetroativa = False
      sEnviaDataRetroativa = ""
  Else
      strLinha = strLinha & "<dhEmi>" & Format(rsSaidas.Fields("DataEmissaoNota").Value, "yyyy-MM-ddTHH:mm:ss") & sFuso & "</dhEmi>"
  End If

  strLinha = strLinha & "<dhSaiEnt>" & Format(rsSaidas.Fields("Data").Value, "yyyy-MM-dd") & "T" & Format(rsSaidas.Fields("DataEmissaoNota").Value, "hh:mm:ss") & sFuso & "</dhSaiEnt>"
  strLinha = strLinha & "<tpNF>1</tpNF>"
      
  If rsParametros.Fields("Estado").Value = rsCliFor.Fields("Estado").Value Then
    strLinha = strLinha & "<idDest>1</idDest>"
  ElseIf rsParametros.Fields("Estado").Value = "EX" Then
    strLinha = strLinha & "<idDest>3</idDest>"
  Else
    strLinha = strLinha & "<idDest>2</idDest>"
    isInterestadual = True
  End If

  ' Apenas verificar se é uma NFe de Devolução
  If rsSaidas.Fields("FinalidadeNFe").Value = 4 Or sTIPO_OPERACAO = "G" Then  ' NFe de Devolução
      
      ' Tipo Operação de Saída = G (Gratis saida / devolução / Remessa)
      '                     Será uma NFe de:
      '                                        Finalidade = 4 e TpOper = G -> Devolução de mercadoria ao Forn ou
      '                                        Finalidade = 1 e TpOper = G -> Remessa de mercadoria para conserto ao seu Forn
      '                                        Finalidade = 1 e TpOper = G -> Remessa de mercadoria para exposição
      '
      ' Pegar os tributos (ex IMCS de entrada para destaque)
      
      bFINALIDADE_NFE_DEVOLUCAO = True
  Else
      bFINALIDADE_NFE_DEVOLUCAO = False
  End If
  iFINALIDADE_NFE = rsSaidas.Fields("FinalidadeNFe").Value
    
  strLinha = strLinha & "<cMunFG>" & GetTerritorioMunicipioCodigoIbge(rsParametros.Fields("Estado").Value & "", rsParametros.Fields("Cidade").Value & "") & "</cMunFG>"
  strLinha = strLinha & "<tpImp>" & rsParametros.Fields("FormatoImpressaoDanfeNfe").Value & "</tpImp>"
  strLinha = strLinha & "<tpEmis>1</tpEmis>"
  strLinha = strLinha & "<cDV>" & intDV & "</cDV>"
  strLinha = strLinha & "<tpAmb>" & rsParametros.Fields("AmbienteNFe").Value & "</tpAmb>"
  
  If rsSaidas.Fields("FinalidadeNFe").Value > 0 Then
    strLinha = strLinha & "<finNFe>" & rsSaidas.Fields("FinalidadeNFe").Value & "</finNFe>"
  Else
    strLinha = strLinha & "<finNFe>1</finNFe>"
  End If
    
  If IsNull(rsSaidas.Fields("Consumidor_Final").Value) Then
    strLinha = strLinha & "<indFinal>0</indFinal>"
  Else
    strLinha = strLinha & "<indFinal>" & rsSaidas.Fields("Consumidor_Final").Value & "</indFinal>"
  End If
  If IsNull(rsSaidas.Fields("Presenca_Comprador").Value) Then
     strLinha = strLinha & "<indPres>1</indPres>"
  Else
     strLinha = strLinha & "<indPres>" & rsSaidas.Fields("Presenca_Comprador").Value & "</indPres>"
  End If
  
  ' =====================================================================
  ' NT2020.006 v1.10 para produção em 2021
  ' Tratamento tag: indIntermed (Indicador de intermediador/marketplace)
  '                 0=Operação sem intermediador (em site ou plataforma própria)
  '                 1=Operação em site ou plataforma de terceiros (intermediadores/marketplace)
  If rsParametros.Fields("Estado").Value = "SP" Or rsParametros.Fields("Estado").Value = "MG" Or _
      rsParametros.Fields("Estado").Value = "GO" Or rsParametros.Fields("Estado").Value = "MA" Then
    
      If Not IsNull(rsSaidas.Fields("Presenca_Comprador").Value) Then
          If rsSaidas.Fields("Presenca_Comprador").Value = 2 Or rsSaidas.Fields("Presenca_Comprador").Value = 3 Or _
              rsSaidas.Fields("Presenca_Comprador").Value = 4 Or rsSaidas.Fields("Presenca_Comprador").Value = 9 Then

              strLinha = strLinha & "<indIntermed>0</indIntermed>"
          End If
      Else
          strLinha = strLinha & "<indIntermed>0</indIntermed>"
      End If
  End If
  ' =====================================================================
      
      
  strLinha = strLinha & "<procEmi>0</procEmi>"
  strLinha = strLinha & "<verProc>" & VERSAO_PROCESSO_EMISSAO & "</verProc>"
  If TpEmis <> "1" Then 'Se for <> 1, é emissão em regime de contingência
    strLinha = strLinha & "<dhCont>" & Format(dte_data_inicio_contingencia, "yyyy-MM-dd") & "T" & Format(str_hora_inicio_contingencia, "hh:mm:ss") & "</dhCont>"
    strLinha = strLinha & "<xJust>" & str_motivo_contingencia & "</xJust>"
  End If
  
  '==========================================================
  Dim rsChaves      As Recordset
  Dim nContaChaves  As Integer
  Dim sSQLChaves    As String
  sSQLChaves = "Select * from SaidasChaves Where Filial = " & Filial & " and Sequencia = " & Sequencia
  Set rsChaves = db.OpenRecordset(sSQLChaves, dbOpenSnapshot)
  If Not (rsChaves.BOF And rsChaves.EOF) Then
      rsChaves.MoveLast
      rsChaves.MoveFirst
      For nContaChaves = 0 To rsChaves.RecordCount - 1
        strLinha = strLinha & "<NFref>"
        strLinha = strLinha & "<refNFe>" & rsChaves.Fields("Chave").Value & "</refNFe>"
        strLinha = strLinha & "</NFref>"

        rsChaves.MoveNext
      Next
  End If
  rsChaves.Close
  Set rsChaves = Nothing
  '==========================================================

  strLinha = strLinha & "</ide>"
  AddLinha clcArquivo, strLinha

   
  '--------------------------------------------------------------------------------------------------
  'C - Identificação do Emitente da Nota Fiscal Eletrônica
  '--------------------------------------------------------------------------------------------------
  With rsParametros
      ' -----------------------------
      ' XML NFe 4.00
      
      sCNPJ_EmitenteXML = FormataCnpjCpf(.Fields("CGC").Value & "")
      
      strLinha = "<emit>"
      strLinha = strLinha & "<CNPJ>" & FormataCnpjCpf(.Fields("CGC").Value & "") & "</CNPJ>"
      
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Razão Social").Value, 60))
      strLinha = strLinha & "<xNome>" & sTrataDado & "</xNome>"
      
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Nome").Value, 60))
      strLinha = strLinha & "<xFant>" & sTrataDado & "</xFant>"
      strLinha = strLinha & "<enderEmit>"
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Endereço").Value, 60))
      strLinha = strLinha & "<xLgr>" & LTrim(RTrim(sTrataDado)) & "</xLgr>"
      
      If .Fields("EnderecoNumero").Value = "" Then
        strLinha = strLinha & "<nro>S/N</nro>"
      Else
        strLinha = strLinha & "<nro>" & Left(.Fields("EnderecoNumero").Value, 60) & "</nro>"
      End If
      
      If Left(.Fields("EnderecoComplemento").Value, 60) <> "" Then
        sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("EnderecoComplemento").Value, 60))
        strLinha = strLinha & "<xCpl>" & sTrataDado & "</xCpl>"
      End If
      
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Bairro").Value, 60))
      strLinha = strLinha & "<xBairro>" & sTrataDado & "</xBairro>"
      
      If UCase(.Fields("Estado").Value) = "EX" Then
        strLinha = strLinha & "<cMun>" & "9999999" & "</cMun>"
        strLinha = strLinha & "<xMun>" & "EXTERIOR" & "</xMun>"
        strLinha = strLinha & "<UF>" & "EX" & "</UF>"
      Else
        strLinha = strLinha & "<cMun>" & GetTerritorioMunicipioCodigoIbge(.Fields("Estado").Value, .Fields("Cidade").Value & "") & "</cMun>"
        strLinha = strLinha & "<xMun>" & trataCaracteresEspeciaisNFE(Left(.Fields("Cidade").Value, 60)) & "</xMun>"
        strLinha = strLinha & "<UF>" & trataCaracteresEspeciaisNFE(.Fields("Estado").Value) & "</UF>"
      End If
      
      ' Pilatti nfe 4.0 2018
      str_FCP_UF_emitente = .Fields("Estado").Value 'Sigla da UF
      
      strLinha = strLinha & "<CEP>" & FormataNumero(.Fields("CEP").Value & "", 8) & "</CEP>"
      
      If Len(.Fields("Pais").Value) > 0 Then
        strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(.Fields("Pais").Value & ""), 4) & "</cPais>"
        strLinha = strLinha & "<xPais>" & Trim(Left(.Fields("Pais").Value, 60)) & "</xPais>"
      End If
      
      If Left(RetornaSomenteNumero(.Fields("Fone").Value), 10) <> "" Then
        strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(.Fields("Fone").Value), 10) & "</fone>"
      End If
      
      strLinha = strLinha & "</enderEmit>"
     
      strLinha = strLinha & "<IE>" & Left(FormataIE(.Fields("Inscrição").Value, False), 14) & "</IE>"
        
      strLinha = strLinha & "<CRT>" & rsParametros.Fields("CodigoRegimeTributario").Value & "</CRT>"
      strLinha = strLinha & "</emit>"
      AddLinha clcArquivo, strLinha

      strSQL = ""
  End With
  '--------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------
  'E - Identificação do Destinatário da Nota Fiscal eletrônica
  '--------------------------------------------------------------------------------------------------
  Dim strEstadoCliente As String
  With rsCliFor
      ' -------------------------
      ' XML NFe 4.00
    
      strLinha = strLinha & "<dest>"
      If (.Fields("Física_Jurídica").Value) = "F" Then
        strLinha = strLinha & "<CPF>" & FormataCnpjCpf(.Fields("CGC").Value & "") & "</CPF>"
      Else
        strLinha = strLinha & "<CNPJ>" & FormataCnpjCpf(.Fields("CGC").Value & "") & "</CNPJ>"
      End If
      
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Nome").Value, 60))
      strLinha = strLinha & "<xNome>" & LTrim(RTrim(sTrataDado)) & "</xNome>"
      
      strLinha = strLinha & "<enderDest>"
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Endereço").Value, 60))
      strLinha = strLinha & "<xLgr>" & LTrim(RTrim(sTrataDado)) & "</xLgr>"
      
      If .Fields("Endereço Número").Value = "" Then
        strLinha = strLinha & "<nro>S/N</nro>"
      Else
        strLinha = strLinha & "<nro>" & .Fields("Endereço Número").Value & "</nro>"
      End If
      
      If Len(.Fields("Complemento").Value) <> 0 Then
        sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Complemento").Value, 60))
        strLinha = strLinha & "<xCpl>" & RTrim(LTrim(sTrataDado)) & "</xCpl>"
      End If
      
      If Left(.Fields("Bairro").Value, 60) = "" Then
        strLinha = strLinha & "<xBairro>" & "CENTRO" & "</xBairro>"
      Else
        sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Bairro").Value, 60))
        strLinha = strLinha & "<xBairro>" & LTrim(RTrim(sTrataDado)) & "</xBairro>"
      End If
      
      If UCase(.Fields("Estado").Value) = "EX" Then
        strLinha = strLinha & "<cMun>" & "9999999" & "</cMun>"
        strLinha = strLinha & "<xMun>" & "EXTERIOR" & "</xMun>"
        strLinha = strLinha & "<UF>" & "EX" & "</UF>"
        str_FCP_UF_destinatario = "EX"
      Else
        strLinha = strLinha & "<cMun>" & GetTerritorioMunicipioCodigoIbge(.Fields("Estado").Value, .Fields("Cidade").Value & "") & "</cMun>"
        strLinha = strLinha & "<xMun>" & trataCaracteresEspeciaisNFE(Left(.Fields("Cidade").Value, 60)) & "</xMun>"
        strLinha = strLinha & "<UF>" & trataCaracteresEspeciaisNFE(.Fields("Estado").Value) & "</UF>"
        str_FCP_UF_destinatario = .Fields("Estado").Value
      End If
      
      If Len(.Fields("CEP").Value) > 0 Then
        strLinha = strLinha & "<CEP>" & LTrim(RTrim(FormataNumero(.Fields("CEP").Value & "", 8))) & "</CEP>"
      End If
      
      If Len(.Fields("Pais").Value) > 0 Then
        strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(.Fields("Pais").Value & ""), 4) & "</cPais>"
        strLinha = strLinha & "<xPais>" & Trim(Left(.Fields("Pais").Value, 60)) & "</xPais>"
      End If
      
      
      ' Mauro - 11/08/2021 Aumentando o tamanho do campo Fone
      
      If Not IsNull(.Fields("Fone 1").Value) Then
          If Left(RetornaSomenteNumero(.Fields("Fone 1").Value), 12) <> "" Then
            strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(.Fields("Fone 1").Value), 12) & "</fone>"
          End If
      End If
      
      strLinha = strLinha & "</enderDest>"
      
      IE = Left(FormataIE(.Fields("Inscrição").Value, UCase(.Fields("Física_Jurídica").Value) = "F"), 14)
      IndicadorIE = .Fields("IndicadorIE").Value
      If IndicadorIE <> 1 And IndicadorIE <> 2 And IndicadorIE <> 9 Then
          IndicadorIE = 9
      End If
      If IndicadorIE = 2 Then
          isIEIsento = True
      End If
      
      ' Pilatti nfe 4.0 2018
      str_FCP_RG_INS_EST_destinatario = IE
      
      strLinha = strLinha & "<indIEDest>" & IndicadorIE & "</indIEDest>"
      
      If IndicadorIE <> "2" Then
        sTrataDado = Left(FormataIE(.Fields("Inscrição").Value, UCase(.Fields("Física_Jurídica").Value) = "F"), 14)
        sTrataDado = trataCaracteresEspeciaisNFE(sTrataDado)
        If (LTrim(RTrim(sTrataDado)) <> "") Then strLinha = strLinha & "<IE>" & LTrim(RTrim(sTrataDado)) & "</IE>" 'IE
      End If
      
      
      If Len(.Fields("email").Value) <> 0 Then
        sTrataDado = trataCaracteresEspeciaisNFE(.Fields("email").Value)
        strLinha = strLinha & "<email>" & LTrim(RTrim(sTrataDado)) & "</email>"
      End If
  
      strLinha = strLinha & "</dest>"
        
      'Adicionar linha
      AddLinha clcArquivo, strLinha
      
      '--------------------------------------------------------------------------------------------------
      ' TRATAMENTO PARA DESTAQUE DE ICMS PARA EMPRESAS <LUCRO REAL>....'NÃO SIMPLES'
      '--------------------------------------------------------------------------------------------------
      If blnSimplesNacional = False Then
          Dim rsEstadosICMS As Recordset
          aliquotaICMS_tab_ICMS_PERC_ESTADOS = ""
    
          Set rsEstadosICMS = db.OpenRecordset("ICMS_Percentual_Estados", dbOpenTable, dbReadOnly)
    
          If Not (rsEstadosICMS.EOF And rsEstadosICMS.BOF) Then
              rsEstadosICMS.MoveFirst
              While Not rsEstadosICMS.EOF
                  If UCase(rsEstadosICMS.Fields("ESTADO_ORIGEM").Value) = UCase(gsEstadoOrigemEmpresaLogado) And _
                    UCase(rsEstadosICMS.Fields("ESTADO_DESTINO").Value) = UCase(.Fields("Estado").Value) Then
                      aliquotaICMS_tab_ICMS_PERC_ESTADOS = rsEstadosICMS.Fields("ALIQUOTA").Value
                      rsEstadosICMS.MoveLast
                  End If
                  rsEstadosICMS.MoveNext
              Wend
          End If
          rsEstadosICMS.Close
          Set rsEstadosICMS = Nothing
          
          If UCase(gsEstadoOrigemEmpresaLogado) = UCase(.Fields("Estado").Value) Then
              bo_AliquotaICMS_interestadual = False
          Else
              bo_AliquotaICMS_interestadual = True
          End If
      End If
      '--------------------------------------------------------------------------------------------------
  
      .Close
      strSQL = ""
  End With
  
  'TODO: inserir CNPF contador
  If FormataCnpjCpf(rsParametros.Fields("CGC").Value) = "10401935000567" Then
    AddLinha clcArquivo, "<autXML><CNPJ>38881249000126</CNPJ></autXML>"
  End If
  
  ' *****************************************************
  'Inicio FCP (cabecalho calculo) Pilatti

  Dim ds_auxiliarSQL As Recordset
  
  vFCP_registroTotal = 0
  vFCPST_registroTotal = 0
  vFCPSTRet_registroTotal = 0
  iAnoFCP = Year(Now)

  ' Obter aliquota na tabela geral de FCP_AdesaoPercentual_Quick (0%, 1% ou 2%)
  str_sql = "Select ALIQUOTA1 From FCP_AdesaoPercentual_Quick where ESTADO_ORIGEM = '" & str_FCP_UF_destinatario & "'"
  
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaAdesaoUFDestFCP = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA1").Value)
  End If
  
  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing
  
  ' Obter aliquota na tabela geral de ICMS_PERCENTUAL_ESTADOS (SELECT PARA ESTADOS DIFERENTES)
  str_sql = "Select ALIQUOTA From ICMS_PERCENTUAL_ESTADOS where ESTADO_ORIGEM = '" & str_FCP_UF_emitente & "' and ESTADO_DESTINO = '" & str_FCP_UF_destinatario & "'"
                
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaFCP_EMIT_DEST = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA").Value)
  End If
  
  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing
                
  ' Obter aliquota na tabela geral de ICMS_PERCENTUAL_ESTADOS (SELECT PARA O MESMO ESTADO EMITENTE PR ->  PR)
  str_sql = "Select ALIQUOTA From ICMS_PERCENTUAL_ESTADOS where ESTADO_ORIGEM = '" & str_FCP_UF_emitente & "' and ESTADO_DESTINO = '" & str_FCP_UF_emitente & "'"
                
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaFCP_Emitente = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA").Value)
  End If
  
  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing
                
  ' Obter aliquota na tabela geral de ICMS_PERCENTUAL_ESTADOS (SELECT PARA O MESMO ESTADO DESTINATARIO SP -> SP)
  str_sql = "Select ALIQUOTA From ICMS_PERCENTUAL_ESTADOS where ESTADO_ORIGEM = '" & str_FCP_UF_destinatario & "' and ESTADO_DESTINO = '" & str_FCP_UF_destinatario & "'"
                
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaFCP_Destinatario = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA").Value)
  End If

  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing

  ' Fim FCP Pilatti (cabecalho calculo)
  ' ************************************************************
    
  
  '--------------------------------------------------------------------------------------------------
  'I - Produtos e Serviços da NF-e
  '--------------------------------------------------------------------------------------------------
  iDetNumItemXML = 0
  AddProdutoSevicosSaida_REMODELADA Filial, Sequencia, clcArquivo, rsSaidas, rsParametros, dblPercentualPIS, dblPercentualCofins, ConcatenarDescricaoAdicional, blnSimplesNacional, dblAliquotaSimples, dblPercentualReducaoSimplesNacional, rsOpSaida.Fields("CSO").Value, strEstadoCliente, PadraoArquivoIntegracao
  '--------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------
  'W - Valores Totais da NF-e
  '--------------------------------------------------------------------------------------------------
  
  'Valores Totais da NF-e
  strSQL = "SELECT [Base ICM], [Valor ICM], [Base ICM Subs], [Valor ICM Subs], Produtos, [Serviços], Frete, Seguro, [Recebe - Dinheiro],  "
  strSQL = strSQL & "Desconto, IPI, [Total PIS], [Total COFINS], Total, Frete, Seguro, obs_Qtde, obs_Especie, obs_Marca, obs_Placa, obs_UF,"
  strSQL = strSQL & "obs_PesoLiquido, obs_PesoBruto, obs_infCpl1, obs_infCpl2, TotalNCM, TotalDesoneracaoICMS "
  ' Pilatti Novembro 28-11-2017 linha abaixo FreteSomaOuNaoEstimativa incluida no select
  strSQL = strSQL & " ,FreteSomaOuNaoEstimativa, TotalCartaoDebito, TotalCartaoCredito, [Total Prazo], [Recebe - Vale] "
  strSQL = strSQL & " FROM [Saídas] WHERE Filial = " & Filial
  strSQL = strSQL & "AND Sequência = " & Sequencia
  Set rsSaidas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  Dim str_volumes_quantidade As String

  If PadraoArquivoIntegracao = "XML" Then
    With rsSaidas
      ' -------------------------------
      ' XML
      strLinha = "<total>"
      strLinha = strLinha & "<ICMSTot>"

      If bFINALIDADE_NFE_DEVOLUCAO = True Or bTEM_ICMS_3179a = True Then
      
            '=============================================================================================
            ' Tratamento se Oper.Saída seja classificada como Devolução/Remessa/Saída Gratis
            '=============================================================================================

            If bo_CalcularICMS = False Then
                v_vlSOMA_DO_ICMS_BaseDeCalculo = 0
                vlSOMA_DO_ICMS = 0
            End If
          
            If (.Fields("Base ICM").Value > 0 Or vlSOMA_DO_ICMS > 0) And bo_CalcularICMS = True Then
                strLinha = strLinha & "<vBC>" & FormataValorTexto(v_vlSOMA_DO_ICMS_BaseDeCalculo, 2) & "</vBC>"
            Else
                strLinha = strLinha & "<vBC>0.00</vBC>"
            End If
            
            If bo_CalcularICMS = True Then
                If vlSOMA_DO_ICMS > 0 Then
                    If vlSOMA_DO_ICMS <> .Fields("Valor ICM").Value Then
                      strLinha = strLinha & " <vICMS>" & FormataValorTexto(vlSOMA_DO_ICMS, 2) & "</vICMS>"
                    Else
                      strLinha = strLinha & "<vICMS>" & FormataValorTexto(.Fields("Valor ICM").Value, 2) & "</vICMS>"
                    End If
                Else
                    strLinha = strLinha & "<vICMS>" & FormataValorTexto(.Fields("Valor ICM").Value, 2) & "</vICMS>"
                End If
            Else
                strLinha = strLinha & "<vICMS>0.00</vICMS>"
            End If
            
            If IsNull(.Fields("TotalDesoneracaoICMS").Value) Then
                strLinha = strLinha & "<vICMSDeson>0.00</vICMSDeson>"
            Else
                strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
            End If
            
            ' FCP tratamento NFe 4.0    Pilatti/2018
            ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
            If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacional = False Then
                Dim dAux3 As Double
    
                'vFCPUFDest ------------------------
                'dAux3 = (vFCP_registroTotal * AliquotaFCP_Destinatario) / 100
                'strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(dAux3, 2) & "</vFCPUFDest>"
                strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(vFCPUFDest_registroTotal, 2) & "</vFCPUFDest>"
    
                'vICMSUFDest --------------------------
                ' DIFAL
                'CalculoFCPEstadosDiferentes = (vFCP_registroTotal * AliquotaFCP_EMIT_DEST) / 100
                'CalculoFCPEstadosIguais = (vFCP_registroTotal * AliquotaFCP_Emitente) / 100
                'CalculoFCP_diferenca = CalculoFCPEstadosIguais - CalculoFCPEstadosDiferentes
                'CalculoFCP_participacaoDest = 100
                'strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(CalculoFCP_participacaoDest, 2) & "</vICMSUFDest>"
                strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(vICMSUFDest_registroTotal, 2) & "</vICMSUFDest>"
    
                ' vICMSUFRemet ------------------------
                CalculoFCP_participacaoEmitente = 0
                strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(CalculoFCP_participacaoEmitente, 2) & "</vICMSUFRemet>"
                
                'vFCP
                strLinha = strLinha & "<vFCP>" & FormataValorTexto(vFCP_registroTotal, 2) & "</vFCP>"
            Else
                strLinha = strLinha & "<vFCP>0.00</vFCP>"
            End If
            ' Fim tratamento FCP NFe 4.0
            
            If bo_CalcularICMS_ST = False Then
                v_vlSOMA_DO_ICMS_ST_BaseDeCalculo = 0
                v_vlSOMA_DO_ICMS_ST = 0
            End If
            
            If bo_CalcularICMS_ST = True And v_vlSOMA_DO_ICMS_ST_BaseDeCalculo > 0 And v_vlSOMA_DO_ICMS_ST > 0 Then
                strLinha = strLinha & "<vBCST>" & FormataValorTexto(v_vlSOMA_DO_ICMS_ST_BaseDeCalculo, 2) & "</vBCST>" 'Base de Cálculo do ICMS ST
                strLinha = strLinha & "<vST>" & FormataValorTexto(v_vlSOMA_DO_ICMS_ST, 2) & "</vST>" 'Valor Total do ICMS ST
            ElseIf bo_CalcularICMS_ST = True Then
                strLinha = strLinha & "<vBCST>" & FormataValorTexto(.Fields("Base ICM Subs").Value, 2) & "</vBCST>" 'Base de Cálculo do ICMS ST
                strLinha = strLinha & "<vST>" & FormataValorTexto(.Fields("Valor ICM Subs").Value, 2) & "</vST>" 'Valor Total do ICMS ST
            Else
                strLinha = strLinha & "<vBCST>0.00</vBCST>" 'Base de Cálculo do ICMS ST
                strLinha = strLinha & "<vST>0.00</vST>" 'Valor Total do ICMS ST
            End If
            '=============================================================================================
     
      ElseIf blnSimplesNacional = False Then
        
            '=============================================================================================
            ' EMPRESA Lucro Real ou Presumido
            '=============================================================================================

            If bo_CalcularICMS = False Then
                v_vlSOMA_DO_ICMS_BaseDeCalculo = 0
                vlSOMA_DO_ICMS = 0
            End If

            If vlSOMA_DO_ICMS > 0 And bTEM_ICMS_3179a = True And bo_CalcularICMS = True Then
                strLinha = strLinha & "<vBC>" & FormataValorTexto(s_vlSOMA_DO_ICMS_BaseDeCalculo, 2) & "</vBC>"
      
                If vlSOMA_DO_ICMS <> .Fields("Valor ICM").Value Then
                  strLinha = strLinha & " <vICMS>" & FormataValorTexto(vlSOMA_DO_ICMS, 2) & "</vICMS>"
                Else
                  strLinha = strLinha & " <vICMS>" & FormataValorTexto(.Fields("Valor ICM").Value, 2) & "</vICMS>"
                End If
            Else
                strLinha = strLinha & "<vBC>0.00</vBC>"
                strLinha = strLinha & " <vICMS>0.00</vICMS>"
            End If
            
            If IsNull(.Fields("TotalDesoneracaoICMS").Value) Then
                strLinha = strLinha & "<vICMSDeson>0.00</vICMSDeson>"
            Else
                strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
            End If
            
            ' FCP tratamento NFe 4.0    Pilatti/2018
            ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
            If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                Dim dAux2 As Double
    
                'vFCPUFDest ------------------------
                'dAux2 = (vFCP_registroTotal * AliquotaFCP_Destinatario) / 100
                'strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(dAux2, 2) & "</vFCPUFDest>"
                strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(vFCPUFDest_registroTotal, 2) & "</vFCPUFDest>"
    
                'vICMSUFDest --------------------------
                ' DIFAL
                'CalculoFCPEstadosDiferentes = (vFCP_registroTotal * AliquotaFCP_EMIT_DEST) / 100
                'CalculoFCPEstadosIguais = (vFCP_registroTotal * AliquotaFCP_Emitente) / 100
                'CalculoFCP_diferenca = CalculoFCPEstadosIguais - CalculoFCPEstadosDiferentes
    
                ' Ano 2018 80% para o Destino   e   20% para o Emitente
                ' Ano maior que 2018 100% para o Destino
                'If (iAnoFCP = 2018) Then
                '    CalculoFCP_participacaoDest = (CalculoFCP_diferenca * 80) / 100   ' 80%
                'Else
                '    CalculoFCP_participacaoDest = CalculoFCP_diferenca   ' 100% para 2019 em diante
                'End If
                'strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(CalculoFCP_participacaoDest, 2) & "</vICMSUFDest>"
                strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(vICMSUFDest_registroTotal, 2) & "</vICMSUFDest>"
    
                ' vICMSUFRemet ------------------------
                If (iAnoFCP = 2018) Then
                    CalculoFCP_participacaoEmitente = (CalculoFCP_diferenca * 20) / 100   ' 20%
                Else
                    CalculoFCP_participacaoEmitente = 0   ' 100% para 2019 em diante
                End If
                strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(CalculoFCP_participacaoEmitente, 2) & "</vICMSUFRemet>"
                
                'vFCP
                strLinha = strLinha & "<vFCP>" & FormataValorTexto(vFCP_registroTotal, 2) & "</vFCP>"
            Else
                strLinha = strLinha & "<vFCP>0.00</vFCP>"
            End If
            ' Fim tratamento FCP NFe 4.0
            
            
            If bo_CalcularICMS_ST = True Then
                strLinha = strLinha & "<vBCST>" & FormataValorTexto(.Fields("Base ICM Subs").Value, 2) & "</vBCST>" 'Base de Cálculo do ICMS ST
                strLinha = strLinha & "<vST>" & FormataValorTexto(.Fields("Valor ICM Subs").Value, 2) & "</vST>" 'Valor Total do ICMS ST
            Else
                strLinha = strLinha & "<vBCST>0.00</vBCST>" 'Base de Cálculo do ICMS ST
                strLinha = strLinha & "<vST>0.00</vST>" 'Valor Total do ICMS ST
            End If
            '=============================================================================================
            
      Else
            '=============================================================================================
      
            strLinha = strLinha & "<vBC>0.00</vBC><vICMS>0.00</vICMS>"
            
            If IsNull(.Fields("TotalDesoneracaoICMS").Value) Then
                strLinha = strLinha & "<vICMSDeson>0.00</vICMSDeson>"
            Else
                strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
            End If
            
            ' FCP tratamento NFe 4.0    Pilatti/2018
            ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
            If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacional = False Then ' Or isIEIsento = True) Then
                dAux2 = 0
    
                'vFCPUFDest ------------------------
                'dAux2 = (vFCP_registroTotal * AliquotaFCP_Destinatario) / 100
                'strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(dAux2, 2) & "</vFCPUFDest>"
                strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(vFCPUFDest_registroTotal, 2) & "</vFCPUFDest>"
    
                'vICMSUFDest --------------------------
                ' DIFAL
                'CalculoFCPEstadosDiferentes = (vFCP_registroTotal * AliquotaFCP_EMIT_DEST) / 100
                'CalculoFCPEstadosIguais = (vFCP_registroTotal * AliquotaFCP_Emitente) / 100
                'CalculoFCP_diferenca = CalculoFCPEstadosIguais - CalculoFCPEstadosDiferentes
    
                ' Ano 2018 80% para o Destino   e   20% para o Emitente
                ' Ano maior que 2018 100% para o Destino
                'If (iAnoFCP = 2018) Then
                '    CalculoFCP_participacaoDest = (CalculoFCP_diferenca * 80) / 100   ' 80%
                'Else
                '    CalculoFCP_participacaoDest = CalculoFCP_diferenca   ' 100% para 2019 em diante
                'End If
                'strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(CalculoFCP_participacaoDest, 2) & "</vICMSUFDest>"
                strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(vICMSUFDest_registroTotal, 2) & "</vICMSUFDest>"
    
                ' vICMSUFRemet ------------------------
                If (iAnoFCP = 2018) Then
                    CalculoFCP_participacaoEmitente = (CalculoFCP_diferenca * 20) / 100   ' 20%
                Else
                    CalculoFCP_participacaoEmitente = 0   ' 100% para 2019 em diante
                End If
                strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(CalculoFCP_participacaoEmitente, 2) & "</vICMSUFRemet>"
                
                'vFCP
                strLinha = strLinha & "<vFCP>" & FormataValorTexto(vFCP_registroTotal, 2) & "</vFCP>"
            Else
                strLinha = strLinha & "<vFCP>0.00</vFCP>"
            End If
            ' Fim tratamento FCP NFe 4.0
            
            If bo_CalcularICMS_ST = True Then
                strLinha = strLinha & "<vBCST>" & FormataValorTexto(v_vlSOMA_DO_ICMS_ST_BaseDeCalculo, 2) & "</vBCST>" 'Base de Cálculo do ICMS ST
                strLinha = strLinha & "<vST>" & FormataValorTexto(v_vlSOMA_DO_ICMS_ST, 2) & "</vST>" 'Valor Total do ICMS ST
            Else
                strLinha = strLinha & "<vBCST>0.00</vBCST>" 'Base de Cálculo do ICMS ST
                strLinha = strLinha & "<vST>0.00</vST>" 'Valor Total do ICMS ST
            End If
            '=============================================================================================
      End If
      
      ' FCP tratamento NFe 4.0    Pilatti/2018
      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
      If blnSimplesNacional = False Then
        If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
  
            'vFCPST
            strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vFCPST_registroTotal, 2) & "</vFCPST>"
            'vFCPSTRet
            strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vFCPSTRet_registroTotal, 2) & "</vFCPSTRet>"
        Else
            strLinha = strLinha & "<vFCPST>0.00</vFCPST>"
            strLinha = strLinha & "<vFCPSTRet>0.00</vFCPSTRet>"
        End If
      Else
        strLinha = strLinha & "<vFCPST>0.00</vFCPST>"
        strLinha = strLinha & "<vFCPSTRet>0.00</vFCPSTRet>"
      End If
      
      strLinha = strLinha & "<vProd>" & FormataValorTexto(dblValorTotalProdutos, 2) & "</vProd>"
      strLinha = strLinha & "<vFrete>" & FormataValorTexto(.Fields("Frete").Value, 2) & "</vFrete>"
      strLinha = strLinha & "<vSeg>" & FormataValorTexto(.Fields("Seguro").Value, 2) & "</vSeg>"

      '26/04/2011 - Andrea
      'Alterado para que quando tiver desconto no item, preencha o valor do desconto no total
      'Se não tiver desconto no item (o desconto dado no sub total muda o valor do item, por isso o desconto não aparece nos itens)
      'nao vai ser preenchido o desconto no total.

      If m_bln_tem_desconto = True Then
        
          If bTemDescontoNaVenda = True Then 'Tratamento para DESCONTO DADO DIRETO NA VENDA (Não desconto por produto)
              strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
          Else
              strLinha = strLinha & "<vDesc>" & FormataValorTexto(.Fields("Desconto").Value, 2) & "</vDesc>"
          End If
      Else
          If bTemDescontoNaVenda = True Then
              strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
          Else
              strLinha = strLinha & "<vDesc>0.00</vDesc>"
          End If
      End If
      m_bln_tem_desconto = False
      
      strLinha = strLinha & "<vII>0.00</vII>" '& FormataValorTexto(.Fields("").Value, 2) Valor Total do II (Imposto de Importação) NOT_FOUND
      
      ' NFe 4.0 Pilatti 2018
      If bo_CalcularIPI = True And bFINALIDADE_NFE_DEVOLUCAO = True And nfeDevolucao_impostoDevol = True Then
          strLinha = strLinha & "<vIPI>0.00</vIPI>" 'Valor Total do IPI
          strLinha = strLinha & "<vIPIDevol>" & FormataValorTexto(.Fields("IPI").Value, 2) & "</vIPIDevol>"
      ElseIf bo_CalcularIPI = True Then
          strLinha = strLinha & "<vIPI>" & FormataValorTexto(vlSOMA_DO_IPI, 2) & "</vIPI>" 'Valor Total do IPI
          strLinha = strLinha & "<vIPIDevol>0.00</vIPIDevol>"
      Else
          strLinha = strLinha & "<vIPI>0.00</vIPI>" 'Valor Total do IPI
          strLinha = strLinha & "<vIPIDevol>0.00</vIPIDevol>"
      End If
      
      'Alterado a fonte do valor do PIS e COFINS, calculado pelo percentual da tabela Parametros Filial
      'Porque na tela de Saidas só grava PIS e COFINS para servicos.
      
      If inst_dblPercentualPIS > 0 Then
          strLinha = strLinha & "<vPIS>" & FormataValorTexto(inst_dblPercentualPIS, 2) & "</vPIS>"
          strLinha = strLinha & "<vCOFINS>" & FormataValorTexto(inst_dblPercentualCofins, 2) & "</vCOFINS>"
      Else
          strLinha = strLinha & "<vPIS>0.00</vPIS>"
          strLinha = strLinha & "<vCOFINS>0.00</vCOFINS>"
      End If

'===========================================================================
' 13/04/2023 - PABLO
' CÓDIGO REFEITO LOGO ABAIXO
'===========================================================================
      Dim pab_tmp_vNF As Double
      pab_tmp_vNF = .Fields("Total").Value + .Fields("Seguro").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal
      
      If vDespAcessorias > 0 Then
        strLinha = strLinha & "<vOutro>" & FormataValorTexto(vDespAcessorias, 2) & "</vOutro>"
      
        pab_tmp_vNF = pab_tmp_vNF + vDespAcessorias
        If boFreteJaSomadoNaVenda = True Then
          If boTEM_SERVICOS = True Then
            pab_tmp_vNF = pab_tmp_vNF - dbl_vSERVICOS
          End If
        Else
          pab_tmp_vNF = pab_tmp_vNF + .Fields("Frete").Value
          If boTEM_SERVICOS = True Then
            pab_tmp_vNF = pab_tmp_vNF - dbl_vSERVICOS
          End If
        End If
      Else
        strLinha = strLinha & "<vOutro>0.00</vOutro>"
      
        '09/03/2023 - Pablo
        Dim v_SomaIPItotalNF As Double
        v_SomaIPItotalNF = 0
        If bo_CalcularIPI And bo_SomarIPItotal Then
          v_SomaIPItotalNF = vlSOMA_DO_IPI
        End If
      
        pab_tmp_vNF = pab_tmp_vNF + v_SomaIPItotalNF
        If boFreteJaSomadoNaVenda = True Then
          If boTEM_SERVICOS = True Then
            pab_tmp_vNF = pab_tmp_vNF - dbl_vSERVICOS
          End If
        Else
          pab_tmp_vNF = pab_tmp_vNF + .Fields("Frete").Value
          If boTEM_SERVICOS = True Then
            pab_tmp_vNF = pab_tmp_vNF - dbl_vSERVICOS
          End If
        End If
      End If
      
      If v_vlSOMA_DO_ICMS_ST > 0 And bFINALIDADE_NFE_DEVOLUCAO = False Then
        pab_tmp_vNF = pab_tmp_vNF - vTotal_Desconto_ICMS_vNF
      End If
      
      strLinha = strLinha & "<vNF>" & FormataValorTexto(pab_tmp_vNF, 2) & "</vNF>" 'Valor Total da NF-e
'=======================================================================================================
      
      '01/10/2013 - Jean
      'Adicionado campo conforme mudança do layout do infinitri TotalNCM
      If .Fields("TotalNCM").Value > 0 Then
        strLinha = strLinha & "<vTotTrib>" & FormataValorTexto(vl_vTotTrib, 2) & "</vTotTrib>"
      
      End If
      
      strLinha = strLinha & "</ICMSTot>"
      strLinha = strLinha & "</total>"
      'Adicionar Linha
      AddLinha clcArquivo, strLinha
      strSQL = ""
    End With
  
    strSQL = ""
    
    '--------------------------------------------------------------------------------------------------
    'X-Informações do Transporte da NF-e
    '--------------------------------------------------------------------------------------------------
    strLinha = "<transp>"
    
    '16ªlinha [5000;0]
    ' Modalidade do frete
    ' 0- Contratação do Frete por conta do Remetente (CIF)
    ' 1- Contratação do Frete por conta do destinatário/remetente (FOB)
    ' 2- Contratação do Frete por conta de terceiros
    ' 3- Transporte próprio por conta do remetente
    ' 4- Transporte próprio por conta do destinatário
    ' 9- Sem Ocorrência de transporte
    If gsFretePago = "" Then
      strLinha = strLinha & "<modFrete>1</modFrete>"
    Else
      strLinha = strLinha & "<modFrete>" & gsFretePago & "</modFrete>"
    End If
     
    'Transporta
    strSQL = "SELECT [Saídas].obs_Transportadora, Transportadoras.CGC, Transportadoras.Nome, "
    strSQL = strSQL & "Transportadoras.[inscrição], Transportadoras.[Endereço], Transportadoras.Cidade, "
    strSQL = strSQL & "Transportadoras.Estado FROM [Saídas], Transportadoras WHERE [Saídas].Filial = " & Filial & "  AND "
    strSQL = strSQL & "[Saídas].Sequência = " & Sequencia & " AND Transportadoras.Nome = [Saídas].obs_Transportadora"
    
    '15/09/2022 - Pablo
    'Caso o <modFrete> seja igual a 9 (sem ocorrência de transporte) então traz a consulta vazia
    strSQL = strSQL & " AND '9' <> '" & gsFretePago & "'"
    
    Set rsTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      
    If Not (rsTransportadora.BOF And rsTransportadora.EOF) Then
      
      strLinha = strLinha & "<transporta>"
      If (rsTransportadora.Fields("CGC").Value <> "") Then
          Dim sCNPJ_CPF_Transportadora As String
          sCNPJ_CPF_Transportadora = RetornaSomenteNumero(rsTransportadora.Fields("CGC").Value & "")
          
          If Len(sCNPJ_CPF_Transportadora) <= 12 Then
              strLinha = strLinha & "<CPF>" & sCNPJ_CPF_Transportadora & "</CPF>"
          Else
              strLinha = strLinha & "<CNPJ>" & sCNPJ_CPF_Transportadora & "</CNPJ>"
          End If
      End If
      If (rsTransportadora.Fields("Nome").Value <> "") Then
          strLinha = strLinha & "<xNome>" & rsTransportadora.Fields("Nome").Value & "</xNome>"
      End If
      If rsTransportadora.Fields("inscrição").Value <> "" Then
          strLinha = strLinha & "<IE>" & RetornaSomenteNumero(rsTransportadora.Fields("inscrição").Value & "") & "</IE>"
      End If
      If (rsTransportadora.Fields("Endereço").Value <> "") Then
          strLinha = strLinha & "<xEnder>" & rsTransportadora.Fields("Endereço").Value & "</xEnder>"
      End If
      If (rsTransportadora.Fields("Cidade").Value <> "") Then
          strLinha = strLinha & "<xMun>" & rsTransportadora.Fields("Cidade").Value & "</xMun>"
      End If
      If (rsTransportadora.Fields("Estado").Value <> "") Then
          strLinha = strLinha & "<UF>" & rsTransportadora.Fields("Estado").Value & "</UF>"
      End If
      strLinha = strLinha & "</transporta>"
      
      If Not IsNull(gsPlaca) And gsPlaca <> "" And gsPlaca <> "0" Then
          strLinha = strLinha & "<veicTransp>"
          strLinha = strLinha & "<placa>" & gsPlaca & "</placa>"
          strLinha = strLinha & "<UF>" & gsUfrmPlaca & "</UF>"
          strLinha = strLinha & "</veicTransp>"
      End If
      
      strLinha = strLinha & "<vol>"
      
      If IsNull(rsSaidas.Fields("obs_Qtde").Value) Then
        str_volumes_quantidade = "0"
      Else
        str_volumes_quantidade = rsSaidas.Fields("obs_Qtde").Value
      End If
       
      str_volumes_quantidade = Format(str_volumes_quantidade, String(15, "0"))
       
      If (str_volumes_quantidade <> "") Then
        strLinha = strLinha & "<qVol>" & str_volumes_quantidade & "</qVol>"
      End If
      
      If (rsSaidas.Fields("obs_Especie").Value <> "") Then
        strLinha = strLinha & "<esp>" & rsSaidas.Fields("obs_Especie").Value & "</esp>"
      End If
       
      If (rsSaidas.Fields("obs_Marca").Value <> "") Then
        strLinha = strLinha & "<marca>" & rsSaidas.Fields("obs_Marca").Value & "</marca>"
      End If
       
      Call IsDataType(dtDouble, rsSaidas.Fields("obs_PesoLiquido").Value, dblPeso)
      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
      strLinha = strLinha & "<pesoL>" & FormataValorTexto(dblPeso, 3) & "</pesoL>"
      End If
      Call IsDataType(dtDouble, rsSaidas.Fields("obs_PesoBruto").Value, dblPeso)
      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
      strLinha = strLinha & "<pesoB>" & FormataValorTexto(dblPeso, 3) & "</pesoB>"
      End If
      strLinha = strLinha & "</vol>"
       
    Else
      '13/09/2022 - Pablo
      'Para evitar mensagem desnecessária
      If gsFretePago <> "9" Then
          MsgBox "Transportadora não encontrada para a Filial " & Filial & ", sequência " & Sequencia, vbExclamation, "Atenção"
      End If
    End If
    strSQL = ""
    rsTransportadora.Close
    
    Set rsTransportadora = Nothing
    
    strLinha = strLinha & "</transp>"
    AddLinha clcArquivo, strLinha

  End If

  'Faturas
  '-------------------------------------------------------------------------------------------------------------
  If PadraoArquivoIntegracao = "XML" Then
  
    '----------------------------
    'XML
  
    Dim sValorLiquidoTotalAux As String
    Dim sValorLiquidoTotalAux2 As String
  
    ' NFe de Devolução não tem grupo <cobr> cobrança
    If bFINALIDADE_NFE_DEVOLUCAO = False Then  ' 4=NFe de Devolução
        'Faturas
        strSQL = "SELECT [Contas a Receber].Parcela, [Contas a Receber].Vencimento, [Contas a Receber].Valor "
        strSQL = strSQL & "FROM [Contas a Receber] WHERE [Contas a Receber].Filial = " & Filial & "  AND "
        strSQL = strSQL & "[Contas a Receber].Sequência = " & Sequencia & " AND [Contas a Receber].Tipo='R' "
        strSQL = strSQL & " order by [Contas a Receber].Vencimento "
        Set rsContasReceber = db.OpenRecordset(strSQL, adOpenDynamic)
        
        With rsContasReceber
          If Not (.BOF And .EOF) Then
            
            .MoveLast
            .MoveFirst
            
            strLinha = "<cobr>"
            strLinha = strLinha & "<fat>"
            strLinha = strLinha & "<nFat>" & lngNotaFiscal & "</nFat>"
            
            If vDespAcessorias > 0 Then
                ' Trato a linha do reg 6000 COM DESPESAS_ACESSORIAS
                strLinha = strLinha & "<vOrig>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value) + vDespAcessorias + v_vlSOMA_DO_ICMS_ST, 2) & "</vOrig>"
                
                If rsSaidas.Fields("Desconto").Value > 0 Then
                    If bTemDescontoNaVenda = True Then
                      strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
                    Else
                      strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsSaidas.Fields("Desconto").Value, 2) & "</vDesc>"
                    End If
                Else
                    If bTemDescontoNaVenda = True Then
                      strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
                    Else
                      strLinha = strLinha & "<vDesc>0.00</vDesc>"
                    End If
                End If
                
                sValorLiquidoTotalAux = FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + vDespAcessorias + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - rsSaidas.Fields("Desconto").Value), 2)
                
                If bTemDescontoNaVenda = True Then
                    strLinha = strLinha & "<vLiq>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + vDespAcessorias + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - rsSaidas.Fields("Desconto").Value) - vSOMA_TOTAL_DESCONTO, 2) & "</vLiq>"
                Else
                    strLinha = strLinha & "<vLiq>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + vDespAcessorias + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - rsSaidas.Fields("Desconto").Value), 2) & "</vLiq>"
                End If
            Else

                strLinha = strLinha & "<vOrig>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value) + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal, 2) & "</vOrig>"
                
                If rsSaidas.Fields("Desconto").Value > 0 Then
                    If bTemDescontoNaVenda = True Then
                      strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
                    Else
                      strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsSaidas.Fields("Desconto").Value, 2) & "</vDesc>"
                    End If
                Else
                    If bTemDescontoNaVenda = True Then
                      strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
                    Else
                      strLinha = strLinha & "<vDesc>0.00</vDesc>"
                    End If
                End If
                
                sValorLiquidoTotalAux = FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - rsSaidas.Fields("Desconto").Value), 2)
                
                If bTemDescontoNaVenda = True Then
                    strLinha = strLinha & "<vLiq>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - vSOMA_TOTAL_DESCONTO), 2) & "</vLiq>"
                Else
                    strLinha = strLinha & "<vLiq>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - rsSaidas.Fields("Desconto").Value), 2) & "</vLiq>"
                End If
            End If
            strLinha = strLinha & "</fat>"
            
            
            '---------------------------------------
            'Se houver outra forma de pagamento (além do parcelado), incluir como primeira DUPLICATA
            Dim bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA As Boolean
            Dim dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA As Double
            
            bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = False
            dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = 0
            
            ' DINHEIRO *********
            If bGerarContaPaga = False And rsSaidas.Fields("Recebe - Dinheiro").Value > 0 Then
               dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = rsSaidas.Fields("Recebe - Dinheiro").Value
               bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True
            End If
            
            ' CHEQUE **********
            Dim rsCheques As Recordset
            Dim totalCheque As Double
            totalCheque = 0
            Set rsCheques = db.OpenRecordset("Select * from [Movimento - Cheques] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
            Do Until rsCheques.EOF
              totalCheque = totalCheque + rsCheques.Fields("Valor")
              rsCheques.MoveNext
            Loop
            rsCheques.Close
            
            If totalCheque > 0 Then
               If totalCheque > rsSaidas.Fields("Total") Then
                 totalCheque = rsSaidas.Fields("Total")
               End If
               dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA + totalCheque
               bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True
            End If
            
            ' CARTAO DE CREDITO ***
            If rsSaidas.Fields("TotalCartaoCredito").Value <> 0 Then
             
               Dim tmpTotalCartaoCredito As Double
               tmpTotalCartaoCredito = rsSaidas.Fields("TotalCartaoCredito").Value
               If tmpTotalCartaoCredito > rsSaidas.Fields("Total") Then
                   tmpTotalCartaoCredito = rsSaidas.Fields("Total").Value
               End If
               dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA + tmpTotalCartaoCredito
               bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True
            End If
            
            ' CARTAO DE DEBITO ***
            If rsSaidas.Fields("TotalCartaoDebito").Value <> 0 Then
            
               Dim tmpTotalCartaoDebito As Double
               tmpTotalCartaoDebito = rsSaidas.Fields("TotalCartaoDebito").Value
               If tmpTotalCartaoDebito > rsSaidas.Fields("Total") Then
                   tmpTotalCartaoDebito = rsSaidas.Fields("Total")
               End If
               
               dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA + tmpTotalCartaoDebito
               bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True
            End If
            
            Dim TotalOutros As Double
            
            If Not IsNull(rsSaidas.Fields("Recebe - Vale").Value) Then
               TotalOutros = TotalOutros + rsSaidas.Fields("Recebe - Vale").Value
            End If

            ' OUTROS PAGAMENTOS ***
            If bGerarContaPaga = False And (TotalOutros > 0 Or vDespAcessorias > 0) Then

               If bGerarContaPaga = False And TotalOutros > 0 Then
                 If TotalOutros > rsSaidas.Fields("Total") Then
                     TotalOutros = rsSaidas.Fields("Total")
                 End If
                 bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True
               Else
                 TotalOutros = 0
               End If

               If vDespAcessorias > 0 Then
                  dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA + TotalOutros + vDespAcessorias
                  bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True
               Else
                 dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA + TotalOutros
               End If
               
            End If
               
            If bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True Then
               strLinha = strLinha & "<dup>"
               strLinha = strLinha & "<nDup>001</nDup>"
               strLinha = strLinha & "<dVenc>" & Format(Now, "YYYY-MM-DD") & "</dVenc>"
               strLinha = strLinha & "<vDup>" & FormataValorTexto(dTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA, 2) & "</vDup>"
               strLinha = strLinha & "</dup>"
               AddLinha clcArquivo, strLinha
             End If
            'Fim outra forma de pagamento (além do parcelado)
             '-----------------------------------------
            
            Dim iNumParc As Integer
            Dim sNumParc As String
            If bTEM_ALGUM_PAGAMENTO_PARA_ULTIMA_DUPLICADA = True Then
              iNumParc = 1
            Else
              iNumParc = 0
            End If
            
            Do Until .EOF
              'Titulos
              
              iNumParc = iNumParc + 1
              sNumParc = CStr(iNumParc)
              If Len(sNumParc) = 1 Then
                  sNumParc = "00" & sNumParc
              ElseIf Len(sNumParc) = 2 Then
                  sNumParc = "0" & sNumParc
              End If
              
              strLinha = strLinha & "<dup>"
              strLinha = strLinha & "<nDup>" & sNumParc & "</nDup>"
              strLinha = strLinha & "<dVenc>" & Format(.Fields("Vencimento").Value, "YYYY-MM-DD") & "</dVenc>"
              
              If vDespAcessorias > 0 Then
                  If .RecordCount = 1 Then
                    'Verificar se o valor liquido da NFe é igual ao total da unica duplicata (tem que ser igual)
                    sValorLiquidoTotalAux2 = FormataValorTexto(.Fields("Valor").Value + vDespAcessorias, 2)
                    If sValorLiquidoTotalAux <> sValorLiquidoTotalAux2 Then
                        sValorLiquidoTotalAux = Replace(sValorLiquidoTotalAux, ".", ",")
                        If bTemDescontoNaVenda = True Then
                          strLinha = strLinha & "<vDup>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - vSOMA_TOTAL_DESCONTO), 2) & "</vDup>"
                        Else
                          strLinha = strLinha & "<vDup>" & FormataValorTexto(sValorLiquidoTotalAux, 2) & "</vDup>"
                        End If
                    Else
                        If bTemDescontoNaVenda = True Then
                            strLinha = strLinha & "<vDup>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - vSOMA_TOTAL_DESCONTO), 2) & "</vDup>"
                        Else
                            strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value + vDespAcessorias, 2) & "</vDup>"
                        End If
                    End If
                  Else
                      strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value + vDespAcessorias, 2) & "</vDup>"
                  End If
              Else
                  If .RecordCount = 1 Then
                    'Verificar se o valor liquido da NFe é igual ao total da unica duplicata (tem que ser igual)
                    sValorLiquidoTotalAux2 = FormataValorTexto(.Fields("Valor").Value, 2)
                    If sValorLiquidoTotalAux <> sValorLiquidoTotalAux2 Then
                        If iNumParc > 1 Then
                              sValorLiquidoTotalAux2 = Replace(sValorLiquidoTotalAux2, ".", ",")

                                If boFreteJaSomadoNaVenda = True Then
                                    strLinha = strLinha & "<vDup>" & FormataValorTexto(sValorLiquidoTotalAux2 + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal, 2) & "</vDup>"
                                Else
                                    strLinha = strLinha & "<vDup>" & FormataValorTexto(sValorLiquidoTotalAux2 + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal, 2) & "</vDup>"
                                End If
                        
                        Else
                            sValorLiquidoTotalAux = Replace(sValorLiquidoTotalAux, ".", ",")
                            If bTemDescontoNaVenda = True Then
                              strLinha = strLinha & "<vDup>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - vSOMA_TOTAL_DESCONTO), 2) & "</vDup>"
                            Else
                              strLinha = strLinha & "<vDup>" & FormataValorTexto(sValorLiquidoTotalAux, 2) & "</vDup>"
                            End If
                        End If
                    
                    Else
                        If bTemDescontoNaVenda = True Then
                            strLinha = strLinha & "<vDup>" & FormataValorTexto((rsSaidas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + v_vlSOMA_DO_ICMS_ST + vFCPST_registroTotal - vSOMA_TOTAL_DESCONTO), 2) & "</vDup>"
                        Else
                            strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value, 2) & "</vDup>"
                        End If
                    End If
                  Else
                      strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value, 2) & "</vDup>"
                  End If
              End If
              strLinha = strLinha & "</dup>"
              AddLinha clcArquivo, strLinha
              strSQL = ""
            .MoveNext
           Loop
           
           strLinha = strLinha & "</cobr>"
           AddLinha clcArquivo, strLinha
          End If
          .Close
        End With
        Set rsContasReceber = Nothing
    End If
    
    
    strLinha = "<pag>"
    
    If bFINALIDADE_NFE_DEVOLUCAO = True Then
        strLinha = strLinha & "<detPag>"
        strLinha = strLinha & "<tPag>90</tPag>"    '90-sem pagamento
        strLinha = strLinha & "<vPag>0.00</vPag>"
        strLinha = strLinha & "</detPag>"
        
        If str_FCP_UF_emitente = "SP" Then
            'Não coloca esta TAG para SP quando o pagamento for 90
        ElseIf str_FCP_UF_emitente <> "PA" Then
            If vDespAcessorias > 0 Then
                strLinha = strLinha & "<vTroco>" & FormataValorTexto(rsSaidas.Fields("Total").Value + rsSaidas.Fields("Frete").Value + vDespAcessorias + v_vlSOMA_DO_ICMS_ST, 2) & "</vTroco>"
            Else
                strLinha = strLinha & "<vTroco>" & FormataValorTexto(rsSaidas.Fields("Total").Value + rsSaidas.Fields("Frete").Value + v_vlSOMA_DO_ICMS_ST, 2) & "</vTroco>"
            End If
        End If
    Else
        ' ****** Tratamento detPag (nfe 4.0) 2018 Pilatti
        
        ' DINHEIRO *********
        If rsSaidas.Fields("Recebe - Dinheiro").Value > 0 Then
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>01</tPag>" ' 01=DINHEIRO
            
            If boTEM_SERVICOS = False Then
                strLinha = strLinha & "<vPag>" & FormataValorTexto((rsSaidas.Fields("Recebe - Dinheiro").Value) + v_vlSOMA_DO_ICMS_ST, 2) & "</vPag>"
            Else
                If rsSaidas.Fields("Recebe - Dinheiro").Value > dbl_vSERVICOS Then
                    strLinha = strLinha & "<vPag>" & FormataValorTexto((rsSaidas.Fields("Recebe - Dinheiro").Value) + v_vlSOMA_DO_ICMS_ST - dbl_vSERVICOS, 2) & "</vPag>"
                    boTEM_SERVICOS = False
                End If
            End If
            strLinha = strLinha & "</detPag>"
            bTEM_ALGUM_PAGAMENTO = True
        ElseIf v_vlSOMA_DO_ICMS_ST > 0 Then
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>01</tPag>" ' 01=DINHEIRO
            strLinha = strLinha & "<vPag>" & FormataValorTexto(v_vlSOMA_DO_ICMS_ST, 2) & "</vPag>"
            strLinha = strLinha & "</detPag>"
        End If
        
        ' CHEQUE **********
        totalCheque = 0
        Set rsCheques = db.OpenRecordset("Select * from [Movimento - Cheques] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
        Do Until rsCheques.EOF
          totalCheque = totalCheque + rsCheques.Fields("Valor")
          rsCheques.MoveNext
        Loop
        rsCheques.Close
        
        If totalCheque > 0 Then
            If totalCheque > rsSaidas.Fields("Total") Then
              totalCheque = rsSaidas.Fields("Total")
            End If
            
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>02</tPag>"    '02=Cheque
            
            If boTEM_SERVICOS = False Then
                strLinha = strLinha & "<vPag>" & FormataValorTexto(totalCheque, 2) & "</vPag>"
            Else
                If totalCheque > dbl_vSERVICOS Then
                    strLinha = strLinha & "<vPag>" & FormataValorTexto(totalCheque - dbl_vSERVICOS, 2) & "</vPag>"
                    boTEM_SERVICOS = False
                End If
            End If
            
            strLinha = strLinha & "</detPag>"
            bTEM_ALGUM_PAGAMENTO = True
        End If
        
        ' CARTAO DE CREDITO ***
        If rsSaidas.Fields("TotalCartaoCredito").Value <> 0 Then
          
            tmpTotalCartaoCredito = rsSaidas.Fields("TotalCartaoCredito").Value
            If tmpTotalCartaoCredito > rsSaidas.Fields("Total") Then
                tmpTotalCartaoCredito = rsSaidas.Fields("Total").Value
            End If
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>1</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>03</tPag>"    '03=Cartão de Crédito
            
            If boTEM_SERVICOS = False Then
                strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoCredito, 2) & "</vPag>"
            Else
                If rsSaidas.Fields("TotalCartaoCredito").Value > dbl_vSERVICOS Then
                    strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoCredito - dbl_vSERVICOS, 2) & "</vPag>"
                    boTEM_SERVICOS = False
                End If
            End If
            
            strLinha = strLinha & "</detPag>"
            bTEM_ALGUM_PAGAMENTO = True
        End If
        
        ' CARTAO DE DEBITO ***
        If rsSaidas.Fields("TotalCartaoDebito").Value <> 0 Then
        
            tmpTotalCartaoDebito = rsSaidas.Fields("TotalCartaoDebito").Value
            If tmpTotalCartaoDebito > rsSaidas.Fields("Total") Then
                tmpTotalCartaoDebito = rsSaidas.Fields("Total")
            End If
            
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>04</tPag>"
            
            If boTEM_SERVICOS = False Then
                strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoDebito, 2) & "</vPag>"
            Else
                If rsSaidas.Fields("TotalCartaoDebito").Value > dbl_vSERVICOS Then
                    strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoDebito - dbl_vSERVICOS, 2) & "</vPag>"
                    boTEM_SERVICOS = False
                End If
            End If
            
            
            strLinha = strLinha & "</detPag>"
            bTEM_ALGUM_PAGAMENTO = True
        End If
        
        If Not IsNull(rsSaidas.Fields("Total Prazo").Value) Then
            TotalOutros = rsSaidas.Fields("Total Prazo").Value
        End If
        
        If Not IsNull(rsSaidas.Fields("Recebe - Vale").Value) Then
            TotalOutros = TotalOutros + rsSaidas.Fields("Recebe - Vale").Value
        End If

        ' OUTROS PAGAMENTOS ***
        If TotalOutros > 0 Or vDespAcessorias > 0 Then
   
            If TotalOutros > 0 Then
              If TotalOutros > rsSaidas.Fields("Total") Then
                  TotalOutros = rsSaidas.Fields("Total")
              End If
            Else
              TotalOutros = 0
            End If
          
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>99</tPag>"
            If vDespAcessorias > 0 Then
              strLinha = strLinha & "<vPag>" & FormataValorTexto(TotalOutros + vDespAcessorias + CDbl(stotal_vICMSST_ParametrizacaoAUX), 2) & "</vPag>"
            Else
              strLinha = strLinha & "<vPag>" & FormataValorTexto(TotalOutros + CDbl(stotal_vICMSST_ParametrizacaoAUX), 2) & "</vPag>"
            End If
            strLinha = strLinha & "</detPag>"
            bTEM_ALGUM_PAGAMENTO = True
        Else
            If FormataValorTexto(TotalOutros + vDespAcessorias + CDbl(stotal_vICMSST_ParametrizacaoAUX), 2) <> "0.00" Then
                strLinha = strLinha & "<detPag>"
                strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
                strLinha = strLinha & "<tPag>99</tPag>"
                strLinha = strLinha & "<vPag>" & FormataValorTexto(TotalOutros + vDespAcessorias + CDbl(stotal_vICMSST_ParametrizacaoAUX), 2) & "</vPag>"
                strLinha = strLinha & "</detPag>"
            End If
        End If
        
        ' Verificar se não teve nenhuma forma de pagamento
        ' Caso não haja, então colocar setando o grupo como 'sem pagamento'
        If bTEM_ALGUM_PAGAMENTO = False Then
        
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<tPag>90</tPag>"    '90-sem pagamento
            strLinha = strLinha & "<vPag>0.00</vPag>"
            strLinha = strLinha & "</detPag>"
        
            If str_FCP_UF_emitente <> "SP" And str_FCP_UF_emitente <> "PA" Then
            
                If vDespAcessorias > 0 Then
                    strLinha = strLinha & "<vTroco>" & FormataValorTexto(rsSaidas.Fields("Total").Value + vDespAcessorias + CDbl(stotal_vICMSST_ParametrizacaoAUX), 2) & "</vTroco>"
                Else
                    'Não coloca esta TAG para SP quando o pagamento for 90
                    If bFINALIDADE_NFE_DEVOLUCAO Then
                      strLinha = strLinha & "<vTroco>" & FormataValorTexto(rsSaidas.Fields("Total").Value, 2) & "</vTroco>"
                    Else
                      strLinha = strLinha & "<vTroco>0.00</vTroco>"
                    End If
                End If
            End If
        End If
    End If
  
    strLinha = strLinha & "</pag>"
    AddLinha clcArquivo, strLinha
  End If
  
Obs:
  '7000;DADOS DO FISCO1;1124 -  - 28 / 35 / 42 / 49 DIAS - / Nr. PEDIDO DE VENDA: 45579
  'Informações adicionais de interesse do fisco
  Dim str_observacoes_NF As String
  If Not IsNull(rsSaidas.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & "Val Aprox dos Tributos R$"
  
  If Not IsNull(rsSaidas.Fields("TotalNCM").Value) Then
      If vl_vTotTrib > 0 Then
          str_observacoes_NF = str_observacoes_NF & vl_vTotTrib & "("
      Else
          str_observacoes_NF = str_observacoes_NF & rsSaidas.Fields("TotalNCM").Value & "("
      End If
  End If
  
  Dim aliquotaNCM As Double
  If Not IsNull(rsSaidas.Fields("TotalNCM").Value) Then
      If vl_vTotTrib > 0 Then
          aliquotaNCM = (Format(vl_vTotTrib, "##,###,##0.00") / Format(rsSaidas("Total"), "##,###,##0.00")) * 100
      Else
          aliquotaNCM = (Format(rsSaidas("TotalNCM"), "##,###,##0.00") / Format(rsSaidas("Total"), "##,###,##0.00")) * 100
      End If
  End If
  
  If Not IsNull(rsSaidas.Fields("TotalNCM").Value) Then
      aliquotaNCM = Format(aliquotaNCM, "##,###,##0.00")
  End If
  
  ' Pilatti Novembro/17
  If rsSaidas.Fields("FreteSomaOuNaoEstimativa").Value = False Then
    If Not IsNull(rsSaidas.Fields("TotalNCM").Value) Then
      If vl_vTotTrib > 0 Then
          aliquotaNCM = (Format(vl_vTotTrib, "##,###,##0.00") / Format(rsSaidas("Produtos"), "##,###,##0.00")) * 100
      Else
          aliquotaNCM = (Format(rsSaidas("TotalNCM"), "##,###,##0.00") / Format(rsSaidas("Produtos"), "##,###,##0.00")) * 100
      End If
      aliquotaNCM = Format(aliquotaNCM, "##,###,##0.00")
    End If
  End If

  If Not IsNull(rsSaidas.Fields("TotalNCM").Value) Then
    str_observacoes_NF = str_observacoes_NF & aliquotaNCM & "%) Fonte IBPT."
  End If
  ' Fim Pilatti
  
  
  If Not IsNull(rsSaidas.Fields("obs_infCpl1").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaidas.Fields("obs_infCpl1").Value
  If Not IsNull(rsSaidas.Fields("obs_infCpl2").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaidas.Fields("obs_infCpl2").Value
  
  str_observacoes_NF = Trim(str_observacoes_NF)
  '---------------------------------------------------------
  '07/04/2011 - Andrea
  'Incluída mensagem de tributacao, no primeiro campo depois da tag 7000
  'Carrega as mensagens para Nota Fiscal
  Dim str_observacoes_NF_fisco As String
  Set m_clcMensagens = New Collection
  Call GetMensagemNotaFiscal(Filial, Sequencia, m_clcMensagens)
  If m_clcMensagens.Count > 0 Then
    If m_clcMensagens.Count >= 1 Then
      If Not IsNull(m_clcMensagens.Item(1)) Then str_observacoes_NF_fisco = str_observacoes_NF_fisco & m_clcMensagens.Item(1)
    End If
    If m_clcMensagens.Count >= 2 Then
      If Not IsNull(m_clcMensagens.Item(2)) Then str_observacoes_NF_fisco = str_observacoes_NF_fisco & m_clcMensagens.Item(2)
    End If
    If m_clcMensagens.Count >= 3 Then
      If Not IsNull(m_clcMensagens.Item(3)) Then str_observacoes_NF_fisco = str_observacoes_NF_fisco & m_clcMensagens.Item(3)
    End If
    If m_clcMensagens.Count >= 4 Then
      If Not IsNull(m_clcMensagens.Item(4)) Then str_observacoes_NF_fisco = str_observacoes_NF_fisco & m_clcMensagens.Item(4)
    End If
    If m_clcMensagens.Count >= 5 Then
      If Not IsNull(m_clcMensagens.Item(5)) Then str_observacoes_NF_fisco = str_observacoes_NF_fisco & m_clcMensagens.Item(5)
    End If
    If m_clcMensagens.Count >= 6 Then
      If Not IsNull(m_clcMensagens.Item(6)) Then str_observacoes_NF_fisco = str_observacoes_NF_fisco & m_clcMensagens.Item(6)
    End If
    str_observacoes_NF_fisco = Trim(str_observacoes_NF_fisco)
  End If
  
  If PadraoArquivoIntegracao = "TXT" Then
    strLinha = "7000;"
    If m_clcMensagens.Count > 0 Then
      strLinha = strLinha & str_observacoes_NF_fisco
    End If
    strLinha = strLinha & ";" & str_observacoes_NF
    AddLinha clcArquivo, strLinha
  ElseIf PadraoArquivoIntegracao = "XML" Then
    strLinha = "<infAdic>"
    If (str_observacoes_NF_fisco <> "") Then
      strLinha = strLinha & "<infAdFisco>" & str_observacoes_NF_fisco & "</infAdFisco>"
    End If
    strLinha = strLinha & "<infCpl>" & str_observacoes_NF & "</infCpl>"
    strLinha = strLinha & "</infAdic>"
    strLinha = strLinha & "</infNFe>"
    strLinha = strLinha & "</NFe>"
    strLinha = strLinha & "</enviNFe>"
    AddLinha clcArquivo, strLinha
  End If
    
 'Fecha tabelas
  rsParametros.Close
  Set rsParametros = Nothing
  rsSaidas.Close
  Set rsSaidas = Nothing
  rsOpSaida.Close
  Set rsOpSaida = Nothing
  Set rsCliFor = Nothing

  'Grava arquivo
  If PadraoArquivoIntegracao = "TXT" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeEnvio" & lngNotaFiscal & ".txt"
  ElseIf PadraoArquivoIntegracao = "XML" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_" & lngNotaFiscal & ".xml"
    
    'Chamada WS (Web Service)
    EnviarXML_SEFAZ Sequencia, lngNotaFiscal, intSerieNotaFiscal, "55", Filial, sCNPJ_EmitenteXML
    
    If gRetOID = -1007 Then
        'Chamada WS (Web Service)
        EnviarXML_SEFAZ Sequencia, lngNotaFiscal, intSerieNotaFiscal, "55", Filial, sCNPJ_EmitenteXML
    End If
  End If
  
  'Retorno da função
  EnviarSaida_4_0_REMODELADA = True
 
  MsgBox "Nota Fiscal Eletrônica enviada com sucesso!", vbInformation, "Atenção"
 
  Exit Function

ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Public Function EnviarXML_SEFAZ(ByVal pSequencia As Long, ByVal pNumero As Long, ByVal pSerie As String, ByVal pModelo As String, ByVal pFilial As Integer, ByVal pCNPJEmitente As String)
On Error GoTo ErroTratamento

    gRetOID = 0

    'CRIA O SOAPCLIENTE E FAZ AS DEVIDAS CONFIGURAÇÕES

    If bSoapClient_MSSoapInit = False Then
      Set soapclient = New SoapClient30
      soapclient.MSSoapInit sSoapClient_MSSoapInit
      soapclient.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL
      bSoapClient_MSSoapInit = True
    End If

    'CRIA VARIAVEIS COM INFORMAÇÕES NECESSÁRIAS PARA CHAMADA DO MÉTODO
    Dim usuario As String
    Dim Senha As String

    'ESTAMOS USANDO O CNPJ DA BENEFIX PARA REALIZAR O TESTE
    'cnpj = "06039615000108"
    usuario = "TaskService"
    Senha = "4sjvpuKl7+Hns//ijN7o"
    
    'CRIA VARIAVEIS DE RETORNO
    Dim Retorno As String
    Dim dtRecebimento As String
    Dim dtRetorno As String
    Dim Desc As String
    Dim Obs As String
    Dim tempoM As String
    Dim retNFeSec As String
    Dim sSql As String
    Dim sStatusDescricao As String
    Dim sStatusDescricao2 As String
    Dim iIndice As Integer
    Dim iIndice2 As Integer
    Dim rsPesq As Recordset
    
    Dim arquivoLote As String
    Dim tipoAutenticacao As Integer

    Dim retOID As Long

    retOID = soapclient.SolicitarImportacaoArquivoXMLasUTF8(sXML_40, pCNPJEmitente, 0, usuario, Senha, retNFeSec)
    
    gRetOID = retOID
    
    If retOID > 0 Then
      'Sucesso no envio
      
      'Realizar o tratamento e INSERT para carga AQUI
      sSql = "UPDATE NFe Set oid_xmlLoteBenefix=" & retOID & ", Status=0 Where Sequencia=" & pSequencia & " And Numero=" & pNumero & " And Filial=" & pFilial
      sSql = sSql & " And Serie ='" & pSerie & "' And Modelo ='" & pModelo & "'"
      db.Execute sSql
      
      sSql = "SELECT * FROM [NFeRetorno] "
      sSql = sSql & "WHERE Filial = " & pFilial & " And Sequencia=" & pSequencia & " And TipoMovimento=0"
      Set rsPesq = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
      If Not (rsPesq.BOF And rsPesq.EOF) Then
        sSql = "Update NFeRetorno set DataHora= '" & Now & "', StatusDescricao='NFe enviada ao Emissor',"
        sSql = sSql & " StatusDescricao2='NFe enviada ao Emissor', status=0 " '& retOID & " "
        sSql = sSql & " Where Filial=" & pFilial & " And Sequencia=" & pSequencia & " And TipoMovimento=0 "
        
        db.Execute sSql
      End If
      rsPesq.Close
      Set rsPesq = Nothing

    Else
      
      sSql = "UPDATE NFe Set oid_xmlLoteBenefix=" & retOID & ", Status=" & retOID & " Where Sequencia=" & pSequencia & " And Numero=" & pNumero & " And Filial=" & pFilial
      sSql = sSql & " And Serie ='" & pSerie & "' And Modelo ='" & pModelo & "'"
      db.Execute sSql
    
      iIndice = InStr(1, retNFeSec, "<ErroXML>")
      If iIndice > 0 Then
        iIndice2 = InStr(iIndice, retNFeSec, "</ErroXML>")
        sStatusDescricao = Mid$(retNFeSec, iIndice + 9, iIndice2 - (iIndice + 9))
        
        If Len(sStatusDescricao) > 255 Then
          sStatusDescricao = Mid$(sStatusDescricao, 1, 255)
        End If
      End If
      sStatusDescricao = TrataCaracteresEspeciaisASCII_Traducao(sStatusDescricao)
      
      iIndice = InStr(1, retNFeSec, "<Mensagem>")
      If iIndice > 0 Then
        iIndice2 = InStr(iIndice, retNFeSec, "</Mensagem>")
        sStatusDescricao2 = Mid$(retNFeSec, iIndice + 10, iIndice2 - (iIndice + 10))
        
        If Len(sStatusDescricao2) > 255 Then
          sStatusDescricao2 = Mid$(sStatusDescricao2, 1, 255)
        End If
      End If
      
      If sStatusDescricao = "" Then
          sStatusDescricao = sStatusDescricao2
      End If
      
      sSql = "SELECT * FROM [NFeRetorno] "
      sSql = sSql & "WHERE Filial = " & pFilial & " And Sequencia=" & pSequencia & " And TipoMovimento=0"
      Set rsPesq = db.OpenRecordset(sSql, dbOpenDynaset, dbReadOnly)
      If (rsPesq.BOF And rsPesq.EOF) Then
        sSql = "Insert into NFeRetorno (Filial, Sequencia, TipoMovimento, DataHora, Protocolo, "
        sSql = sSql & " StatusDescricao, StatusDescricao2, status)"
        sSql = sSql & " values(" & pFilial & "," & pSequencia & ",0,'" & Now & "','',"
        sSql = sSql & "'" & sStatusDescricao & "','" & sStatusDescricao2 & "', " & retOID & ")"
        
        db.Execute sSql
      Else
        sSql = "Update NFeRetorno set DataHora= '" & Now & "', StatusDescricao='" & sStatusDescricao & "',"
        sSql = sSql & " StatusDescricao2='" & sStatusDescricao2 & "', status=" & retOID & " "
        sSql = sSql & " Where Filial=" & pFilial & " And Sequencia=" & pSequencia & " And TipoMovimento=0 "
        
        db.Execute sSql
      End If
      rsPesq.Close
      Set rsPesq = Nothing
    
      If retOID = -1007 Then
          MsgBox "Erro genérico -1007 na transmissão. Selecione a NF-e e envie novamente.", vbInformation
      Else
          MsgBox "Inconsistência na emissão da NF-e: " & pNumero & "  CódigoErro: " & retOID & " -> Mais detalhes, visualize na Aba cor amarela 'NFe Erros e Críticas'.", vbInformation
      End If
    End If
    
    Exit Function
    
ErroTratamento:
  MsgBox "Erro na chamada do WS: Descricao Erro: " & Err.Description & " -- " & retNFeSec, vbInformation
End Function

Public Function CancelarXML_SEFAZ(ByVal pSequencia As Long, ByVal pNumero As Long, ByVal pSerie As String, ByVal pModelo As String, ByVal pFilial As Integer, ByVal pCNPJEmitente As String, ByVal pJustificativa As String, ByVal pTipoDocumento As String)
On Error GoTo ErroTratamento
  
    'CRIA O SOAPCLIENTE E FAZ AS DEVIDAS CONFIGURAÇÕES
    If bSoapClient_MSSoapInit = False Then
      Set soapclient = New SoapClient30
      soapclient.MSSoapInit sSoapClient_MSSoapInit
      soapclient.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL
      bSoapClient_MSSoapInit = True
    End If
    
    'CRIA VARIAVEIS COM INFORMAÇÕES NECESSÁRIAS PARA CHAMADA DO MÉTODO
    Dim usuario As String
    Dim Senha As String
    
    'ESTAMOS USANDO O CNPJ DA BENEFIX PARA REALIZAR O TESTE
    'cnpj = "06039615000108"
    usuario = "TaskService"
    Senha = "4sjvpuKl7+Hns//ijN7o"
    
    'CRIA VARIAVEIS DE RETORNO
    Dim retNFeSec As String
    Dim retCancNFe As String
    Dim sSql As String
    Dim retOID As Long
        
    retOID = soapclient.SolicitarCancelamentoDaNFeNaSEFAZ(pCNPJEmitente, 0, usuario, Senha, pJustificativa, pSerie, pNumero, pTipoDocumento, retNFeSec, retCancNFe)
    
    If retOID = 1 Then
      'Sucesso no cancelamento da nfe
      
      'Realizar o tratamento e INSERT para carga AQUI
      sSql = "UPDATE NFe Set Status=101 Where Sequencia=" & pSequencia & " And Numero=" & pNumero & " And Filial=" & pFilial
      sSql = sSql & " And Serie ='" & pSerie & "' And Modelo ='" & pModelo & "'"
      db.Execute sSql

      MsgBox "Sucesso no cancelamento da NFe: " & CStr(pNumero), vbInformation
    Else
    
      'Condicao de Erro e deve ser tratado pelo ERP.
      'Detalhes na estrutura XML do parametro OUT [retNFeSec], note que [retCancNFe] sera NULO

      MsgBox "Inconsistência no cancelamento da NFe: " & CStr(pNumero) & "  CódigoErro: " & retOID & " -> Mais detalhes, Botão 'Pesquise Retornos' e Visualize na Aba 'Retornos'", vbInformation
    End If
    
    Exit Function
    
ErroTratamento:
  MsgBox "Erro na chamada do WS: " & retNFeSec, vbInformation
End Function

Public Function InutilizarXML_SEFAZ(ByVal pFilial As Integer, ByVal pNumeroInicial As Long, ByVal pNumeroFinal As Long, ByVal pSerie As String, ByVal pAno As Integer, ByVal pCNPJEmitente As String, ByVal pJustificativa As String, ByVal pModelo As String)
On Error GoTo ErroTratamento
  
    'CRIA O SOAPCLIENTE E FAZ AS DEVIDAS CONFIGURAÇÕES
    If bSoapClient_MSSoapInit = False Then
      Set soapclient = New SoapClient30
      soapclient.MSSoapInit sSoapClient_MSSoapInit
      soapclient.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL
      bSoapClient_MSSoapInit = True
    End If
    
    'CRIA VARIAVEIS COM INFORMAÇÕES NECESSÁRIAS PARA CHAMADA DO MÉTODO
    Dim usuario As String
    Dim Senha As String
    
    'ESTAMOS USANDO O CNPJ DA BENEFIX PARA REALIZAR O TESTE
    'cnpj = "06039615000108"
    usuario = "TaskService"
    Senha = "4sjvpuKl7+Hns//ijN7o"
    
    'CRIA VARIAVEIS DE RETORNO
    Dim retNFeSec As String
    Dim retInutNFe As String
    Dim sSql As String

    Dim retOID As Long
        
    retOID = soapclient.SolicitarInutilizacaoDaNFeNaSEFAZ(pCNPJEmitente, 0, usuario, Senha, pAno, pModelo, pSerie, pNumeroInicial, pNumeroFinal, pJustificativa, retNFeSec, retInutNFe)
    
    If retOID = 1 Then
      'Sucesso na Inutilizacao da nfe
      
      'Realizar o tratamento e INSERT
      sSql = "INSERT INTO NFeInutilizadas (Filial, CNPJ, Ano, Serie, NumeroInicial, NumeroFinal, Justificativa, DataHora, Modelo)"
      sSql = sSql & " VALUES (" & pFilial & ",'" & pCNPJEmitente & "'," & pAno & ",'" & pSerie & "'," & pNumeroInicial & "," & pNumeroFinal & ",'" & pJustificativa & "','" & Now & "'," & pModelo & ")"
      db.Execute sSql

      If pModelo = "55" Then
          MsgBox "Sucesso na inutilização da Nota Fiscal.", vbInformation
      Else
          MsgBox "Sucesso na inutilização do Cupom Fiscal.", vbInformation
      End If
    Else
      If pModelo = "55" Then
          MsgBox "Inconsistência na inutilização da Nota Fiscal" & "  CódigoErro: " & retOID & " -> Detalhes: " & retNFeSec, vbInformation
      Else
          MsgBox "Inconsistência na inutilização do Cupom Fiscal" & "  CódigoErro: " & retOID & " -> Detalhes: " & retNFeSec, vbInformation
      End If
    End If
    
    Exit Function
    
ErroTratamento:
  MsgBox "Erro na chamada do WS: " & retNFeSec, vbInformation
End Function

Public Function CartaCorrecaoXML_SEFAZ(ByVal pFilial As Integer, ByVal pNumero As Long, ByVal pSerie As String, ByVal pCNPJEmitente As String, ByVal pDescricaoCC As String, ByVal pTipoDoc As String)
On Error GoTo ErroTratamento
  
    'CRIA O SOAPCLIENTE E FAZ AS DEVIDAS CONFIGURAÇÕES
    If bSoapClient_MSSoapInit = False Then
      Set soapclient = New SoapClient30
      soapclient.MSSoapInit sSoapClient_MSSoapInit
      soapclient.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL
      bSoapClient_MSSoapInit = True
    End If
    
    'CRIA VARIAVEIS COM INFORMAÇÕES NECESSÁRIAS PARA CHAMADA DO MÉTODO
    Dim usuario As String
    Dim Senha As String
    
    'ESTAMOS USANDO O CNPJ DA BENEFIX PARA REALIZAR O TESTE
    'cnpj = "06039615000108"
    usuario = "TaskService"
    Senha = "4sjvpuKl7+Hns//ijN7o"
    
    'CRIA VARIAVEIS DE RETORNO
    Dim retNFeSec As String
    Dim retCCe As String
    Dim pdfNomeDanfe As String
    Dim sSql As String

    Dim retOID As Long
        
    retOID = soapclient.SolicitarRegistroCCeComPDF(pCNPJEmitente, 0, usuario, Senha, pDescricaoCC, pSerie, pNumero, pTipoDoc, retNFeSec, retCCe, pdfNomeDanfe)
    
    If retOID = 1 Then
      'Sucesso na Carta Correcao da nfe
      
      'Realizar o tratamento e INSERT
      sSql = "INSERT INTO NFeCartaCorrecao (Filial, CNPJ, DataHora, Serie, Numero, Descricao, arquivoDanfeCC)"
      sSql = sSql & " VALUES (" & pFilial & ",'" & pCNPJEmitente & "','" & Now & "','" & pSerie & "'," & pNumero & ",'" & pDescricaoCC & "','" & pdfNomeDanfe & "')"
      db.Execute sSql

      MsgBox "Sucesso na realização da Carta Correção da NFe.", vbInformation
    Else
    
      MsgBox "Inconsistência na realização da Carta Correção da NFe." & "  CódigoErro: " & retOID & " -> Detalhes: " & retNFeSec, vbInformation
    End If
    
    Exit Function
    
ErroTratamento:
  MsgBox "Erro na chamada do WS: " & retNFeSec, vbInformation
End Function

Private Function EnviarEntrada_4_0(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal TpEmis As Integer, ByVal dte_data_inicio_contingencia As Date, ByVal str_hora_inicio_contingencia As String, ByVal str_motivo_contingencia As String, ByVal PadraoArquivoIntegracao As String)
  
  '29/11/2010 - Andrea
  m_versao_layout_envio = "4.00"
  Dim strEmail As String
  Dim strDocumento As String
  Dim clcArquivo As New Collection
  Dim strLinha As String
  Dim strChaveNFe As String
  Dim strCodigoAleatorioChaveNFe As String
  Dim intDV As Integer
  Dim intAux As Integer
  Dim strSQL As String
  Dim rsParametros As Recordset
  Dim rsEntradas As Recordset
  Dim rsOpEntrada As Recordset
  Dim rsCliFor As Recordset
  Dim rsCadastroTransportadora As Recordset
  Dim strPastaEnvio As String
  Dim lngNotaFiscal As Long
  Dim intSerieNotaFiscal As Integer
  Dim dblValorTotalProdutos As Double
  Dim dblValorTotalServicos As Double
  Dim rsTransportadora As Recordset
  Dim rsContasPagar As Recordset
    
  'Calculo de PIS e COFINS usa os percentuais da tabela Parâmetros Filial
  Dim dblPercentualPIS As Double
  Dim dblPercentualCofins As Double
  
  '25/01/2010 - mpdea
  'Tratamento para peso bruto e líquido
  Dim dblPeso As Double
  
  '31/03/2011 - Andrea
  Dim blnSimplesNacional As Boolean
  Dim dblAliquotaSimples As Double
  Dim dblPercentualReducaoSimplesNacional As Double
  
  Dim IndicadorIE As String
  Dim IE As String
  
  On Error GoTo ErrHandler
  
  Dim bTEM_ALGUM_PAGAMENTO As Boolean
  bTEM_ALGUM_PAGAMENTO = False
  
  '--------------------------------------------------------------------------------------------------
  'Dados
  '--------------------------------------------------------------------------------------------------
  '
  'Parâmetros da Filial
  strSQL = "SELECT CGC, [Razão Social], Nome, Inscrição, Endereço, EnderecoNumero, EnderecoComplemento, "
  strSQL = strSQL & "Bairro, Cidade, Estado, CEP, Pais, Fone, "
  strSQL = strSQL & "AmbienteNfe, FormatoImpressaoDanfeNfe, ModDetBaseCalculoIcms, ModDetBaseCalculoIcmsSt, "
  strSQL = strSQL & "PastaEnvioNfe, PastaRetornoNfe, PIS, COFINS, InscricaoSuframa, CodigoRegimeTributario, PercentualSimplesNacional, PercentualReducaoBCSimplesNacional  "
  strSQL = strSQL & "FROM [Parâmetros Filial] "
  strSQL = strSQL & "WHERE Filial = " & Filial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsParametros
    If (.BOF And .EOF) Then
      MsgBox "Registro não localizado para a Filial: " & Filial, vbExclamation, "Atenção"
      .Close
      Set rsEntradas = Nothing
      Exit Function
    Else
      'Validações
      strPastaEnvio = Trim(.Fields("PastaEnvioNfe").Value)
      If strPastaEnvio = "" Then
        MsgBox "Pasta de envio não configurada para a Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsEntradas = Nothing
        Exit Function
      Else
        If Right(strPastaEnvio, 1) <> "\" Then
          strPastaEnvio = strPastaEnvio & "\"
        End If
      End If
      
      '25/01/2010 - mpdea
      'Valida código IBGE para UF
      If GetTerritorioUfCodigoIbge(.Fields("Estado").Value & "") = 0 Then
        MsgBox "Código IBGE inválido para a UF da Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsParametros = Nothing
        Exit Function
      End If
      'Valida CNPJ
      If (Trim(.Fields("CGC").Value) & "") = "" Or Not Valid_CGC(.Fields("CGC").Value & "") Then
        MsgBox "CNPJ inválido para a Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsParametros = Nothing
        Exit Function
      End If
      
      sEstadoFilial = .Fields("Estado").Value
    
      dblPercentualPIS = (.Fields("PIS").Value / 100)
      dblPercentualCofins = (.Fields("COFINS").Value / 100)
          
      '31/03/2011 - Andrea
      If (.Fields("CodigoRegimeTributario").Value) = 1 Or (.Fields("CodigoRegimeTributario").Value) = 2 Then
        blnSimplesNacional = True
      Else
      
        blnSimplesNacional = False
      End If
      blnSimplesNacionalGeral = blnSimplesNacional
      
      dblAliquotaSimples = .Fields("PercentualSimplesNacional").Value
      dblPercentualReducaoSimplesNacional = (1 - (.Fields("PercentualReducaoBCSimplesNacional").Value / 100))
    
    End If
  End With
  '
  sPadraoArquivoIntegracao = PadraoArquivoIntegracao
  
  ' *******************************************
  ' PILATTI Outubro/17
  ' Tratar horario de verão
  Dim sFuso As String
  sFuso = ""
  func_horarioFuso bHorarioDeVerao, sEstadoFilial, sFuso
  ' *******************************************

  
  
  'Entrada
  strSQL = "SELECT * From Entradas "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequência = " & Sequencia
  '26/01/2010 - Andrea
  'Tirei o ReadOnly pq estava dando erro na hora de atualizar o numero da NFe na tbl de Entradas
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  '
  'Fornecedor
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM [Cli_For] "
  strSQL = strSQL & "WHERE Código = " & rsEntradas.Fields("Fornecedor").Value
  Set rsCliFor = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsCliFor
    If (.BOF And .EOF) Then
      MsgBox "Fornecedor não localizado para o código: " & rsEntradas.Fields("Fornecedor").Value, vbExclamation, "Atenção"
      .Close
      Set rsCliFor = Nothing
      Exit Function
    Else
      '20/08/2010 - mpdea
      'Verifica condição para Exterior
      If UCase(.Fields("Estado").Value) <> "EX" Then
        '25/01/2010 - mpdea
        'Valida código IBGE para UF
        If GetTerritorioUfCodigoIbge(.Fields("Estado").Value & "") = 0 Then
          MsgBox "Código IBGE inválido para a UF do Fornecedor: " & rsEntradas.Fields("Fornecedor").Value, vbExclamation, "Atenção"
          .Close
          Set rsCliFor = Nothing
          Exit Function
        End If
      End If
      'Valida CNPJ
      If (Trim(.Fields("CGC").Value) & "") = "" Or Not Valid_CGC(.Fields("CGC").Value & "") Then
        If Not Valid_CPF(.Fields("CGC").Value & "") Then
          MsgBox "CNPJ inválido para o Fornecedor: " & rsEntradas.Fields("Fornecedor").Value, vbExclamation, "Atenção"
          .Close
          Set rsCliFor = Nothing
          Exit Function
        End If
      End If
    End If
  End With
  '
  
  '
  'Operação de Entrada
  strSQL = "SELECT Nome, CSO, Tipo  "
  strSQL = strSQL & "FROM [Operações Entrada] "
  strSQL = strSQL & "WHERE Código = " & rsEntradas.Fields("Operação").Value
  Set rsOpEntrada = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rsOpEntrada
    If (.BOF And .EOF) Then
      MsgBox "Operação não localizado para o código: " & rsEntradas.Fields("Operação").Value, vbExclamation, "Atenção"
      .Close
      Set rsOpEntrada = Nothing
      Exit Function
    End If
  End With
  
  ' C = Compra
  ' T = Transferencia
  ' A = Ajuste de entrada
  ' D = Devolução
  ' E = Recebimento de Emprestimo
  ' P = Pedido de Compra para Fornecedor
  If rsOpEntrada.Fields("Tipo").Value = "D" Then
    bFINALIDADE_NFE_DEVOLUCAO = True
  End If
  
  If IsNull(rsEntradas.Fields("obs_Transportadora").Value) Then
    'Cadastro da transportadora
    strSQL = "SELECT Nome FROM Transportadoras WHERE Código = " & rsCliFor.Fields("Transportadora").Value
    Set rsCadastroTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    If rsCadastroTransportadora.BOF And rsCadastroTransportadora.EOF Then
      MsgBox "Transportadora do cliente não localizada, favor verificar. Cliente código: " & rsEntradas.Fields("Fornecedor").Value, vbExclamation, "Atenção"
    End If
  End If
    
  With rsEntradas
    If (.BOF And .EOF) Then
      MsgBox "Movimento não localizado para a Filial: " & Filial & ", sequência: " & Sequencia, vbExclamation, "Atenção"
      .Close
      Set rsEntradas = Nothing
      Exit Function
    Else
      Call IsDataType(dtDouble, .Fields("Produtos").Value, dblValorTotalProdutos)
      dblValorTotalServicos = 0
      '
      'Número da Nota Fiscal
      Call IsDataType(dtLong, .Fields("Nota Impressa").Value, lngNotaFiscal)
      If lngNotaFiscal = 0 Then
        lngNotaFiscal = g_lngNextNotaFiscal_e(Filial)
        .Edit
        
        If IsNull(.Fields("obs_Transportadora").Value) Then
          If rsCadastroTransportadora.RecordCount > 0 Then
            .Fields("obs_Transportadora").Value = rsCadastroTransportadora.Fields("Nome").Value
          Else
            rsCadastroTransportadora.Close
            Set rsCadastroTransportadora = Nothing
          End If
        End If
        
        .Fields("Nota Impressa").Value = lngNotaFiscal
        .Fields("Data Emissão").Value = Now
        .Update
      End If
      '
      'Série da Nota Fiscal
      Call IsDataType(dtInteger, .Fields("SerieNF").Value, intSerieNotaFiscal)
      If intSerieNotaFiscal = 0 Then
          'Tratamento NrSerie pelo arquivo .txt
          If gNrSerieNF > 0 Then
              Dim xConta As Integer
              Dim sCGC As String
              sCGC = rsParametros.Fields("CGC").Value
              sCGC = Replace(sCGC, " ", "")
              sCGC = Replace(sCGC, "/", "")
              sCGC = Replace(sCGC, "\", "")
              sCGC = Replace(sCGC, ".", "")
              sCGC = Replace(sCGC, ",", "")
              sCGC = Replace(sCGC, "-", "")
              sCGC = Replace(sCGC, ";", "")
              For xConta = 0 To gNrSerieNF - 1
                  If gArrayNrSerieNF(xConta, 0) = sCGC Then
                      'NrCnpj1 , SerieNFe1, SerieNFCe1
                      intSerieNotaFiscal = gArrayNrSerieNF(xConta, 1)
                      Exit For
                  End If
              Next
              
              If intSerieNotaFiscal = 0 Then
                  MsgBox "Atenção, não foi encontrado o cadastro da <Serie> da NFe...estará saindo com valor '1'", vbCritical, "Atenção"
                  intSerieNotaFiscal = 1
              End If
          Else
              intSerieNotaFiscal = 1
          End If
          'Fim tratamento NrSerie
      End If
    End If
  End With

  '
  'Chave de Acesso
  strChaveNFe = GetChaveAcesso(Filial, Sequencia, 2)
  If strChaveNFe <> "" Then
    
    '29/11/2010 - Andrea
    '1. Alterações layout 2.0
    'Alterado de 9 para 8 posicoes
    'Código Numérico (aleatório)
    strCodigoAleatorioChaveNFe = Mid(strChaveNFe, 36, 8)
    
    'Dígito verificador
    intDV = Right(strChaveNFe, 1)
    
    Dim intTipoEmissao As Integer
    intTipoEmissao = GetTipoEmissao(Filial, Sequencia, 1)

    '29/11/2010 - Andrea
    '2. Alteraçoes layout 2.0
    If intTipoEmissao <> TpEmis Then 'Se a forma de emissao que está gravada na NFe for diferente da forma de emissao atual
      If TpEmis <> 1 Then 'Se for emissão em estado de contingência
        'Tem que refazer a chave de acesso, como o tipo de emissao de contingencia
        'Código Numérico (aleatório)
        Randomize
        strCodigoAleatorioChaveNFe = Format(Int((99999999) * Rnd + 1), String(8, "0"))
        'Monta o Identificador da TAG a ser assinada
        'Exemplo (campos separados por espaço): 41 0908 79077343000104 55 001 000027025 000000001 4
        strChaveNFe = GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") 'Código da UF
        strChaveNFe = strChaveNFe & Format(Date, "yyMM") 'AAMM da emissão
        strChaveNFe = strChaveNFe & FormataCnpjCpf(rsParametros.Fields("CGC").Value & "") 'CNPJ do Emitente
        strChaveNFe = strChaveNFe & "55" 'Modelo da NFe
        strChaveNFe = strChaveNFe & Format(intSerieNotaFiscal, String(3, "0")) 'Serie da NFe
        strChaveNFe = strChaveNFe & Format(lngNotaFiscal, String(9, "0")) 'Número da NFe
        strChaveNFe = strChaveNFe & TpEmis
        strChaveNFe = strChaveNFe & strCodigoAleatorioChaveNFe 'Código Numérico (aleatório)
        'Dígito verificador
        intDV = CalculaDV(strChaveNFe)
        strChaveNFe = strChaveNFe & intDV
      End If
    End If
    
    
  Else
    'Código Numérico (aleatório)
    Randomize
    '29/11/2010 - Andrea
    '3. Alterações layout 2.0
    'Alterado de 9 para 8 posicoes
    strCodigoAleatorioChaveNFe = Format(Int((99999999) * Rnd + 1), String(8, "0"))
    'Monta o Identificador da TAG a ser assinada
    'Exemplo (campos separados por espaço): 41 0908 79077343000104 55 001 000027025 000000001 4
    
    strChaveNFe = GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") 'Código da UF
    strChaveNFe = strChaveNFe & Format(Date, "yyMM") 'AAMM da emissão
    
    '18/01/2010 - Andrea
    'Para NFe de Entrada, o Fornecedor é o Emitente
    '30/06/2010 - Andrea
    'Trocado o CGC do cliente pelo do fornecedor
    strChaveNFe = strChaveNFe & FormataCnpjCpf(rsParametros.Fields("CGC").Value & "") 'CNPJ do Emitente
    strChaveNFe = strChaveNFe & "55" 'Modelo da NFe
    strChaveNFe = strChaveNFe & Format(intSerieNotaFiscal, String(3, "0")) 'Serie da NFe
    strChaveNFe = strChaveNFe & Format(lngNotaFiscal, String(9, "0")) 'Número da NFe
    
    '29/11/2010 - Andrea
    '4. Alterações layout 2.0
    'Outra importante alteração está na construção da chave de acesso, que agora passa a ter o tpEmis como componente.
    'A construção da chave agora é obtida através da concatenação dos seguintes campos:
    '•  Código UF: 2 posições
    '•  AAMM (Ano e mês): 4 posições
    '•  CNPJ Emitente: 14 posições (precedidos de 0 caso o número não atinja 14 dígitos)
    '•  Modelo: 2 posições
    '•  Série: 3 posições
    '•  Número NF-e: 9 dígitos
    '•  Forma de emissão da NF-e tpEmis: 1 digito
    '1 – Normal – emissão normal;
    '2 – Contingência FS – emissão em contingência com impressão do DANFE em Formulário de Segurança;
    '3 – Contingência SCAN – emissão em contingência no Sistema de Contingência do Ambiente Nacional – SCAN;
    '4 – Contingência DPEC - emissão em contingência com envio da Declaração Prévia de Emissão em Contingência – DPEC;
    '5 – Contingência FS-DA - emissão em contingência com impressão do DANFE em
    '•  Código Numérico (controle ERP): 8 posições
    '•  Dígito Verificador: 1 posição
    'A alteração está no campo Código Numérico que perdeu um dígito para o tpEmis.
    
    strChaveNFe = strChaveNFe & TpEmis
    strChaveNFe = strChaveNFe & strCodigoAleatorioChaveNFe 'Código Numérico (aleatório)
    'Dígito verificador
    intDV = CalculaDV(strChaveNFe)
    strChaveNFe = strChaveNFe & intDV
  End If
  
  '20/09/2009 - mpdea
  'Salva registro
  '29/11/2010 - Andrea
  'Incluído argumento tipo emissao (TpEmis)
  SaveEnvioNFe_4_0 Filial, Sequencia, 2, Now, 0, CByte(rsParametros.Fields("AmbienteNFe").Value), 1, _
    lngNotaFiscal, intSerieNotaFiscal, "55", strChaveNFe, TpEmis
  '--------------------------------------------------------------------------------------------------
  
  
  '--------------------------------------------------------------------------------------------------
  ' Cabeçalho
  '--------------------------------------------------------------------------------------------------
  If PadraoArquivoIntegracao = "TxtERP" Then
     strLinha = "NOTA FISCAL|1"
    AddLinha clcArquivo, strLinha
  ElseIf PadraoArquivoIntegracao = "TXT" Then
    strLinha = FormataCnpjCpf(rsParametros.Fields("CGC").Value & "") & ";" 'CNPJ da Filial
    strLinha = strLinha & "1;" 'Código fixo
    strLinha = strLinha & "ENVIO;" 'Tipo de operação
    strLinha = strLinha & lngNotaFiscal 'Número do Documento Fiscal
    AddLinha clcArquivo, strLinha
  End If
  '--------------------------------------------------------------------------------------------------
  
  
  '--------------------------------------------------------------------------------------------------
  ' A- Dados da Nota Fiscal eletrônica
  '--------------------------------------------------------------------------------------------------
  If PadraoArquivoIntegracao = "XML" Then
    
    ' -----------------------
    ' XML versão NFe 4.00
    
    Dim slngNotaFiscal As String
    slngNotaFiscal = CStr(lngNotaFiscal)
    
    If Len(slngNotaFiscal) < 15 Then
        If Len(slngNotaFiscal) = 1 Then
            slngNotaFiscal = "00000000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 2 Then
            slngNotaFiscal = "0000000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 3 Then
            slngNotaFiscal = "000000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 4 Then
            slngNotaFiscal = "00000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 5 Then
            slngNotaFiscal = "0000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 6 Then
            slngNotaFiscal = "000000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 7 Then
            slngNotaFiscal = "00000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 8 Then
            slngNotaFiscal = "0000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 9 Then
            slngNotaFiscal = "000000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 10 Then
            slngNotaFiscal = "00000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 11 Then
            slngNotaFiscal = "0000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 12 Then
            slngNotaFiscal = "000" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 13 Then
            slngNotaFiscal = "00" + slngNotaFiscal
        ElseIf Len(slngNotaFiscal) = 14 Then
            slngNotaFiscal = "0" + slngNotaFiscal
        End If
    End If
    
    strLinha = "<?xml version=""1.0"" encoding=""UTF-8""?>" '& vbCrLf
    strLinha = strLinha & "<enviNFe versao=""" & VERSAO_PROCESSO_EMISSAO & """ xmlns=""http://www.portalfiscal.inf.br/nfe""><idLote>" + slngNotaFiscal + "</idLote><indSinc>0</indSinc>"
    strLinha = strLinha & "<NFe xmlns=""http://www.portalfiscal.inf.br/nfe"">"
    strLinha = strLinha & "<infNFe versao=""" & VERSAO_PROCESSO_EMISSAO & """ Id=""NFe" & strChaveNFe & """>"
    strLinha = strLinha & "<ide>"
    strLinha = strLinha & "<cUF>" & GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") & "</cUF>"
    strLinha = strLinha & "<cNF>" & strCodigoAleatorioChaveNFe & "</cNF>"
    strLinha = strLinha & "<natOp>" & Left(rsOpEntrada.Fields("Nome").Value, 60) & "</natOp>"
    strLinha = strLinha & "<mod>55</mod>"
    strLinha = strLinha & "<serie>" & intSerieNotaFiscal & "</serie>"
    strLinha = strLinha & "<nNF>" & lngNotaFiscal & "</nNF>"
    strLinha = strLinha & "<dhEmi>" & Format(rsEntradas.Fields("Data Emissão").Value, "yyyy-MM-ddTHH:mm:ss") & sFuso & "</dhEmi>"
    strLinha = strLinha & "<dhSaiEnt>" & Format(rsEntradas.Fields("Data").Value, "yyyy-MM-dd") & "T" & Format(rsEntradas.Fields("Data Emissão").Value, "hh:mm:ss") & sFuso & "</dhSaiEnt>"
    strLinha = strLinha & "<tpNF>0</tpNF>"    ' 0-Entrada; 1-Saída
    
    If rsParametros.Fields("Estado").Value = rsCliFor.Fields("Estado").Value Then
      strLinha = strLinha & "<idDest>1</idDest>"
    ElseIf rsParametros.Fields("Estado").Value = "EX" Then
      strLinha = strLinha & "<idDest>3</idDest>"
    Else
      strLinha = strLinha & "<idDest>2</idDest>"
      isInterestadual = True
    End If

    strLinha = strLinha & "<cMunFG>" & GetTerritorioMunicipioCodigoIbge(rsParametros.Fields("Estado").Value & "", rsParametros.Fields("Cidade").Value & "") & "</cMunFG>"
    strLinha = strLinha & "<tpImp>" & rsParametros.Fields("FormatoImpressaoDanfeNfe").Value & "</tpImp>"
    strLinha = strLinha & "<tpEmis>1</tpEmis>"
    strLinha = strLinha & "<cDV>" & intDV & "</cDV>"
    strLinha = strLinha & "<tpAmb>" & rsParametros.Fields("AmbienteNFe").Value & "</tpAmb>"
    
    If rsEntradas.Fields("FinalidadeNFe").Value > 0 Then
      'Finalidade de emissão da NF-e
      '1- NF-e normal;  2-NF-e complementar;  3 – NF-e de ajuste NOT_FOUND
      strLinha = strLinha & "<finNFe>" & rsEntradas.Fields("FinalidadeNFe").Value & "</finNFe>"
    Else
      strLinha = strLinha & "<finNFe>1</finNFe>"
    End If
    
    If IsNull(rsEntradas.Fields("Consumidor_Final").Value) Then
      strLinha = strLinha & "<indFinal>0</indFinal>"
    Else
      strLinha = strLinha & "<indFinal>" & rsEntradas.Fields("Consumidor_Final").Value & "</indFinal>"
    End If
    If IsNull(rsEntradas.Fields("Presenca_Comprador").Value) Then
      strLinha = strLinha & "<indPres>1</indPres>"
    Else
      strLinha = strLinha & "<indPres>" & rsEntradas.Fields("Presenca_Comprador").Value & "</indPres>"
    End If
    
    strLinha = strLinha & "<procEmi>0</procEmi>"
    strLinha = strLinha & "<verProc>" & VERSAO_PROCESSO_EMISSAO & "</verProc>"
    If TpEmis <> "1" Then 'Se for <> 1, é emissão em regime de contingência
      strLinha = strLinha & "<dhCont>" & Format(dte_data_inicio_contingencia, "yyyy-MM-dd") & "T" & Format(str_hora_inicio_contingencia, "hh:mm:ss") & "</dhCont>"
      strLinha = strLinha & "<xJust>" & str_motivo_contingencia & "</xJust>"
    End If
    
    If Len(rsEntradas.Fields("ChaveReferenciada").Value) > 0 Then  ' NFe de Devolução
      strLinha = strLinha & "<NFref>"
      strLinha = strLinha & "<refNFe>" & rsEntradas.Fields("ChaveReferenciada").Value & "</refNFe>"
      strLinha = strLinha & "</NFref>"
    End If
    
    strLinha = strLinha & "</ide>"
    AddLinha clcArquivo, strLinha
  
  End If
  
  '--------------------------------------------------------------------------------------------------
  'C - Identificação do Emitente da Nota Fiscal Eletrônica  = Emitente = Matriz
  '--------------------------------------------------------------------------------------------------
  If PadraoArquivoIntegracao = "XML" Then
    ' -----------------------------
    ' XML NFe 4.00
    sCNPJ_EmitenteXML = FormataCnpjCpf(rsParametros.Fields("CGC").Value & "")

    strLinha = "<emit>"
    strLinha = strLinha & "<CNPJ>" & FormataCnpjCpf(rsParametros.Fields("CGC").Value & "") & "</CNPJ>"
    strLinha = strLinha & "<xNome>" & Left(Trim(rsParametros.Fields("Razão Social").Value), 60) & "</xNome>"
    strLinha = strLinha & "<xFant>" & Left(Trim(rsParametros.Fields("Nome").Value), 60) & "</xFant>"
    strLinha = strLinha & "<enderEmit>"
    strLinha = strLinha & "<xLgr>" & Left(Trim(rsParametros.Fields("Endereço").Value), 60) & "</xLgr>"
     
    If rsParametros.Fields("EnderecoNumero").Value = "" Then
      strLinha = strLinha & "<nro>S/N</nro>"
    Else
      strLinha = strLinha & "<nro>" & Left(Trim(rsParametros.Fields("EnderecoNumero").Value), 60) & "</nro>"
    End If
     
    If Left(rsParametros.Fields("EnderecoComplemento").Value, 60) <> "" Then
      strLinha = strLinha & "<xCpl>" & Left(Trim(rsParametros.Fields("EnderecoComplemento").Value), 60) & "</xCpl>"
    End If

    strLinha = strLinha & "<xBairro>" & Left(Trim(rsParametros.Fields("Bairro").Value), 60) & "</xBairro>"
    
    If UCase(rsParametros.Fields("Estado").Value) = "EX" Then
      strLinha = strLinha & "<cMun>" & "9999999" & "</cMun>"
      strLinha = strLinha & "<xMun>" & "EXTERIOR" & "</xMun>"
      strLinha = strLinha & "<UF>" & "EX" & "</UF>"
    Else
      strLinha = strLinha & "<cMun>" & GetTerritorioMunicipioCodigoIbge(rsParametros.Fields("Estado").Value, rsParametros.Fields("Cidade").Value & "") & "</cMun>"
      strLinha = strLinha & "<xMun>" & Left(rsParametros.Fields("Cidade").Value, 60) & "</xMun>"
      strLinha = strLinha & "<UF>" & Trim(rsParametros.Fields("Estado").Value) & "</UF>"
    End If
     
    ' Pilatti nfe 4.0 2018
    str_FCP_UF_emitente = Trim(rsParametros.Fields("Estado").Value) 'Sigla da UF
     
    strLinha = strLinha & "<CEP>" & FormataNumero(rsParametros.Fields("CEP").Value & "", 8) & "</CEP>"
     
    If Len(rsParametros.Fields("Pais").Value) > 0 Then
      strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(rsParametros.Fields("Pais").Value & ""), 4) & "</cPais>"
      strLinha = strLinha & "<xPais>" & Trim(Left(rsParametros.Fields("Pais").Value, 60)) & "</xPais>"
    End If
     
     If Left(RetornaSomenteNumero(rsParametros.Fields("Fone").Value), 10) <> "" Then
       strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(rsParametros.Fields("Fone").Value), 10) & "</fone>"
     End If
     
     strLinha = strLinha & "</enderEmit>"
    
     strLinha = strLinha & "<IE>" & Left(FormataIE(rsParametros.Fields("Inscrição").Value, False), 14) & "</IE>"
     
     If dblValorTotalProdutos > 0 And dblValorTotalServicos > 0 Then
       strLinha = strLinha & "<IM>" & Left(rsParametros.Fields("InscricaoMunicipal").Value, 15) & "</IM>"
       strLinha = strLinha & "<CNAE>" & Left(rsParametros.Fields("CNAE").Value, 7) & "</CNAE>"
     End If
       
     strLinha = strLinha & "<CRT>" & rsParametros.Fields("CodigoRegimeTributario").Value & "</CRT>"
     strLinha = strLinha & "</emit>"
     AddLinha clcArquivo, strLinha

  End If
  
  
  
  '----------------------------------------------------------------------------------------------------
  'E - Identificação do Destinatário da Nota Fiscal eletrônica = Destinatário = Fornecedor (30/06/2010)
  '----------------------------------------------------------------------------------------------------
  Dim strEstadoCliente As String
  
  If PadraoArquivoIntegracao = "XML" Then

    ' -------------------------
    ' XML NFe 4.00

    strLinha = strLinha & "<dest>"
    If (rsCliFor.Fields("Física_Jurídica").Value) = "F" Then
      strLinha = strLinha & "<CPF>" & FormataCnpjCpf(rsCliFor.Fields("CGC").Value & "") & "</CPF>"
    Else
      strLinha = strLinha & "<CNPJ>" & FormataCnpjCpf(rsCliFor.Fields("CGC").Value & "") & "</CNPJ>"
    End If
    strLinha = strLinha & "<xNome>" & Left(RemoveCaracteresEspeciaisParaNFE(rsCliFor.Fields("Nome").Value), 60) & "</xNome>"
    strLinha = strLinha & "<enderDest>"
    strLinha = strLinha & "<xLgr>" & Left(RemoveCaracteresEspeciaisParaNFE(rsCliFor.Fields("Endereço").Value), 60) & "</xLgr>"

    If rsCliFor.Fields("Endereço Número").Value = "" Then
      strLinha = strLinha & "<nro>S/N</nro>"
    Else
      strLinha = strLinha & "<nro>" & Left(rsCliFor.Fields("Endereço Número").Value, 60) & "</nro>"
    End If

    If Len(rsCliFor.Fields("Complemento").Value) <> 0 Then
      strLinha = strLinha & "<xCpl>" & Left(RemoveCaracteresEspeciaisParaNFE(rsCliFor.Fields("Complemento").Value), 60) & "</xCpl>"
    End If

    If Left(rsCliFor.Fields("Bairro").Value, 60) = "" Then
      strLinha = strLinha & "<xBairro>" & "CENTRO" & "</xBairro>"
    Else
      strLinha = strLinha & "<xBairro>" & Left(RemoveCaracteresEspeciaisParaNFE(rsCliFor.Fields("Bairro").Value), 60) & "</xBairro>"
    End If
    
    If UCase(rsCliFor.Fields("Estado").Value) = "EX" Then
      strLinha = strLinha & "<cMun>" & "9999999" & "</cMun>"
      strLinha = strLinha & "<xMun>" & "EXTERIOR" & "</xMun>"
      strLinha = strLinha & "<UF>" & "EX" & "</UF>"
    Else
      strLinha = strLinha & "<cMun>" & GetTerritorioMunicipioCodigoIbge(rsCliFor.Fields("Estado").Value, rsCliFor.Fields("Cidade").Value & "") & "</cMun>"
      strLinha = strLinha & "<xMun>" & Left(rsCliFor.Fields("Cidade").Value, 60) & "</xMun>"
      strLinha = strLinha & "<UF>" & rsCliFor.Fields("Estado").Value & "</UF>"
    End If
    
    strLinha = strLinha & "<CEP>" & FormataNumero(rsCliFor.Fields("CEP").Value & "", 8) & "</CEP>"
    strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(rsCliFor.Fields("Pais").Value & ""), 4) & "</cPais>"
    strLinha = strLinha & "<xPais>" & Trim(Left(rsCliFor.Fields("Pais").Value, 60)) & "</xPais>"
    
    ' Mauro - 11/08/2021 Aumentando o tamanho do campo Fone
    
    If Left(RetornaSomenteNumero(rsCliFor.Fields("Fone 1").Value), 12) <> "" Then
      strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(rsCliFor.Fields("Fone 1").Value), 12) & "</fone>"
    End If
    
    strLinha = strLinha & "</enderDest>"
   
    IE = Left(FormataIE(rsCliFor.Fields("Inscrição").Value, UCase(rsCliFor.Fields("Física_Jurídica").Value) = "F"), 14)
    IndicadorIE = rsCliFor.Fields("IndicadorIE").Value
    If IndicadorIE <> 1 And IndicadorIE <> 2 And IndicadorIE <> 9 Then
        IndicadorIE = 9
    End If
    If IndicadorIE = 9 Then
        isIEIsento = True
    End If
    
    ' Pilatti nfe 4.0 2018
    str_FCP_RG_INS_EST_destinatario = IE
      
    strLinha = strLinha & "<indIEDest>" & IndicadorIE & "</indIEDest>"
      
    If IndicadorIE = "1" Then
      strLinha = strLinha & "<IE>" & Left(FormataIE(rsCliFor.Fields("Inscrição").Value, UCase(rsCliFor.Fields("Física_Jurídica").Value) = "F"), 14) & "</IE>" 'IE
    End If
    
    If Len(rsCliFor.Fields("email").Value) <> 0 Then
      strLinha = strLinha & "<email>" & rsCliFor.Fields("email").Value & "</email>"
    End If

    strLinha = strLinha & "</dest>"
      
    'Adicionar linha
    AddLinha clcArquivo, strLinha
  End If
  
    
  ' *****************************************************
  'Inicio FCP (cabecalho calculo) Pilatti
  Dim ds_auxiliarSQL As Recordset
  
  vFCP_registroTotal = 0
  vFCPST_registroTotal = 0
  vFCPSTRet_registroTotal = 0
  iAnoFCP = Year(Now)

  ' Obter aliquota na tabela geral de FCP_AdesaoPercentual_Quick (0%, 1% ou 2%)
  str_sql = "Select ALIQUOTA1 From FCP_AdesaoPercentual_Quick where ESTADO_ORIGEM = '" & str_FCP_UF_destinatario & "'"
  
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaAdesaoUFDestFCP = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA1").Value)
  End If
  
  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing
  
  ' Obter aliquota na tabela geral de ICMS_PERCENTUAL_ESTADOS (SELECT PARA ESTADOS DIFERENTES)
  str_sql = "Select ALIQUOTA From ICMS_PERCENTUAL_ESTADOS where ESTADO_ORIGEM = '" & str_FCP_UF_emitente & "' and ESTADO_DESTINO = '" & str_FCP_UF_destinatario & "'"
                
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaFCP_EMIT_DEST = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA").Value)
  End If
  
  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing
                
  ' Obter aliquota na tabela geral de ICMS_PERCENTUAL_ESTADOS (SELECT PARA O MESMO ESTADO EMITENTE PR ->  PR)
  str_sql = "Select ALIQUOTA From ICMS_PERCENTUAL_ESTADOS where ESTADO_ORIGEM = '" & str_FCP_UF_emitente & "' and ESTADO_DESTINO = '" & str_FCP_UF_emitente & "'"
                
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaFCP_Emitente = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA").Value)
  End If
  
  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing
                
  ' Obter aliquota na tabela geral de ICMS_PERCENTUAL_ESTADOS (SELECT PARA O MESMO ESTADO DESTINATARIO SP -> SP)
  str_sql = "Select ALIQUOTA From ICMS_PERCENTUAL_ESTADOS where ESTADO_ORIGEM = '" & str_FCP_UF_destinatario & "' and ESTADO_DESTINO = '" & str_FCP_UF_destinatario & "'"
                
  Set ds_auxiliarSQL = db.OpenRecordset(str_sql, dbOpenDynaset, dbReadOnly)
  If Not (ds_auxiliarSQL.EOF And ds_auxiliarSQL.BOF) Then
    AliquotaFCP_Destinatario = CDbl(ds_auxiliarSQL.Fields("ALIQUOTA").Value)
  End If
  
  ds_auxiliarSQL.Close
  Set ds_auxiliarSQL = Nothing
                
  ' Fim FCP Pilatti (cabecalho calculo)
  ' ************************************************************
    
 
  '--------------------------------------------------------------------------------------------------
  'I - Produtos e Serviços da NF-e
  '--------------------------------------------------------------------------------------------------
  iDetNumItemXML = 0
  
  Dim sCSO_operEntrada As String
  If Not IsNull(rsOpEntrada.Fields("CSO").Value) Then
      sCSO_operEntrada = rsOpEntrada.Fields("CSO").Value
  Else
      sCSO_operEntrada = ""
  End If
  
  AddProdutoEntrada Filial, Sequencia, clcArquivo, rsEntradas, rsParametros, dblPercentualPIS, dblPercentualCofins, blnSimplesNacional, dblAliquotaSimples, dblPercentualReducaoSimplesNacional, sCSO_operEntrada, UCase(rsCliFor.Fields("Estado").Value), PadraoArquivoIntegracao
  '--------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------
  'W - Valores Totais da NF-e
  '--------------------------------------------------------------------------------------------------
  
  'Valores Totais da NF-e
  strSQL = "SELECT * FROM [Entradas] WHERE Filial = " & Filial
  strSQL = strSQL & "AND Sequência = " & Sequencia
  
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  Dim str_volumes_quantidade As String
  Dim dbl_seguro As Double
        
  If PadraoArquivoIntegracao = "XML" Then
  
    ' -------------------------------
    ' XML
    strLinha = "<total>"
    strLinha = strLinha & "<ICMSTot>"
    
    With rsEntradas
      ' -------------------------------
      ' XML
      strLinha = "<total>"
      strLinha = strLinha & "<ICMSTot>"

      If bFINALIDADE_NFE_DEVOLUCAO = True Or bTEM_ICMS_3179a = True Then
        If .Fields("Base ICM").Value > 0 Then
            If dblValorTotalProdutos + dblValorTotalServicos > .Fields("Base ICM").Value Then
                strLinha = strLinha & "<vBC>" & FormataValorTexto(dblValorTotalProdutos + dblValorTotalServicos, 2) & "</vBC>"
            Else
                strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Base ICM").Value, 2) & "</vBC>"
            End If
        Else
            strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Base ICM").Value, 2) & "</vBC>"
        End If
        
        If vlSOMA_DO_ICMS > 0 Then
            If vlSOMA_DO_ICMS <> .Fields("Valor ICM").Value Then
              strLinha = strLinha & " <vICMS>" & FormataValorTexto(vlSOMA_DO_ICMS, 2) & "</vICMS>"
            Else
              strLinha = strLinha & "<vICMS>" & FormataValorTexto(.Fields("Valor ICM").Value, 2) & "</vICMS>"
            End If
        Else
            strLinha = strLinha & "<vICMS>" & FormataValorTexto(.Fields("Valor ICM").Value, 2) & "</vICMS>"
        End If
        
        If IsNull(.Fields("TotalDesoneracaoICMS").Value) Then
            strLinha = strLinha & "<vICMSDeson>0.00</vICMSDeson>"
        Else
            strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
        End If
        
        ' FCP tratamento NFe 4.0    Pilatti/2018
        ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
        If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacional = False Then ' Or isIEIsento = True) Then
            Dim dAux3 As Double

            'vFCPUFDest ------------------------
            dAux3 = (vFCP_registroTotal * AliquotaFCP_Destinatario) / 100
            strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(dAux3, 2) & "</vFCPUFDest>"

            'vICMSUFDest --------------------------
            ' DIFAL
            CalculoFCPEstadosDiferentes = (vFCP_registroTotal * AliquotaFCP_EMIT_DEST) / 100
            CalculoFCPEstadosIguais = (vFCP_registroTotal * AliquotaFCP_Emitente) / 100
            CalculoFCP_diferenca = CalculoFCPEstadosIguais - CalculoFCPEstadosDiferentes

            ' Ano 2018 80% para o Destino   e   20% para o Emitente
            ' Ano maior que 2018 100% para o Destino
            If (iAnoFCP = 2018) Then
                CalculoFCP_participacaoDest = (CalculoFCP_diferenca * 80) / 100   ' 80%
            Else
                CalculoFCP_participacaoDest = CalculoFCP_diferenca   ' 100% para 2019 em diante
            End If
            strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(CalculoFCP_participacaoDest, 2) & "</vICMSUFDest>"

            ' vICMSUFRemet ------------------------
            If (iAnoFCP = 2018) Then
                CalculoFCP_participacaoEmitente = (CalculoFCP_diferenca * 20) / 100   ' 20%
            Else
                CalculoFCP_participacaoEmitente = 0   ' 100% para 2019 em diante
            End If
            strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(CalculoFCP_participacaoEmitente, 2) & "</vICMSUFRemet>"
            
            'vFCP
            strLinha = strLinha & "<vFCP>" & FormataValorTexto(vFCP_registroTotal, 2) & "</vFCP>"
        Else
            strLinha = strLinha & "<vFCP>0.00</vFCP>"
        End If
        ' Fim tratamento FCP NFe 4.0
        
        strLinha = strLinha & "<vBCST>" & FormataValorTexto(.Fields("Base ICM Subs").Value, 2) & "</vBCST>" 'Base de Cálculo do ICMS ST
        strLinha = strLinha & "<vST>" & FormataValorTexto(.Fields("Valor ICM Subs").Value, 2) & "</vST>" 'Valor Total do ICMS ST
      
     
      ElseIf blnSimplesNacional = False Then
        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Base ICM").Value, 2) & "</vBC>"
        
        If vlSOMA_DO_ICMS > 0 Then
          If vlSOMA_DO_ICMS <> .Fields("Valor ICM").Value Then
            strLinha = strLinha & " <vICMS>" & FormataValorTexto(vlSOMA_DO_ICMS, 2) & "</vICMS>"
          Else
            strLinha = strLinha & " <vICMS>" & FormataValorTexto(.Fields("Valor ICM").Value, 2) & "</vICMS>"
          End If
        Else
          strLinha = strLinha & " <vICMS>" & FormataValorTexto(.Fields("Valor ICM").Value, 2) & "</vICMS>"
        End If
        
        If IsNull(.Fields("TotalDesoneracaoICMS").Value) Then
          strLinha = strLinha & "<vICMSDeson>0.00</vICMSDeson>"
        Else
          strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
        End If
        
        ' FCP tratamento NFe 4.0    Pilatti/2018
        ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
        If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
            Dim dAux2 As Double

            'vFCPUFDest ------------------------
            dAux2 = (vFCP_registroTotal * AliquotaFCP_Destinatario) / 100
            strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(dAux2, 2) & "</vFCPUFDest>"

            'vICMSUFDest --------------------------
            ' DIFAL
            CalculoFCPEstadosDiferentes = (vFCP_registroTotal * AliquotaFCP_EMIT_DEST) / 100
            CalculoFCPEstadosIguais = (vFCP_registroTotal * AliquotaFCP_Emitente) / 100
            CalculoFCP_diferenca = CalculoFCPEstadosIguais - CalculoFCPEstadosDiferentes

            ' Ano 2018 80% para o Destino   e   20% para o Emitente
            ' Ano maior que 2018 100% para o Destino
            If (iAnoFCP = 2018) Then
                CalculoFCP_participacaoDest = (CalculoFCP_diferenca * 80) / 100   ' 80%
            Else
                CalculoFCP_participacaoDest = CalculoFCP_diferenca   ' 100% para 2019 em diante
            End If
            strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(CalculoFCP_participacaoDest, 2) & "</vICMSUFDest>"

            ' vICMSUFRemet ------------------------
            If (iAnoFCP = 2018) Then
                CalculoFCP_participacaoEmitente = (CalculoFCP_diferenca * 20) / 100   ' 20%
            Else
                CalculoFCP_participacaoEmitente = 0   ' 100% para 2019 em diante
            End If
            strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(CalculoFCP_participacaoEmitente, 2) & "</vICMSUFRemet>"
            
            'vFCP
            strLinha = strLinha & "<vFCP>" & FormataValorTexto(vFCP_registroTotal, 2) & "</vFCP>"
        Else
            strLinha = strLinha & "<vFCP>0.00</vFCP>"
        End If
        ' Fim tratamento FCP NFe 4.0
        
        strLinha = strLinha & "<vBCST>" & FormataValorTexto(.Fields("Base ICM Subs").Value, 2) & "</vBCST>" 'Base de Cálculo do ICMS ST
        strLinha = strLinha & "<vST>" & FormataValorTexto(.Fields("Valor ICM Subs").Value, 2) & "</vST>" 'Valor Total do ICMS ST
      
      Else
      
        strLinha = strLinha & "<vBC>0.00</vBC><vICMS>0.00</vICMS>"
        If IsNull(.Fields("TotalDesoneracaoICMS").Value) Then
          strLinha = strLinha & "<vICMSDeson>0.00</vICMSDeson>"
        Else
          strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
        End If
        
        ' FCP tratamento NFe 4.0    Pilatti/2018
        ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
        If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacional = False Then ' Or isIEIsento = True) Then
            dAux2 = 0

            'vFCPUFDest ------------------------
            dAux2 = (vFCP_registroTotal * AliquotaFCP_Destinatario) / 100
            strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(dAux2, 2) & "</vFCPUFDest>"

            'vICMSUFDest --------------------------
            ' DIFAL
            CalculoFCPEstadosDiferentes = (vFCP_registroTotal * AliquotaFCP_EMIT_DEST) / 100
            CalculoFCPEstadosIguais = (vFCP_registroTotal * AliquotaFCP_Emitente) / 100
            CalculoFCP_diferenca = CalculoFCPEstadosIguais - CalculoFCPEstadosDiferentes

            ' Ano 2018 80% para o Destino   e   20% para o Emitente
            ' Ano maior que 2018 100% para o Destino
            If (iAnoFCP = 2018) Then
                CalculoFCP_participacaoDest = (CalculoFCP_diferenca * 80) / 100   ' 80%
            Else
                CalculoFCP_participacaoDest = CalculoFCP_diferenca   ' 100% para 2019 em diante
            End If
            strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(CalculoFCP_participacaoDest, 2) & "</vICMSUFDest>"

            ' vICMSUFRemet ------------------------
            If (iAnoFCP = 2018) Then
                CalculoFCP_participacaoEmitente = (CalculoFCP_diferenca * 20) / 100   ' 20%
            Else
                CalculoFCP_participacaoEmitente = 0   ' 100% para 2019 em diante
            End If
            strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(CalculoFCP_participacaoEmitente, 2) & "</vICMSUFRemet>"
            
            'vFCP
            strLinha = strLinha & "<vFCP>" & FormataValorTexto(vFCP_registroTotal, 2) & "</vFCP>"
        Else
            strLinha = strLinha & "<vFCP>0.00</vFCP>"
        End If
        ' Fim tratamento FCP NFe 4.0
        
        strLinha = strLinha & "<vBCST>0.00</vBCST>" 'Base de Cálculo do ICMS ST
        strLinha = strLinha & "<vST>0.00</vST>"
      End If
      
      ' FCP tratamento NFe 4.0    Pilatti/2018
      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
      If blnSimplesNacional = False Then
        If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
  
            'vFCPST
            strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vFCPST_registroTotal, 2) & "</vFCPST>"
            'vFCPSTRet
            strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vFCPSTRet_registroTotal, 2) & "</vFCPSTRet>"
        Else
            strLinha = strLinha & "<vFCPST>0.00</vFCPST>"
            strLinha = strLinha & "<vFCPSTRet>0.00</vFCPSTRet>"
        End If
      Else
        strLinha = strLinha & "<vFCPST>0.00</vFCPST>"
        strLinha = strLinha & "<vFCPSTRet>0.00</vFCPSTRet>"
      End If
      
      strLinha = strLinha & "<vProd>" & FormataValorTexto(dblValorTotalProdutos + dblValorTotalServicos, 2) & "</vProd>"
      strLinha = strLinha & "<vFrete>" & FormataValorTexto(.Fields("Frete").Value, 2) & "</vFrete>"
      strLinha = strLinha & "<vSeg>0.00</vSeg>"

      '26/04/2011 - Andrea
      'Alterado para que quando tiver desconto no item, preencha o valor do desconto no total
      'Se não tiver desconto no item (o desconto dado no sub total muda o valor do item, por isso o desconto não aparece nos itens)
      'nao vai ser preenchido o desconto no total.
      If m_bln_tem_desconto = True Then
        strLinha = strLinha & "<vDesc>" & FormataValorTexto(.Fields("Desconto").Value, 2) & "</vDesc>"
      Else
        strLinha = strLinha & "<vDesc>0.00</vDesc>"
      End If
      m_bln_tem_desconto = False
      
      strLinha = strLinha & "<vII>0.00</vII>" '& FormataValorTexto(.Fields("").Value, 2) Valor Total do II (Imposto de Importação) NOT_FOUND
      strLinha = strLinha & "<vIPI>" & FormataValorTexto(.Fields("IPI").Value, 2) & "</vIPI>" 'Valor Total do IPI
      
      ' NFe 4.0 Pilatti 2018
      'vIPIDevol
      strLinha = strLinha & "<vIPIDevol>0.00</vIPIDevol>"
      
      'Alterado a fonte do valor do PIS e COFINS, calculado pelo percentual da tabela Parametros Filial
      'Porque na tela de Saidas só grava PIS e COFINS para servicos.
      strLinha = strLinha & "<vPIS>" & FormataValorTexto((.Fields("Total").Value * dblPercentualPIS), 2) & "</vPIS>"
      strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Total").Value * dblPercentualCofins), 2) & "</vCOFINS>"
      
      If vDespAcessorias > 0 Then
        strLinha = strLinha & "<vOutro>" & FormataValorTexto(vDespAcessorias, 2) & "</vOutro>"
        strLinha = strLinha & "<vNF>" & FormataValorTexto(.Fields("Total").Value + vDespAcessorias, 2) & "</vNF>" 'Valor Total da NF-e
      Else
        strLinha = strLinha & "<vOutro>0.00</vOutro>" '& FormataValorTexto(.Fields("").Value, 2) Outras Despesas acessórias Criar campo
        strLinha = strLinha & "<vNF>" & FormataValorTexto(.Fields("Total").Value, 2) & "</vNF>" 'Valor Total da NF-e
      End If
      
      '01/10/2013 - Jean
      'Adicionado campo conforme mudança do layout do infinitri TotalNCM
      If .Fields("TotalNCM").Value > 0 Then
        strLinha = strLinha & "<vTotTrib>" & FormataValorTexto(.Fields("TotalNCM").Value, 2) & "</vTotTrib>"
      End If
      
      strLinha = strLinha & "</ICMSTot>"
      strLinha = strLinha & "</total>"
      'Adicionar Linha
      AddLinha clcArquivo, strLinha
      strSQL = ""
    End With
  
    strSQL = ""
    
    '--------------------------------------------------------------------------------------------------
    'X-Informações do Transporte da NF-e
    '--------------------------------------------------------------------------------------------------
    strLinha = "<transp>"
    
    '16ªlinha [5000;0]
    ' Modalidade do frete
    ' 0- Contratação do Frete por conta do Remetente (CIF)
    ' 1- Contratação do Frete por conta do destinatário/remetente (FOB)
    ' 2- Contratação do Frete por conta de terceiros
    ' 3- Transporte próprio por conta do remetente
    ' 4- Transporte próprio por conta do destinatário
    ' 9- Sem Ocorrência de transporte
    If gsFretePago = "" Then
      strLinha = strLinha & "<modFrete>1</modFrete>"
    Else
      strLinha = strLinha & "<modFrete>" & gsFretePago & "</modFrete>"
    End If
     
    'Transporta
    strSQL = "SELECT [Entradas].obs_Transportadora, Transportadoras.CGC, Transportadoras.Nome, "
    strSQL = strSQL & "Transportadoras.[inscrição], Transportadoras.[Endereço], Transportadoras.Cidade, "
    strSQL = strSQL & "Transportadoras.Estado FROM [Entradas], Transportadoras WHERE [Entradas].Filial = " & Filial & "  AND "
    strSQL = strSQL & "[Entradas].Sequência = " & Sequencia & " AND Transportadoras.Nome = [Entradas].obs_Transportadora"

    '15/09/2022 - Pablo
    'Caso o <modFrete> seja igual a 9 (sem ocorrência de transporte) então traz a consulta vazia
    strSQL = strSQL & " AND '9' <> '" & gsFretePago & "'"

    Set rsTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      
    If Not (rsTransportadora.BOF And rsTransportadora.EOF) Then
      
      strLinha = strLinha & "<transporta>"
      If (rsTransportadora.Fields("CGC").Value <> "") Then
          strLinha = strLinha & "<CNPJ>" & RetornaSomenteNumero(rsTransportadora.Fields("CGC").Value & "") & "</CNPJ>"
      End If
      If (rsTransportadora.Fields("Nome").Value <> "") Then
          strLinha = strLinha & "<xNome>" & rsTransportadora.Fields("Nome").Value & "</xNome>"
      End If
      If rsTransportadora.Fields("inscrição").Value <> "" Then
          strLinha = strLinha & "<IE>" & RetornaSomenteNumero(rsTransportadora.Fields("inscrição").Value & "") & "</IE>"
      End If
      If (rsTransportadora.Fields("Endereço").Value <> "") Then
          strLinha = strLinha & "<xEnder>" & rsTransportadora.Fields("Endereço").Value & "</xEnder>"
      End If
      If (rsTransportadora.Fields("Cidade").Value <> "") Then
          strLinha = strLinha & "<xMun>" & rsTransportadora.Fields("Cidade").Value & "</xMun>"
      End If
      If (rsTransportadora.Fields("Estado").Value <> "") Then
          strLinha = strLinha & "<UF>" & rsTransportadora.Fields("Estado").Value & "</UF>"
      End If
      strLinha = strLinha & "</transporta>"
    
      If Not IsNull(rsEntradas.Fields("obs_Placa").Value) Then
        If rsEntradas.Fields("obs_Placa").Value <> "" Then
          strLinha = strLinha & "<veicTransp>"
          strLinha = strLinha & "<placa>" & rsEntradas.Fields("obs_Placa").Value & "</placa>"
          strLinha = strLinha & "<UF>" & rsEntradas.Fields("obs_UF").Value & "</UF>"
          strLinha = strLinha & "</veicTransp>"
        End If
      End If
      
      strLinha = strLinha & "<vol>"
      
      If IsNull(rsEntradas.Fields("obs_Qtde").Value) Then
        str_volumes_quantidade = "0"
      Else
        str_volumes_quantidade = rsEntradas.Fields("obs_Qtde").Value
      End If
       
      str_volumes_quantidade = Format(str_volumes_quantidade, String(15, "0"))
      
      If (str_volumes_quantidade <> "") Then
        strLinha = strLinha & "<qVol>" & str_volumes_quantidade & "</qVol>"
      End If
      
      If (rsEntradas.Fields("obs_Especie").Value <> "") Then
        strLinha = strLinha & "<esp>" & rsEntradas.Fields("obs_Especie").Value & "</esp>"
      End If
       
      If (rsEntradas.Fields("obs_Marca").Value <> "") Then
        strLinha = strLinha & "<marca>" & rsEntradas.Fields("obs_Marca").Value & "</marca>"
      End If
       
      Call IsDataType(dtDouble, rsEntradas.Fields("obs_PesoLiquido").Value, dblPeso)
      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
      strLinha = strLinha & "<pesoL>" & FormataValorTexto(dblPeso, 3) & "</pesoL>"
      End If
      Call IsDataType(dtDouble, rsEntradas.Fields("obs_PesoBruto").Value, dblPeso)
      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
      strLinha = strLinha & "<pesoB>" & FormataValorTexto(dblPeso, 3) & "</pesoB>"
      End If
      strLinha = strLinha & "</vol>"
       
    Else
      MsgBox "Transportadora não encontrada para a Filial " & Filial & ", sequência " & Sequencia, vbExclamation, "Atenção"
    End If
    strSQL = ""
    rsTransportadora.Close
    
    Set rsTransportadora = Nothing
    
    strLinha = strLinha & "</transp>"
    AddLinha clcArquivo, strLinha
  
  End If
  
  'Fatura
  If PadraoArquivoIntegracao = "TxtERP" Then
    If rsEntradas.Fields("Dinheiro Caixa").Value <> 0 Then
    GoTo obse
    Else
  strLinha = "Y|"
  AddLinha clcArquivo, strLinha
  strLinha = "Y02"
    strLinha = strLinha & "|" & lngNotaFiscal
    strLinha = strLinha & "|" & FormataValorTexto(rsEntradas.Fields("Produtos").Value, 2)
    If rsEntradas.Fields("Desconto").Value > 0 Then
      strLinha = strLinha & "|" & FormataValorTexto(rsEntradas.Fields("Desconto").Value, 2)
    Else
      strLinha = strLinha & "|"
    End If
    strLinha = strLinha & "|" & FormataValorTexto(rsEntradas.Fields("Produtos").Value, 2)
    AddLinha clcArquivo, strLinha
  
  'Faturas
  '-------------------------------------------------------------------------------------------------------------
  '[6100;TITULO: 35072/1;2009-08-10;591.45]
  'Faturas
    strSQL = "SELECT [Contas a Pagar].Contador, [Contas a Pagar].Vencimento, [Contas a Pagar].Valor "
    strSQL = strSQL & "FROM [Contas a Pagar] WHERE [Contas a Pagar].Filial = " & Filial & "  AND "
    strSQL = strSQL & "[Contas a Pagar].Sequência = " & Sequencia & " "
    Set rsContasPagar = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    With rsContasPagar
      If Not (.BOF And .EOF) Then
        Do Until .EOF
        
          strLinha = "Y07"
          strLinha = strLinha & "|" & .Fields("Contador").Value
          
          strLinha = strLinha & "|" & Format(.Fields("Vencimento").Value, "YYYY-MM-DD") 'AAMM da emissão
  
          strLinha = strLinha & "|" & FormataValorTexto(.Fields("Valor").Value, 2)
          AddLinha clcArquivo, strLinha
        
          strSQL = ""
        .MoveNext
        
       Loop
      End If
      .Close
    End With
    Set rsContasPagar = Nothing
  End If
  End If
  If PadraoArquivoIntegracao = "TXT" Then
    
    strLinha = "6000"
    strLinha = strLinha & ";" & lngNotaFiscal
    strLinha = strLinha & ";" & FormataValorTexto(rsEntradas.Fields("Produtos").Value, 2)
    If rsEntradas.Fields("Desconto").Value > 0 Then
      strLinha = strLinha & ";" & FormataValorTexto(rsEntradas.Fields("Desconto").Value, 2)
    Else
      strLinha = strLinha & ";"
    End If
    strLinha = strLinha & ";" & FormataValorTexto(rsEntradas.Fields("Produtos").Value, 2)
    AddLinha clcArquivo, strLinha
  
  'Faturas
  '-------------------------------------------------------------------------------------------------------------
  '[6100;TITULO: 35072/1;2009-08-10;591.45]
  'Faturas
    strSQL = "SELECT [Contas a Pagar].Contador, [Contas a Pagar].Vencimento, [Contas a Pagar].Valor "
    strSQL = strSQL & "FROM [Contas a Pagar] WHERE [Contas a Pagar].Filial = " & Filial & "  AND "
    strSQL = strSQL & "[Contas a Pagar].Sequência = " & Sequencia & " "
    Set rsContasPagar = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    With rsContasPagar
      If Not (.BOF And .EOF) Then
        Do Until .EOF
        
          strLinha = "6100"
          strLinha = strLinha & ";" & .Fields("Contador").Value
          
          strLinha = strLinha & ";" & Format(.Fields("Vencimento").Value, "YYYY-MM-DD") 'AAMM da emissão
  
          strLinha = strLinha & ";" & FormataValorTexto(.Fields("Valor").Value, 2)
          AddLinha clcArquivo, strLinha
        
          strSQL = ""
        .MoveNext
        
       Loop
      End If
      .Close
    End With
    Set rsContasPagar = Nothing
    
    '------------------------------------------------------------------------------------
    ' 6200;
    ' ****** Tratamento troco (nfe 4.0) 2018 Pilatti
    ' Registro troco
    strLinha = "6200;0.00"
    AddLinha clcArquivo, strLinha
    
    ' ****** Tratamento detPag (nfe 4.0) 2018 Pilatti
    ' Registro 6201;
    ' DINHEIRO *********
    If rsSaidas.Fields("Recebe - Dinheiro").Value > 0 Then
        strLinha = strLinha & "6201;0;01;"  '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
        strLinha = strLinha & FormataValorTexto((rsSaidas.Fields("Recebe - Dinheiro").Value), 2)
        AddLinha clcArquivo, strLinha
    End If
    
    ' Registro 6201;
    ' CHEQUE **********
    Dim rsCheques As Recordset
    Dim totalCheque As Double
    totalCheque = 0
    Set rsCheques = db.OpenRecordset("Select * from [Movimento - Cheques] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
    Do Until rsCheques.EOF
      totalCheque = totalCheque + rsCheques.Fields("Valor")
      rsCheques.MoveNext
    Loop
    If totalCheque > 0 Then
        strLinha = strLinha & "6201;0;02;"  '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
    
        If totalCheque > rsSaidas.Fields("Total") Then
          totalCheque = rsSaidas.Fields("Total")
        End If
        
        strLinha = strLinha & FormataValorTexto(totalCheque, 2)
        AddLinha clcArquivo, strLinha
    End If
    
    ' Registro 6201;
    ' CARTAO DE CREDITO ***
    If rsSaidas.Fields("TotalCartaoCredito").Value <> 0 Then
        strLinha = strLinha & "6201;0;03;"  '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
        
        Dim tmpTotalCartaoCredito As Double
        tmpTotalCartaoCredito = rsSaidas.Fields("TotalCartaoCredito").Value
        If tmpTotalCartaoCredito > rsSaidas.Fields("Total") Then
            tmpTotalCartaoCredito = rsSaidas.Fields("Total").Value
        End If
      
        strLinha = strLinha & FormataValorTexto(tmpTotalCartaoCredito, 2)
        AddLinha clcArquivo, strLinha
    End If
    
    ' Registro 6201;
    ' CARTAO DE DEBITO ***
    If rsSaidas.Fields("TotalCartaoDebito").Value <> 0 Then
        strLinha = strLinha & "6201;0;04;"  '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
    
        Dim tmpTotalCartaoDebito As Double
        tmpTotalCartaoDebito = rsSaidas.Fields("TotalCartaoDebito").Value
        If tmpTotalCartaoDebito > rsSaidas.Fields("Total") Then
            tmpTotalCartaoDebito = rsSaidas.Fields("Total")
        End If
        
        strLinha = strLinha & FormataValorTexto(tmpTotalCartaoDebito, 2)
        AddLinha clcArquivo, strLinha
    End If
    
    Dim TotalOutros As Double
    If Not IsNull(rsSaidas.Fields("Total Prazo").Value) Then
        TotalOutros = rsSaidas.Fields("Total Prazo").Value
    End If
    
    If Not IsNull(rsSaidas.Fields("Recebe - Vale").Value) Then
        TotalOutros = TotalOutros + rsSaidas.Fields("Recebe - Vale").Value
    End If
  
    ' Registro 6201;
    ' OUTROS PAGAMENTOS ***
    If TotalOutros > 0 Then
        strLinha = strLinha & "6201;0;99;"  '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros

        If TotalOutros > rsSaidas.Fields("Total") Then
            TotalOutros = rsSaidas.Fields("Total")
        End If
      
        strLinha = strLinha & FormataValorTexto(TotalOutros, 2)
        AddLinha clcArquivo, strLinha
    End If
  
  ElseIf PadraoArquivoIntegracao = "XML" Then
  
    '----------------------------
    'XML
  
    Dim sValorLiquidoTotalAux As String
    Dim sValorLiquidoTotalAux2 As String
    
    ' NFe de Devolução não tem grupo <cobr> cobrança
    If bFINALIDADE_NFE_DEVOLUCAO = False Then  ' 4=NFe de Devolução
        'Faturas
        strSQL = "SELECT [Contas a Pagar].Contador, [Contas a Pagar].Vencimento, [Contas a Pagar].Valor "
        strSQL = strSQL & "FROM [Contas a Pagar] WHERE [Contas a Pagar].Filial = " & Filial & "  AND "
        strSQL = strSQL & "[Contas a Pagar].Sequência = " & Sequencia & " "
        Set rsContasPagar = db.OpenRecordset(strSQL, dbOpenDynaset)
        With rsContasPagar
          If Not (.BOF And .EOF) Then
          
            .MoveLast
            .MoveFirst
            
            strLinha = "<cobr>"
            strLinha = strLinha & "<fat>"
            strLinha = strLinha & "<nFat>" & lngNotaFiscal & "</nFat>"
            
            If vDespAcessorias > 0 Then
                ' Trato a linha do reg 6000 COM DESPESAS_ACESSORIAS
                strLinha = strLinha & "<vOrig>" & FormataValorTexto((rsEntradas.Fields("Produtos").Value + rsEntradas.Fields("IPI").Value) + vDespAcessorias, 2) & "</vOrig>"
                If rsSaidas.Fields("Desconto").Value > 0 Then
                  strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsEntradas.Fields("Desconto").Value, 2) & "</vDesc>"
                Else
                  strLinha = strLinha & "<vDesc>0.00</vDesc>"
                End If
                sValorLiquidoTotalAux = FormataValorTexto((rsEntradas.Fields("Produtos").Value + rsSaidas.Fields("IPI").Value + vDespAcessorias - rsSaidas.Fields("Desconto").Value), 2)
                strLinha = strLinha & "<vLiq>" & FormataValorTexto((rsEntradas.Fields("Produtos").Value + rsEntradas.Fields("IPI").Value + vDespAcessorias - rsSaidas.Fields("Desconto").Value), 2) & "</vLiq>"
            Else
                strLinha = strLinha & "<vOrig>" & FormataValorTexto((rsEntradas.Fields("Produtos").Value + rsEntradas.Fields("IPI").Value), 2) & "</vOrig>"
                If rsSaidas.Fields("Desconto").Value > 0 Then
                  strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsEntradas.Fields("Desconto").Value, 2) & "</vDesc>"
                Else
                  strLinha = strLinha & "<vDesc>0.00</vDesc>"
                End If
                sValorLiquidoTotalAux = FormataValorTexto((rsEntradas.Fields("Produtos").Value + rsEntradas.Fields("IPI").Value - rsSaidas.Fields("Desconto").Value), 2)
                strLinha = strLinha & "<vLiq>" & FormataValorTexto((rsEntradas.Fields("Produtos").Value + rsEntradas.Fields("IPI").Value - rsSaidas.Fields("Desconto").Value), 2) & "</vLiq>"
            End If
            strLinha = strLinha & "</fat>"
            
            Dim iNumParc As Integer
            Dim sNumParc As String
            iNumParc = 0
            
            Do Until .EOF
              'Titulos
              
              iNumParc = iNumParc + 1
              sNumParc = CStr(iNumParc)
              If Len(sNumParc) = 1 Then
                  sNumParc = "00" & sNumParc
              ElseIf Len(sNumParc) = 2 Then
                  sNumParc = "0" & sNumParc
              End If
              
              strLinha = strLinha & "<dup>"
              strLinha = strLinha & "<nDup>" & sNumParc & "</nDup>"
              strLinha = strLinha & "<dVenc>" & Format(.Fields("Vencimento").Value, "YYYY-MM-DD") & "</dVenc>"
              
              If vDespAcessorias > 0 Then
                  If sNumParc = "001" And .RecordCount = 1 Then
                    'Verificar se o valor liquido da NFe é igual ao total da unica duplicata (tem que ser igual)
                    sValorLiquidoTotalAux2 = FormataValorTexto(.Fields("Valor").Value + vDespAcessorias, 2)
                    If sValorLiquidoTotalAux <> sValorLiquidoTotalAux2 Then
                        sValorLiquidoTotalAux = Replace(sValorLiquidoTotalAux, ".", ",")
                        strLinha = strLinha & "<vDup>" & FormataValorTexto(sValorLiquidoTotalAux, 2) & "</vDup>"
                    Else
                        strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value + vDespAcessorias, 2) & "</vDup>"
                    End If
                  Else
                      strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value + vDespAcessorias, 2) & "</vDup>"
                  End If
              Else
                  If sNumParc = "001" And .RecordCount = 1 Then
                    'Verificar se o valor liquido da NFe é igual ao total da unica duplicata (tem que ser igual)
                    sValorLiquidoTotalAux2 = FormataValorTexto(.Fields("Valor").Value, 2)
                    If sValorLiquidoTotalAux <> sValorLiquidoTotalAux2 Then
                        sValorLiquidoTotalAux = Replace(sValorLiquidoTotalAux, ".", ",")
                        strLinha = strLinha & "<vDup>" & FormataValorTexto(sValorLiquidoTotalAux, 2) & "</vDup>"
                    Else
                        strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value, 2) & "</vDup>"
                    End If
                  Else
                      strLinha = strLinha & "<vDup>" & FormataValorTexto(.Fields("Valor").Value, 2) & "</vDup>"
                  End If
              End If
              strLinha = strLinha & "</dup>"
              AddLinha clcArquivo, strLinha
              strSQL = ""
            .MoveNext
           Loop
           strLinha = strLinha & "</cobr>"
           AddLinha clcArquivo, strLinha
          End If
          .Close
        End With
        Set rsContasPagar = Nothing
    End If
    
    
    strLinha = "<pag>"
    
    If bFINALIDADE_NFE_DEVOLUCAO = True Then
        strLinha = strLinha & "<detPag>"
        strLinha = strLinha & "<tPag>90</tPag>"    '90-sem pagamento
        strLinha = strLinha & "<vPag>0.00</vPag>"
        strLinha = strLinha & "</detPag>"
        
        If str_FCP_UF_emitente <> "PA" And str_FCP_UF_emitente <> "SP" Then
            strLinha = strLinha & "<vTroco>" & FormataValorTexto(rsEntradas.Fields("Total").Value, 2) & "</vTroco>"
        ElseIf str_FCP_UF_emitente = "PA" Then
            strLinha = strLinha & "<vTroco>0.00</vTroco>"
        End If
    Else
        ' ****** Tratamento detPag (nfe 4.0) 2018 Pilatti
        ' DINHEIRO *********
        If rsEntradas.Fields("Dinheiro Caixa").Value > 0 Then
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>01</tPag>" ' 01=DINHEIRO
            strLinha = strLinha & "<vPag>" & FormataValorTexto((rsEntradas.Fields("Dinheiro Caixa").Value), 2) & "</vPag>"
            strLinha = strLinha & "</detPag>"
            bTEM_ALGUM_PAGAMENTO = True
        End If
        
        ' CHEQUE **********
        totalCheque = 0
        Set rsCheques = db.OpenRecordset("Select * from [Movimento - Cheques] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
        Do Until rsCheques.EOF
          totalCheque = totalCheque + rsCheques.Fields("Valor")
          rsCheques.MoveNext
        Loop
        If totalCheque > 0 Then
            If totalCheque > rsSaidas.Fields("Total") Then
              totalCheque = rsSaidas.Fields("Total")
            End If
            
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à vista; 1 - Pagamento a prazo; 2- Outros
            strLinha = strLinha & "<tPag>02</tPag>"    '02=Cheque
            strLinha = strLinha & "<vPag>" & FormataValorTexto(totalCheque, 2) & "</vPag>"
            strLinha = strLinha & "</detPag>"
            bTEM_ALGUM_PAGAMENTO = True
        End If
        
        ' Verificar se não teve nenhuma forma de pagamento
        ' Caso não haja, então colocar setando o grupo como 'sem pagamento'
        If bTEM_ALGUM_PAGAMENTO = False Then
        
            strLinha = strLinha & "<detPag>"
            strLinha = strLinha & "<tPag>90</tPag>"    '90-sem pagamento
            strLinha = strLinha & "<vPag>0.00</vPag>"
            strLinha = strLinha & "</detPag>"
            
            If str_FCP_UF_emitente <> "PA" And str_FCP_UF_emitente <> "SP" Then
                strLinha = strLinha & "<vTroco>" & FormataValorTexto(rsEntradas.Fields("Total").Value, 2) & "</vTroco>"
            End If
        End If
    End If
  
    strLinha = strLinha & "</pag>"
    AddLinha clcArquivo, strLinha
  End If
  
obse:
  'INFADIC
  '7000;DADOS DO FISCO1;1124 -  - 28 / 35 / 42 / 49 DIAS - / Nr. PEDIDO DE VENDA: 45579 - Informações adicionais de interesse do fisco
  Dim str_observacoes_NF As String
  If Not IsNull(rsEntradas.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & "Val Aprox dos Tributos R$"
  If Not IsNull(rsEntradas.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & rsEntradas.Fields("TotalNCM").Value
  If Not IsNull(rsEntradas.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & rsEntradas.Fields("TotalNCM").Value & "("
  Dim aliquotaNCM As Double
  aliquotaNCM = (Format(rsEntradas("TotalNCM"), "##,###,##0.00") / Format(rsEntradas("Total"), "##,###,##0.00")) * 100
  aliquotaNCM = Format(aliquotaNCM, "##,###,##0.00")
  If Not IsNull(rsEntradas.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & aliquotaNCM & "%) Fonte IBPT."

  If Not IsNull(rsEntradas.Fields("obs_infCpl1").Value) Then
      str_observacoes_NF = str_observacoes_NF & rsEntradas.Fields("obs_infCpl1").Value
  End If

  str_observacoes_NF = Trim(str_observacoes_NF)
  
  Dim str_observacoes_NF_fisco As String

  If PadraoArquivoIntegracao = "TxtERP" Then
    strLinha = "Z|"
    strLinha = strLinha & str_observacoes_NF
    strLinha = strLinha & "||"
    AddLinha clcArquivo, strLinha
  ElseIf PadraoArquivoIntegracao = "TXT" Then
    strLinha = "7000;"
    strLinha = strLinha & ";" & str_observacoes_NF
    AddLinha clcArquivo, strLinha
  ElseIf PadraoArquivoIntegracao = "XML" Then
    strLinha = "<infAdic>"
    If (str_observacoes_NF_fisco <> "") Then
      strLinha = strLinha & "<infAdFisco>" & str_observacoes_NF_fisco & "</infAdFisco>"
    End If
    strLinha = strLinha & "<infCpl>" & str_observacoes_NF & "</infCpl>"
    strLinha = strLinha & "</infAdic>"
    strLinha = strLinha & "</infNFe>"
    strLinha = strLinha & "</NFe>"
    strLinha = strLinha & "</enviNFe>"
    AddLinha clcArquivo, strLinha
  End If
    
 'Fecha tabelas
  rsParametros.Close
  Set rsParametros = Nothing
  rsEntradas.Close
  Set rsEntradas = Nothing
  rsOpEntrada.Close
  Set rsOpEntrada = Nothing
  Set rsCliFor = Nothing
    
  'Grava arquivo
  If PadraoArquivoIntegracao = "TxtERP" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeEnvioEntrada" & lngNotaFiscal & ".txt"
  ElseIf PadraoArquivoIntegracao = "TXT" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeEnvioEntrada" & lngNotaFiscal & ".txt"
  ElseIf PadraoArquivoIntegracao = "XML" Then
    GravarArquivo clcArquivo, strPastaEnvio & "NFeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_" & lngNotaFiscal & ".xml"
    
    'Chamada WS (Web Service)
    EnviarXML_SEFAZ Sequencia, lngNotaFiscal, intSerieNotaFiscal, "55", Filial, sCNPJ_EmitenteXML
  End If
  
  'Retorno da função
  EnviarEntrada_4_0 = True
 
  MsgBox "Nota Fiscal Eletrônica de Entrada enviada com sucesso!", vbInformation, "Atenção"
 
  Exit Function

ErrHandler:
  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  
End Function

Public Function InutilizarSaida(ByVal Filial As Integer, ByVal MotivoInutilizacao As String, ByVal Serie As String, ByVal NumeroNFeInicial As String, ByVal NumeroNFeFinal As String) As Boolean

  Dim clcArquivo As New Collection
  Dim Data As Date
  Dim strNomeArquivo As String
  Dim strLinha As String
  Dim strSQL As String
  Dim rsParametros As Recordset
  Dim strPastaEnvio As String

  Data = Date$
  strNomeArquivo = "NFeINUTILIZACAOSaida"

  'Demais dados do cancelamento
  strSQL = "SELECT [Parâmetros Filial].CGC, [Parâmetros Filial].AmbienteNFe, [Parâmetros Filial].Estado, "
  strSQL = strSQL & "[Parâmetros Filial].PastaEnvioNFe, [Parâmetros Filial].PadraoArquivoIntegracao "
  strSQL = strSQL & "FROM [Parâmetros Filial] WHERE [Parâmetros Filial].Filial = " & Filial
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
  With rsParametros
    If (.BOF And .EOF) Then
      'Caso não seja localizado campos
      MsgBox "Cadastro da Filial não encontrado.Filial: " & Filial, vbExclamation, "Atenção"
      .Close
      Set rsParametros = Nothing
      Exit Function
    Else
      'Validações
      strPastaEnvio = Trim(.Fields("PastaEnvioNfe").Value)
      If strPastaEnvio = "" Then
        MsgBox "Pasta de envio não configurada para a Filial: " & Filial, vbExclamation, "Atenção"
        .Close
        Set rsParametros = Nothing
        Exit Function
      Else
        If Right(strPastaEnvio, 1) <> "\" Then
          strPastaEnvio = strPastaEnvio & "\"
        End If
      End If
    End If

  End With
    
  Dim strNumeroNotaInicial As String
  Dim strNumeroNotaFinal As String
  Dim m_str_chave_acesso_NF As String
  Dim int_codigo_UF As Integer
  Dim str_CNPJ_CPF As String

  Dim intSerieNotaFiscal As Integer
  
  intSerieNotaFiscal = Serie
  
  strNumeroNotaInicial = NumeroNFeInicial
  strNumeroNotaFinal = NumeroNFeFinal
  str_CNPJ_CPF = Format(RetornaSomenteNumero(rsParametros.Fields("CGC").Value), String(14, "0"))
  int_codigo_UF = GetTerritorioUfCodigoIbge(rsParametros.Fields("Estado").Value & "") 'Código da UF

  m_str_chave_acesso_NF = int_codigo_UF
  m_str_chave_acesso_NF = m_str_chave_acesso_NF & Format(Date, "yy") 'Ano da emissao
  m_str_chave_acesso_NF = m_str_chave_acesso_NF & str_CNPJ_CPF & "55"
  m_str_chave_acesso_NF = m_str_chave_acesso_NF & Format(intSerieNotaFiscal, String(3, "0")) 'Serie da NFe
  m_str_chave_acesso_NF = m_str_chave_acesso_NF & Format(NumeroNFeInicial, String(9, "0")) & Format(NumeroNFeFinal, String(9, "0"))
  
  '--------------------------------------------------------------------------------
  'Estrutura da Inutilizacao
  '-------------------------------------------------------------------------------------------------------------
  strLinha = "<inutNFe xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""2.00"">" & vbCrLf
  strLinha = strLinha & "  <infInut Id=""ID" & m_str_chave_acesso_NF & """>" & vbCrLf
  strLinha = strLinha & "    <tpAmb>" & rsParametros.Fields("AmbienteNFe").Value & "</tpAmb>" & vbCrLf
  strLinha = strLinha & "    <xServ>" & "INUTILIZAR" & "</xServ>" & vbCrLf
  strLinha = strLinha & "    <cUF>" & int_codigo_UF & "</cUF>" & vbCrLf
  strLinha = strLinha & "    <ano>" & Format(Data, "yy") & "</ano>" & vbCrLf
  strLinha = strLinha & "    <CNPJ>" & str_CNPJ_CPF & "</CNPJ>" & vbCrLf
  strLinha = strLinha & "    <mod>" & "55" & "</mod>" & vbCrLf
  strLinha = strLinha & "    <serie>" & intSerieNotaFiscal & "</serie>" & vbCrLf
  strLinha = strLinha & "    <nNFIni>" & strNumeroNotaInicial & "</nNFIni>" & vbCrLf
  strLinha = strLinha & "    <nNFFin>" & strNumeroNotaFinal & "</nNFFin>" & vbCrLf
  strLinha = strLinha & "    <xJust>" & MotivoInutilizacao & "</xJust>" & vbCrLf
  strLinha = strLinha & "  </infInut>" & vbCrLf
  strLinha = strLinha & "</inutNFe>"
  AddLinha clcArquivo, strLinha
  
  If rsParametros.Fields("PadraoArquivoIntegracao") = "TXT" Then
    strNomeArquivo = strNomeArquivo & strNumeroNotaInicial & "_a_" & strNumeroNotaFinal & ".txt"
  Else ' XML
    strNomeArquivo = strNomeArquivo & strNumeroNotaInicial & "_a_" & strNumeroNotaFinal & ".xml"
  End If
     
  
  GravarArquivo clcArquivo, strPastaEnvio & strNomeArquivo
  
  rsParametros.Close
  strSQL = ""
  Set rsParametros = Nothing
  
  InutilizarSaida = True

End Function

Public Function InutilizarEntrada(ByVal Filial As Integer) As Boolean

  Dim clcArquivo As New Collection
  Dim Data As Date
  Dim strSQL As String
  Dim strNomeTXT As String
  Dim strLinha As String
  Dim rsEntradas As Recordset
  
  Data = Date$
  strNomeTXT = "NFeINUTILIZACAOEntrada"
  
  'Versão do arquivo e tipo de operação
  strLinha = "0000;" & VERSAO_LAYOUT_INUTILIZACAO & ";inutNFe"
  AddLinha clcArquivo, strLinha
  
  'ID
  strLinha = "1000;ID" '& NOT_FOUND
  AddLinha clcArquivo, strLinha
  
  'Demais dados do cancelamento
  strSQL = "SELECT [Cli_For].CGC, Entradas.SerieNF, Entradas.[Nota Impressa] FROM [Cli_For], Entradas WHERE Entradas.Filial = " & Filial
  
  Set rsEntradas = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  With rsEntradas
    If Not (.BOF And .EOF) Then
      strLinha = "1100;2;INUTILIZAR;" 'Verificar identificador do ambiente NOT_FOUND
      strLinha = strLinha '& 'Código da UF do solicitante NOT_FOUND
      strLinha = strLinha & ";" & Format(Data, "yy")
      strLinha = strLinha & ";" & Format(RetornaSomenteNumero(.Fields("CGC").Value), String(14, "0"))
      strLinha = strLinha & ";55" ' confirmar se campo será fixo NOT_FOUND
      strLinha = strLinha & ";" & .Fields("SerieNF").Value 'Série da NF-e
      strLinha = strLinha & ";" & .Fields("Nota Impressa").Value
      strLinha = strLinha & ";" & .Fields("Nota Impressa").Value 'confirmar como será feito: Número da NF-e final a ser utilizada NOT_FOUND
      strLinha = strLinha & ";" '& Justificativa NOT_FOUND
      strLinha = strLinha & ";" '& Assinatura XML do grupo identificado pelo atributo "ID" NOT_FOUND
      
      'Adiciona linha
      AddLinha clcArquivo, strLinha
    Else
      'Caso não seja localizado campos
      MsgBox "Movimentação não encontrada para a Filial " & Filial, vbExclamation, "Atenção"
      .Close
      Set rsEntradas = Nothing
      Exit Function
    End If
    .Close
    strSQL = ""
  End With
  Set rsEntradas = Nothing
  
  strNomeTXT = strNomeTXT & ".txt"
  GravarArquivo clcArquivo, strNomeTXT
  
  InutilizarEntrada = True

End Function

Private Function AddProdutoSevicosSaida_REMODELADA(ByVal Filial As String, ByVal Sequencia As Long, ByRef clcArquivo As Collection, ByRef rsSaidas As Recordset, ByRef rsParametros As Recordset, PercentualPIS As Double, PercentualCOFINS As Double, ByVal ConcatenarDescicaoAdicional As Boolean, ByVal blnSimplesNacional As Boolean, ByVal dblAliquotaSimples As Double, ByVal dblPercentualReducaoBC_SN As Double, ByVal CSO_NaturezaOperacao As String, ByVal strEstadoCliente As String, ByVal PadraoArquivoIntegracao As String)
  '15/03/2023 - Pablo
  Dim p_tmp_vIPI As Double 'Valor do IPI
  
  Dim strSQL As String
  Dim strLinha As String
  Dim strTributacao As String
  Dim dblValor As Double
  Dim lngCasasDecimais As Long
  Dim rsSaidasProdutos As Recordset
  Dim rsSaidasServicos As Recordset
  Dim rsSaidasAUXILIAR_ICMS As Recordset
  Dim rsNCM_Cest As Recordset
  Dim sResp As String
  
  Dim dbl_percentual_frete_rateadoACUMULADO As Double
  Dim s_vFreteAux As String
  Dim dbl_vFrete As Double
  Dim dbl_vFreteAcumulado As Double
  Dim dbl_FreteAuxDif As Double
  dbl_vFreteAcumulado = 0
  
  Dim vAuxiliar As Double
  Dim vTesteAux As Double
  Dim v_vlICMS_ST As Double
                
  Dim s_vlSOMA_DO_DESCONTO As String
  Dim vlSOMA_DO_DESCONTO As Double
  vlSOMA_DO_DESCONTO = 0
  s_vlSOMA_DO_DESCONTO = ""
  
  Dim iContaRS As Integer
  Dim iContaLoop As Integer
  Dim s_vlSOMA_DO_ICMS As String
  vlSOMA_DO_ICMS = 0
  s_vlSOMA_DO_ICMS = ""
  iContaRS = 0
  iContaLoop = 0
  
  s_vlSOMA_DO_ICMS_BaseDeCalculo = ""
  v_vlSOMA_DO_ICMS_BaseDeCalculo = 0
  
  Dim iDescontoAuxZ_2 As Integer
  Dim sDescontoAuxZ_2 As String
  
  Dim sCodigoCEST As String
  
  Dim dblBaseCalculoSimplesItem As Double
  
  
  s_vTotTrib = ""
  vl_vTotTrib = 0
  
  dbl_percentual_frete_rateadoACUMULADO = 0
  
  'Valores Totais da NF-e
  strSQL = "SELECT [Base ICM], [Valor ICM], [Base ICM Subs], [Valor ICM Subs], Produtos, [Serviços], Frete, Seguro, [Recebe - Dinheiro],  "
  strSQL = strSQL & "Desconto, IPI, [Total PIS], [Total COFINS], Total, Frete, Seguro, obs_Qtde, obs_Especie, obs_Marca, obs_Placa, obs_UF,"
  strSQL = strSQL & "obs_PesoLiquido, obs_PesoBruto, obs_infCpl1, obs_infCpl2, TotalNCM, TotalDesoneracaoICMS "
  strSQL = strSQL & " ,FreteSomaOuNaoEstimativa, TotalCartaoDebito, TotalCartaoCredito, [Total Prazo], [Recebe - Vale] "
  strSQL = strSQL & " FROM [Saídas] WHERE Filial = " & Filial
  strSQL = strSQL & "AND Sequência = " & Sequencia
  Set rsSaidasAUXILIAR_ICMS = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  If Not IsNull(rsSaidasAUXILIAR_ICMS.Fields("Frete").Value) Then
      dbl_vFrete = rsSaidasAUXILIAR_ICMS.Fields("Frete").Value
  End If
  
  Dim rsSaidasProdutos_CFOP As Recordset
   
  'Rateio dos valores de frete e seguro pelos itens da movimentacao
  Dim dbl_percentual_frete_rateado As Double
  Dim dbl_percentual_seguro_rateado As Double
  Dim dbl_valor_frete_rateado As Double
  Dim dbl_valor_frete_rateado_SOMA As Double
  Dim dbl_valor_seguro_rateado As Double
  Dim s_valor_frete_rateado As String
   
  'Add Produtos
  strSQL = "SELECT Produtos.Código, Produtos.Nome, Produtos.CodigoNBM, Produtos.AliqNCM, [Saídas - Produtos].CFOP,Produtos.[Nome Nota], "
  strSQL = strSQL & "Produtos.[Tipo ICM], Produtos.[Base Cálculo], Produtos.[Redução ICM],  "
  strSQL = strSQL & "Produtos.[Unidade Venda], [Saídas - Produtos].Qtde, [Saídas - Produtos].Preço, "
  strSQL = strSQL & "[Saídas - Produtos].[Preço Final],[Saídas - Produtos].[Desconto Valor], "
  strSQL = strSQL & "Saídas.Produtos, Saídas.Frete, Saídas.Seguro, Saídas.Desconto, [Saídas - Produtos].[Descricao Adicional], "
  strSQL = strSQL & "[Saídas - Produtos].[Situação Tributária], [Saídas - Produtos].ICM, [Saídas - Produtos].[Código sem Grade],  "
  strSQL = strSQL & "[Saídas - Produtos].IPI, Produtos.[IPI_Reduzido], [Saídas - Produtos].[Linha],  "
  strSQL = strSQL & "[Saídas - Produtos].Valor_Aprox_Impostos, [Saídas - Produtos].MotivoDesoneracaoICMS,"
  strSQL = strSQL & "[Saídas - Produtos].ValorICMSDesonerado, [Saídas - Produtos].Percentual_Diferimento, "
  strSQL = strSQL & "[Saídas - Produtos].Desp_Acessorias, Saídas.Operação, [Saídas - Produtos].Desconto, Produtos.TipoSituacaoTributariaPIS, "
  strSQL = strSQL & "Produtos.[Percentual Icm Saida], Produtos.CodigoBeneficio, Produtos.[Percentual Icm Entrada], Produtos.Percentual_IPI_Entrada, "
  strSQL = strSQL & "Produtos.Obs,  Produtos.BaseCalculoICMSST_Saida, Produtos.BaseCalculoICMSST_Entrada, Produtos.Percentual_ICMSST_Entrada, "
  strSQL = strSQL & "Produtos.Percentual_ICMSST_Saida, Produtos.[Percentual IPI], Produtos.SituacaoTributariaEntrada, Produtos.[Situação Tributária] as Prod_Sit_trib, "
  strSQL = strSQL & "Produtos.Tipo, [Saídas - Produtos].Código as CodigoComGrade "
  strSQL = strSQL & " FROM [Saídas - Produtos], Produtos, Saídas "
  strSQL = strSQL & " WHERE [Saídas - Produtos].Filial = " & Filial
  strSQL = strSQL & " AND [Saídas - Produtos].Sequência = " & Sequencia
  strSQL = strSQL & " AND Produtos.Código = [Saídas - Produtos].[Código sem Grade] "
  strSQL = strSQL & " AND Saídas.Filial = " & Filial
  strSQL = strSQL & " AND Saídas.Sequência = " & Sequencia
  strSQL = strSQL & " ORDER BY [Saídas - Produtos].[Linha] "

  Set rsSaidasProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
 
  Dim ValorBaseCalculoICMS As Currency
  Dim AliquotaICMS_Saida As Double
  Dim AliquotaICMS_Entrada As Double
  Dim AliquotaICMSST_Saida As Double
  Dim ValorBCICMSST_Saida As Double
  Dim AliquotaICMSST_Entrada As Double
  Dim ValorBCICMSST_Entrada As Double
  Dim PercentualIPI_Entrada As Double
  Dim PercentualIPI_Saida As Double
  
  
  Dim AliquotaICMS As Double
  Dim BaseCalculoICMS As Double
  Dim ValorICMS As Double
  Dim BaseCalculoICMS_ST As Double
  Dim AliquotaICMS_ST As Double
  Dim ValorICMS_ST As Double
  Dim MargemValorAdicionado_ICMS As Double
  Dim PercentualReducaoBaseICMS As Double
  Dim nomeProduto As String
  Dim Linha As String
  
  Dim dbl_total_produtos As Double
  Dim str_CSO_Produto As String
  Dim str_Desconto_Produto As String
  Dim dbl_Desconto_Produto As Double
      
  With rsSaidasProdutos

    'Total de produtos para rateio do frete e do seguro
    dbl_total_produtos = 0
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        'Trocado o campo preco final pelo calculo preco X qtde porque o preco final vem com ipi
        dbl_total_produtos = dbl_total_produtos + CDbl((.Fields("Qtde").Value * .Fields("Preço").Value))
        iContaRS = iContaRS + 1
        .MoveNext
      Loop
      .MoveFirst
    End If
    
    Dim str_CSO As String
    Dim str_NCM As String
    Dim str_CFOP As String
    
    Do Until .EOF
      Dim dblValorTotalItem As Double
      Dim strValorItem As String
      
      strValorItem = Round(.Fields("Preço").Value, 4)
      
      If Not IsNull(.Fields("TipoSituacaoTributariaPIS").Value) Then
          localTipoSituacaoTributariaPIS = .Fields("TipoSituacaoTributariaPIS").Value
      Else
          'Caso não tenha um valor por produto....então recebe o valor configurado para a FILIAL
          localTipoSituacaoTributariaPIS = gTipoSituacaoTributariaPIS
      End If
      
     
      ' -----------------------------------------------------------------------------
      ' Obter CEST pelo NCM do produtos (se houver)
      strSQL = " Select CEST FROM AliquotasNCM WHERE Codigo = '" & rsSaidasProdutos.Fields("CodigoNBM").Value & "'"
      Set rsNCM_Cest = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      
      If Not (rsNCM_Cest.EOF And rsNCM_Cest.BOF) Then
        If Not IsNull(rsNCM_Cest.Fields("CEST").Value) Then
            sCodigoCEST = rsNCM_Cest.Fields("CEST").Value
        Else
            sCodigoCEST = ""
        End If
      Else
        sCodigoCEST = ""
      End If
      rsNCM_Cest.Close
      ' -----------------------------------------------------------------------------
         
      'Saidas Produtos CFOP
      strSQL = "SELECT CFOP, CSO  "
      strSQL = strSQL & " FROM ProdutoCFOP "
      strSQL = strSQL & " WHERE CodProduto = '" & rsSaidasProdutos.Fields("Código").Value & "'"
      strSQL = strSQL & " And CodOperacao = " & .Fields("Operação").Value

      Set rsSaidasProdutos_CFOP = db.OpenRecordset(strSQL, dbOpenDynaset)
    
      '============================================================================================
      'Rateio do Frete por item
      Dim dbl_percentual_frete_rateadoACUMULADO_dif As Double
      dbl_percentual_frete_rateadoACUMULADO_dif = 100 - dbl_percentual_frete_rateadoACUMULADO
      
      If rsSaidas.Fields("Frete").Value > 0 And dbl_total_produtos > 0 And dbl_percentual_frete_rateadoACUMULADO < 100 Then
          dbl_percentual_frete_rateado = ((rsSaidasProdutos.Fields("Preço Final").Value * 100) / dbl_total_produtos)

          dbl_percentual_frete_rateadoACUMULADO = dbl_percentual_frete_rateadoACUMULADO + dbl_percentual_frete_rateado

          If dbl_percentual_frete_rateadoACUMULADO > 100 Then
              dbl_percentual_frete_rateado = dbl_percentual_frete_rateadoACUMULADO_dif
              If dbl_percentual_frete_rateado < 0 Then
                  dbl_percentual_frete_rateado = 0
                  dbl_valor_frete_rateado = 0
              End If
          End If
          
          If iContaRS = iContaLoop + 1 Then
              If dbl_percentual_frete_rateado <> dbl_percentual_frete_rateadoACUMULADO_dif Then
                  dbl_percentual_frete_rateado = dbl_percentual_frete_rateadoACUMULADO_dif
              End If
          End If

          If dbl_percentual_frete_rateado > 0 Then
              dbl_valor_frete_rateado = ((rsSaidas.Fields("Frete").Value * dbl_percentual_frete_rateado) / 100)
              
              If CDbl(dbl_valor_frete_rateado) > CDbl(rsSaidas.Fields("Frete").Value) Then
                  dbl_valor_frete_rateado = rsSaidas.Fields("Frete").Value
              End If
          End If
      Else
          dbl_valor_frete_rateado = 0
      End If
      '============================================================================================

      'Rateio do valor do Seguro por item
      If rsSaidas.Fields("Seguro").Value > 0 And dbl_total_produtos > 0 Then
        dbl_percentual_seguro_rateado = ((rsSaidasProdutos.Fields("Preço Final").Value * 100) / dbl_total_produtos)
        If dbl_percentual_seguro_rateado > 0 Then
          dbl_valor_seguro_rateado = ((rsSaidas.Fields("Seguro").Value * dbl_percentual_seguro_rateado) / 100)
        End If
      End If

      'Pegar o nome nota se tiver, ao invés do nome do produto
      If Len(rsSaidasProdutos.Fields("Nome Nota").Value) > 0 Then
        nomeProduto = .Fields("Nome Nota").Value
      Else
        nomeProduto = .Fields("Nome").Value
      End If
      
      '13-04-2025 pablo
      If rsParametros("EditarNomeProduto").Value Then
        Dim QUERY As String
        QUERY = "SELECT Nome FROM ProdutoNomeNFe WHERE "
        QUERY = QUERY & "Filial = " & gnCodFilial & " AND "
        QUERY = QUERY & "Sequencia = " & rsSaidas("Sequência") & " AND "
        QUERY = QUERY & "Codigo = '" & .Fields("Código").Value & "';"
        
        Dim rsNomeProd As Recordset
        Set rsNomeProd = db.OpenRecordset(QUERY, dbOpenSnapshot)
        
        If Not (rsNomeProd.BOF And rsNomeProd.EOF) Then
          rsNomeProd.MoveLast
          rsNomeProd.MoveFirst
          If rsNomeProd.RecordCount = 1 Then
            nomeProduto = Trim(CStr(rsNomeProd.Fields("Nome").Value))
          End If
        End If
        
        rsNomeProd.Close
        Set rsNomeProd = Nothing
      End If

      If ConcatenarDescicaoAdicional = True Then
        Dim temp As String
        temp = IIf(IsNull(.Fields("Descricao Adicional").Value), "", .Fields("Descricao Adicional").Value)
        nomeProduto = nomeProduto & " " & RemoveCaracteresEspeciaisParaNFE(temp)
      End If
      
      Linha = .Fields("Linha").Value
      
      ' ------------------------------
      ' XML NFe 4.00
        
      iDetNumItemXML = iDetNumItemXML + 1
      strLinha = "<det nItem=""" & iDetNumItemXML & """>"
      strLinha = strLinha & "<prod>"
      strLinha = strLinha & "<cProd>" & .Fields("Código").Value & "</cProd>"
      strLinha = strLinha & "<cEAN>SEM GTIN</cEAN>"

      nomeProduto = RemoveCaracteresEspeciaisParaNFE(Left(nomeProduto, 120))

      If .Fields("Tipo").Value = "G" Then
          strLinha = strLinha & "<xProd>" & nomeProduto & " " & .Fields("CodigoComGrade").Value & "</xProd>"
      Else
          strLinha = strLinha & "<xProd>" & nomeProduto & "</xProd>"
      End If
      strLinha = strLinha & "<NCM>" & .Fields("CodigoNBM").Value & "</NCM>"
        
      If IsNull(.Fields("CodigoNBM").Value) Or .Fields("CodigoNBM").Value = "" Then
          gstr_posicaoNFeErro_telaAcelerador = gstr_posicaoNFeErro_telaAcelerador & "Produto " & rsSaidasProdutos.Fields("Código").Value & " : não possui NCM vinculado" & vbCrLf
      End If
        
      If sCodigoCEST <> "" Then
          strLinha = strLinha & "<CEST>" & sCodigoCEST & "</CEST>"
      End If
      
      ' ===========================================================================================
      ' Busca Situação Tributaria
      If bFINALIDADE_NFE_DEVOLUCAO = True Then
          ' Se for Devolução, Remessa ou GratisSaida
          If Not IsNull(rsSaidasProdutos.Fields("SituacaoTributariaEntrada").Value) Then
              If Trim(rsSaidasProdutos.Fields("SituacaoTributariaEntrada").Value) <> "" Then
                  If Len(Trim(rsSaidasProdutos.Fields("SituacaoTributariaEntrada").Value)) > 1 Then
                      strTributacao = rsSaidasProdutos.Fields("SituacaoTributariaEntrada").Value
                  Else
                      strTributacao = rsSaidasProdutos.Fields("Situação Tributária").Value
                  End If
              Else
                  strTributacao = rsSaidasProdutos.Fields("Situação Tributária").Value
              End If
          Else
              strTributacao = rsSaidasProdutos.Fields("Situação Tributária").Value
          End If
      Else
          ' Se for Saída Normal (venda)
          strTributacao = rsSaidasProdutos.Fields("Situação Tributária").Value
      End If
      
      If (Trim(strTributacao) = "" Or Trim(strTributacao) = "0") And blnSimplesNacional = True Then
          ' Simples Nacional
          strTributacao = "0"
      ElseIf (Trim(strTributacao) = "" Or CDbl(Trim(strTributacao)) = 0) And blnSimplesNacional = False Then
          ' Lucro Real
          
          If Not IsNull(rsSaidasProdutos.Fields("Prod_Sit_trib").Value) Then
              If Trim(rsSaidasProdutos.Fields("Prod_Sit_trib").Value) = "" Then
                  strTributacao = "0"
                  gstr_posicaoNFeErro_telaAcelerador = gstr_posicaoNFeErro_telaAcelerador & "Produto " & rsSaidasProdutos.Fields("Código").Value & " : não possui cadastro de Situação Tributária" & vbCrLf
              Else
                  strTributacao = Trim(rsSaidasProdutos.Fields("Prod_Sit_trib").Value)
              End If
          Else
              strTributacao = "0"
          End If
      End If
      ' ===========================================================================================
      
      ' Inicio tratamento cBenef (para empresa LUCRO REAL)
      If blnSimplesNacional = False Then
          If Right(strTributacao, 2) = "40" Then
              If Not IsNull(.Fields("CodigoBeneficio").Value) Then
                  If .Fields("CodigoBeneficio").Value = "" Then
                      strLinha = strLinha & "<cBenef>SEM CBENEF</cBenef>"
                  Else
                      strLinha = strLinha & "<cBenef>" & .Fields("CodigoBeneficio").Value & "</cBenef>"
                  End If
              Else
                  'Ex: <cBenef>PR810002</cBenef>
                  strLinha = strLinha & "<cBenef>SEM CBENEF</cBenef>"
              End If
          End If
      End If
      ' Fim tratamento cBenef
        
      If Not rsSaidasProdutos_CFOP.EOF Then
          If Not IsNull(rsSaidasProdutos_CFOP.Fields("CFOP").Value) And rsSaidasProdutos_CFOP.Fields("CFOP").Value <> 0 Then
              str_CFOP = Format(rsSaidasProdutos_CFOP.Fields("CFOP").Value, String(4, "0")) 'CFOP do cadastro de produto X CFOP
          Else
              str_CFOP = Format(.Fields("CFOP").Value, String(4, "0")) 'CFOP dos itens da venda
          End If
                
          If Not IsNull(rsSaidasProdutos_CFOP.Fields("CSO").Value) And rsSaidasProdutos_CFOP.Fields("CSO").Value <> 0 Then
              str_CSO_Produto = Format(rsSaidasProdutos_CFOP.Fields("CSO").Value, String(3, "0")) 'CSO do produto X Natureza de operacao
          Else
              str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
          End If
      Else
          str_CFOP = Format(.Fields("CFOP").Value, String(4, "0")) 'CFOP do Produto
          str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
      End If
        
      If str_CFOP = "" Then
          Dim rsAuxCFOP2 As Recordset
          Set rsAuxCFOP2 = db.OpenRecordset("Select [Código Fiscal], CSO From [Operações Saída] WHERE Código = " & rsSaidas.Fields("Operação").Value, dbOpenDynaset, dbReadOnly)
          If Not (rsAuxCFOP2.BOF And rsAuxCFOP2.EOF) Then
            If Not IsNull(rsAuxCFOP2.Fields("Código Fiscal").Value) Then
              str_CFOP = rsAuxCFOP2.Fields("Código Fiscal").Value
              
              If str_CSO_Produto = "" And Not IsNull(rsAuxCFOP2.Fields("CSO").Value) Then
                  str_CSO_Produto = rsAuxCFOP2.Fields("CSO").Value
              End If
            End If
          End If
          rsAuxCFOP2.Close
          Set rsAuxCFOP2 = Nothing
      End If

      ' ***************************************************************
      ' Tratamento de inconsistência no cadastro de CFOP
      If IsNull(str_CFOP) Or str_CFOP = "" Then
          While str_CFOP = ""
              sResp = InputBox("O cadastro do produto " & vbCrLf & "CÓDIGO: " & rsSaidasProdutos.Fields("Código").Value & vbCrLf & "NOME: " & nomeProduto & vbCrLf & " está INCOMPLETO ! " & vbCrLf & vbCrLf & "Pontualmente lhe ajudaremos a seguir com a emissão desta NF-e. Porém, você deverá COMPLETAR este cadastro posteriormente." & vbCrLf & vbCrLf & "Agora, Digite o CFOP correto vinculado a este produto " & vbCrLf & "Ex: 5102, 6102, entre outros", "Cadastro INCOMPLETO - Vincular CFOP")
              
              If IsNull(sResp) Then
                  str_CFOP = ""
              ElseIf Not IsNumeric(sResp) Then
                  str_CFOP = ""
              ElseIf Len(Trim(sResp)) <> 4 Then
                  str_CFOP = ""
              Else
                  str_CFOP = sResp
              End If
          Wend
      End If
      ' ***************************************************************

      strLinha = strLinha & "<CFOP>" & str_CFOP & "</CFOP>"

      Dim sTrataDado As String
      sTrataDado = trataCaracteresEspeciaisNFE(.Fields("Unidade Venda").Value)
        
      If Trim(sTrataDado) = "" Then
          sTrataDado = "UN"
      End If
      strLinha = strLinha & "<uCom>" & sTrataDado & "</uCom>"
      strLinha = strLinha & "<qCom>" & FormataValorTexto(.Fields("Qtde").Value, 4) & "</qCom>"
      strLinha = strLinha & "<vUnCom>" & FormataValorTexto(strValorItem, 4) & "</vUnCom>"
      strLinha = strLinha & "<vProd>" & FormataValorTexto((.Fields("Qtde").Value * .Fields("Preço").Value), 2) & "</vProd>"
      strLinha = strLinha & "<cEANTrib>SEM GTIN</cEANTrib>"
      strLinha = strLinha & "<uTrib>" & sTrataDado & "</uTrib>"
      strLinha = strLinha & "<qTrib>" & FormataValorTexto(.Fields("Qtde").Value, 4) & "</qTrib>"
      strLinha = strLinha & "<vUnTrib>" & FormataValorTexto(strValorItem, 4) & "</vUnTrib>"

      ' Frete
      If dbl_valor_frete_rateado <> 0 Then
          s_vFreteAux = FormataValorTexto(dbl_valor_frete_rateado, 2)
          s_vFreteAux = Replace(s_vFreteAux, ".", ",")
          dbl_valor_frete_rateado = CDbl(s_vFreteAux)
          dbl_vFreteAcumulado = dbl_vFreteAcumulado + dbl_valor_frete_rateado

          If dbl_vFreteAcumulado > dbl_vFrete Then
              dbl_FreteAuxDif = dbl_vFreteAcumulado - dbl_vFrete
              strLinha = strLinha & "<vFrete>" & FormataValorTexto(dbl_valor_frete_rateado - dbl_FreteAuxDif, 2) & "</vFrete>"
          ElseIf dbl_vFreteAcumulado < dbl_vFrete Then
              If iContaLoop + 1 = iContaRS Then
                  dbl_FreteAuxDif = dbl_vFrete - dbl_vFreteAcumulado
                  strLinha = strLinha & "<vFrete>" & FormataValorTexto(dbl_valor_frete_rateado + dbl_FreteAuxDif, 2) & "</vFrete>"
              Else
                  strLinha = strLinha & "<vFrete>" & FormataValorTexto(dbl_valor_frete_rateado, 2) & "</vFrete>"
              End If
          Else
              strLinha = strLinha & "<vFrete>" & FormataValorTexto(dbl_valor_frete_rateado, 2) & "</vFrete>"
          End If
      End If
        
      If dbl_valor_seguro_rateado <> 0 Then
          strLinha = strLinha & "<vSeg>" & FormataValorTexto(dbl_valor_seguro_rateado, 2) & "</vSeg>"
      End If
        
      '=======================================================
      ' Verificar se tem IPI e ou DESCONTO na linha do produto
      Dim vDesconto_verifica As Double
      Dim vIPI_verifica As Double
      vDesconto_verifica = 0
      vIPI_verifica = 0
      If .Fields("IPI").Value > 0 Then
          vDesconto_verifica = ((.Fields("Qtde").Value * .Fields("Preço").Value) * .Fields("Saídas - Produtos.Desconto").Value) / 100
          vIPI_verifica = (((.Fields("Qtde").Value * .Fields("Preço").Value) - vDesconto_verifica) * .Fields("IPI").Value) / 100
      ElseIf .Fields("Saídas - Produtos.Desconto").Value > 0 Then 'Coluna DESCONTO da tabela [Saídas - Produtos]
          vDesconto_verifica = ((.Fields("Qtde").Value * .Fields("Preço").Value) * .Fields("Saídas - Produtos.Desconto").Value) / 100
      End If
      
      If vDesconto_verifica > 0 Then
          m_bln_tem_desconto = True
      End If
      ' Fim verifica se tem IPI
      '=========================================================

      Dim vAuxiliarCalculoVendaDesconto As Double
      Dim sDescontoAuxZ As String
      sDescontoAuxZ = ""
      If ((.Fields("Qtde").Value * .Fields("Preço").Value) + vIPI_verifica - vDesconto_verifica - .Fields("Preço Final").Value) < 0.01 Then
          
          If bTemDescontoNaVenda = True Then
              vAuxiliarCalculoVendaDesconto = ((.Fields("Qtde").Value * .Fields("Preço").Value) * vTemDescontoNaVendaPercentual) / 100
              s_vlSOMA_DO_DESCONTO = FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2)
              s_vlSOMA_DO_DESCONTO = Replace(s_vlSOMA_DO_DESCONTO, ".", ",")
              vlSOMA_DO_DESCONTO = vlSOMA_DO_DESCONTO + CDbl(s_vlSOMA_DO_DESCONTO)
              
              If iContaRS = 1 Then
                  sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
  
                  If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
                    strLinha = strLinha & sDescontoAuxZ
                  End If

              ElseIf iContaRS = iContaLoop + 1 Then
                  Dim s_vSOMA_TOTAL_DESCONTO As String
                  Dim v_vSOMA_TOTAL_DESCONTO As Double
                  Dim v_diferenca_DESCONTO As Double
                  s_vSOMA_TOTAL_DESCONTO = FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2)
                  s_vSOMA_TOTAL_DESCONTO = Replace(s_vSOMA_TOTAL_DESCONTO, ".", ",")
                  v_vSOMA_TOTAL_DESCONTO = CDbl(s_vSOMA_TOTAL_DESCONTO)
                      
                  If vlSOMA_DO_DESCONTO = v_vSOMA_TOTAL_DESCONTO Then
                      sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                  ElseIf vlSOMA_DO_DESCONTO < v_vSOMA_TOTAL_DESCONTO Then
                      v_diferenca_DESCONTO = v_vSOMA_TOTAL_DESCONTO - vlSOMA_DO_DESCONTO
                      sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto + v_diferenca_DESCONTO, 2) & "</vDesc>"
                  ElseIf vlSOMA_DO_DESCONTO > v_vSOMA_TOTAL_DESCONTO Then
                      v_diferenca_DESCONTO = FormataValorTextoComVirgula(vlSOMA_DO_DESCONTO - v_vSOMA_TOTAL_DESCONTO, 2)
                      sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto - v_diferenca_DESCONTO, 2) & "</vDesc>"
                  End If
                  
                  If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
                    strLinha = strLinha & sDescontoAuxZ
                  End If
              Else
                  sDescontoAuxZ = "<vDesc>" & FormataValorTexto(s_vlSOMA_DO_DESCONTO, 2) & "</vDesc>"
              
                  If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
                    strLinha = strLinha & sDescontoAuxZ
                  End If
              
              End If
          ElseIf vDesconto_verifica > 0 Then
              sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vDesconto_verifica, 2) & "</vDesc>"
              strLinha = strLinha & sDescontoAuxZ
          End If
      Else
          m_bln_tem_desconto = True
          
          If bTemDescontoNaVenda = True Then
              vAuxiliarCalculoVendaDesconto = ((.Fields("Qtde").Value * .Fields("Preço").Value) * vTemDescontoNaVendaPercentual) / 100
              s_vlSOMA_DO_DESCONTO = FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2)
              s_vlSOMA_DO_DESCONTO = Replace(s_vlSOMA_DO_DESCONTO, ".", ",")
              vlSOMA_DO_DESCONTO = vlSOMA_DO_DESCONTO + CDbl(s_vlSOMA_DO_DESCONTO)
              
              If iContaRS = 1 Then
                sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"

                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                  strLinha = strLinha & sDescontoAuxZ
                End If

              ElseIf iContaRS = iContaLoop + 1 Then
                  s_vSOMA_TOTAL_DESCONTO = FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2)
                  s_vSOMA_TOTAL_DESCONTO = Replace(s_vSOMA_TOTAL_DESCONTO, ".", ",")
                  v_vSOMA_TOTAL_DESCONTO = CDbl(s_vSOMA_TOTAL_DESCONTO)
                      
                  If vlSOMA_DO_DESCONTO = v_vSOMA_TOTAL_DESCONTO Then
                      sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                  ElseIf vlSOMA_DO_DESCONTO < v_vSOMA_TOTAL_DESCONTO Then
                      v_diferenca_DESCONTO = v_vSOMA_TOTAL_DESCONTO - vlSOMA_DO_DESCONTO
                      sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto + v_diferenca_DESCONTO, 2) & "</vDesc>"
                  ElseIf vlSOMA_DO_DESCONTO > v_vSOMA_TOTAL_DESCONTO Then
                      v_diferenca_DESCONTO = vlSOMA_DO_DESCONTO - v_vSOMA_TOTAL_DESCONTO
                      sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto - v_diferenca_DESCONTO, 2) & "</vDesc>"
                  End If
                  
                  If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                    strLinha = strLinha & sDescontoAuxZ
                  End If
              Else
                  sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
              
                  If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                    strLinha = strLinha & sDescontoAuxZ
                  End If
              End If
              
          Else
              sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value), 2) & "</vDesc>"
          
              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                strLinha = strLinha & sDescontoAuxZ
              End If
          End If
      End If
        
      If Not IsNull(.Fields("Desp_Acessorias").Value) And .Fields("Desp_Acessorias").Value <> "" And .Fields("Desp_Acessorias").Value <> "0" Then
          strLinha = strLinha & "<vOutro>" & FormataValorTexto(.Fields("Desp_Acessorias").Value, 2) & "</vOutro>"
          vDespAcessorias = vDespAcessorias + CDbl(.Fields("Desp_Acessorias").Value)
      End If

      strLinha = strLinha & "<indTot>" & "1" & "</indTot>"

      If nfe_xPed_nItemPed = True Then
          If gSequenciaSaidas = Sequencia Then
              Dim iContadorXPed As Integer
              For iContadorXPed = 0 To gProdutosXPed_nItemPedContador - 1
                  If gProdutosXPed_nItemPed(iContadorXPed, 0) = .Fields("Código").Value Then
                      strLinha = strLinha & "<xPed>" & gProdutosXPed_nItemPed(iContadorXPed, 1) & "</xPed>"
                      strLinha = strLinha & "<nItemPed>" & gProdutosXPed_nItemPed(iContadorXPed, 2) & "</nItemPed>"
                      Exit For
                  End If
              Next
          End If
      End If
        
      strLinha = strLinha & "</prod>"
      AddLinha clcArquivo, strLinha
      
      '---------------------------------------------------------------------
      ValorBaseCalculoICMS = .Fields("Qtde").Value * .Fields("Preço").Value
      
      If Not IsNull(.Fields("Percentual Icm Saida").Value) Then
        AliquotaICMS_Saida = .Fields("Percentual Icm Saida").Value
      Else
        AliquotaICMS_Saida = 0
      End If
      
      If Not IsNull(.Fields("Percentual Icm Entrada").Value) Then
        AliquotaICMS_Entrada = .Fields("Percentual Icm Entrada").Value
      Else
        AliquotaICMS_Entrada = 0
      End If
      
      If bFINALIDADE_NFE_DEVOLUCAO = True Then
        vTotal_Desconto_ICMS_vNF = vTotal_Desconto_ICMS_vNF + (ValorBaseCalculoICMS * (AliquotaICMS_Entrada / 100))
      Else
        vTotal_Desconto_ICMS_vNF = vTotal_Desconto_ICMS_vNF + (ValorBaseCalculoICMS * (AliquotaICMS_Saida / 100))
      End If
      
      If Not IsNull(.Fields("Percentual_ICMSST_Entrada").Value) Then
        AliquotaICMSST_Entrada = .Fields("Percentual_ICMSST_Entrada").Value
      Else
        AliquotaICMSST_Entrada = 0
      End If
      
      If Not IsNull(.Fields("BaseCalculoICMSST_Entrada").Value) Then
        ValorBCICMSST_Entrada = .Fields("BaseCalculoICMSST_Entrada").Value * .Fields("Qtde").Value
      Else
        ValorBCICMSST_Entrada = 0
      End If
  
      If Not IsNull(.Fields("Percentual_ICMSST_Saida").Value) Then
        AliquotaICMSST_Saida = .Fields("Percentual_ICMSST_Saida").Value
      Else
        AliquotaICMSST_Saida = 0
      End If
      
      If Not IsNull(.Fields("BaseCalculoICMSST_Saida").Value) Then
        ValorBCICMSST_Saida = .Fields("BaseCalculoICMSST_Saida").Value * .Fields("Qtde").Value
      Else
        ValorBCICMSST_Saida = 0
      End If
      
      If Not IsNull(.Fields("Percentual IPI").Value) Then
        PercentualIPI_Saida = .Fields("Percentual IPI").Value
      Else
        PercentualIPI_Saida = 0
      End If
      
      If Not IsNull(.Fields("Percentual_IPI_Entrada").Value) Then
        PercentualIPI_Entrada = .Fields("Percentual_IPI_Entrada").Value
      Else
        PercentualIPI_Entrada = 0
      End If
      
      '14/03/2023 - Pablo
      If bo_CalcularIPI = True And bo_SomarIPIBC = True Then
        Dim p_tmp_Desconto As Double
        p_tmp_Desconto = 0
        If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
          p_tmp_Desconto = CDbl(sDescontoAuxZ_2)
        End If
        
        Dim p_tmp_pIPI As Double  'Aliquota do IPI
        If bFINALIDADE_NFE_DEVOLUCAO = True And .Fields("Percentual_IPI_Entrada").Value > 0 Then ' IPI para OPERAÇÃO DE DEVOLUÇÃO
          p_tmp_pIPI = PercentualIPI_Entrada
        ElseIf bFINALIDADE_NFE_DEVOLUCAO = False And .Fields("Percentual IPI").Value > 0 Then ' IPI para OPERAÇÃO DE SAIDA NORMAL
          p_tmp_pIPI = PercentualIPI_Saida
        End If
        
        p_tmp_vIPI = Truncate((p_tmp_pIPI * (ValorBaseCalculoICMS - p_tmp_Desconto)) / 100, 2)
        
        ValorBaseCalculoICMS = ValorBaseCalculoICMS + p_tmp_vIPI
      Else
        p_tmp_vIPI = 0
      End If
      
      '10/04/2023 - Pablo
      'cálculo ValorICMS_ST para consumidor final (saída)
      If bFINALIDADE_NFE_DEVOLUCAO = False Then
          ValorICMS_ST = ValorBCICMSST_Saida * AliquotaICMSST_Saida / 100
          v_vlSOMA_DO_ICMS_ST_BaseDeCalculo = v_vlSOMA_DO_ICMS_ST_BaseDeCalculo + ValorBCICMSST_Saida
          v_vlSOMA_DO_ICMS_ST = v_vlSOMA_DO_ICMS_ST + ValorICMS_ST
      End If
      
      ' AQUI CONTA LOOP
      iContaLoop = iContaLoop + 1

      strLinha = "<imposto>"
        
      If Not IsNull(.Fields("Valor_Aprox_Impostos").Value) And .Fields("Valor_Aprox_Impostos").Value <> 0 Then
          strLinha = strLinha & "<vTotTrib>" & FormataValorTexto(.Fields("Valor_Aprox_Impostos").Value, 2) & "</vTotTrib>"
          
          s_vTotTrib = FormataValorTexto(.Fields("Valor_Aprox_Impostos").Value, 2)
          s_vTotTrib = Replace(s_vTotTrib, ".", ",")
          vl_vTotTrib = vl_vTotTrib + CDbl(s_vTotTrib)
      End If
      AddLinha clcArquivo, strLinha
        
      ' ICMS
      strLinha = "<ICMS>"
      AddLinha clcArquivo, strLinha
      
      Dim v_pMVAST_01 As Double
      Dim s_pMVAST_01 As String
      Dim v_pMVAST_02 As Double
      Dim s_pMVAST_02 As String
      Dim v_pMVAST As Double
      
      Dim sTAG_vBC As String
      Dim sTAG_pICMS As String
      Dim sTAG_vICMS As String
      
      sDescontoAuxZ_2 = "0"
   
      If blnSimplesNacional = False Then
          ' ======================================================================
          ' Empresa LUCRO REAL / LUCRO PRESUMIDO

          Select Case Right(strTributacao, 2)
            ' ==============================================================================================
            ' 00
            Case "00"
              'Tributos Incidentes no Produto ou no serviço
              strLinha = strLinha & "<ICMS00>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(strTributacao, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"

              ' ======================================================================
              ' Tratamento ValorBaseCalculoICMS
              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                  iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                  sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                  sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                  sTAG_vBC = FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2)
                  
                  ValorICMS = (ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2) * AliquotaICMS_Entrada) / 100
                  
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2)
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                  v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
              Else
                  sTAG_vBC = FormataValorTexto(ValorBaseCalculoICMS, 2)
                  
                  ValorICMS = (ValorBaseCalculoICMS * AliquotaICMS_Entrada) / 100
                 
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS, 2)
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                  v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
              End If
              ' ======================================================================
              
              If bFINALIDADE_NFE_DEVOLUCAO = True Then
                  AliquotaICMS_Entrada = .Fields("Percentual Icm Entrada").Value
                  sTAG_pICMS = FormataValorTexto(.Fields("Percentual Icm Entrada").Value, 2)
                  ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Entrada & "")) / 100
                  
                  If AliquotaICMS_Entrada > 0 Then
                      bTEM_ICMS_3179a = True
                  End If
              
                  If m_bln_tem_desconto = True Or bTemDescontoNaVenda = True Then
                      If sDescontoAuxZ_2 = "" Then
                          sDescontoAuxZ_2 = "0"
                      End If
                  
                      s_vlSOMA_DO_ICMS = ((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * AliquotaICMS_Entrada / 100)
                      s_vlSOMA_DO_ICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  Else
                      s_vlSOMA_DO_ICMS = FormataValorTexto(ValorICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  End If
              ElseIf bFINALIDADE_NFE_DEVOLUCAO = False Then
                  If bo_AliquotaICMS_interestadual = False Then
                      'VENDA ESTADUAL
                      sTAG_pICMS = FormataValorTexto(AliquotaICMS_Saida, 2)
                      
                      If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                          ValorICMS = (ValorBaseCalculoICMS - sDescontoAuxZ_2) * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      Else
                          ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      End If
                  Else
                      'VENDA INTERESTADUAL
                      sTAG_pICMS = FormataValorTexto(aliquotaICMS_tab_ICMS_PERC_ESTADOS, 2)
                      AliquotaICMS_Saida = aliquotaICMS_tab_ICMS_PERC_ESTADOS
                      
                      If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                          ValorICMS = (ValorBaseCalculoICMS - sDescontoAuxZ_2) * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      Else
                          ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      End If
                  End If

                  If AliquotaICMS_Saida > 0 Then
                      bTEM_ICMS_3179a = True
                  End If
    
                  If m_bln_tem_desconto = True Or bTemDescontoNaVenda = True Then
                      If sDescontoAuxZ_2 = "" Then
                          sDescontoAuxZ_2 = "0"
                      End If
    
                      s_vlSOMA_DO_ICMS = ((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * AliquotaICMS_Saida / 100)
                      s_vlSOMA_DO_ICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
    
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  Else
                      s_vlSOMA_DO_ICMS = FormataValorTexto(ValorICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  End If
              End If
              
              strLinha = strLinha & "<vBC>" & sTAG_vBC & "</vBC>"
              strLinha = strLinha & "<pICMS>" & sTAG_pICMS & "</pICMS>"
              strLinha = strLinha & "<vICMS>" & sTAG_vICMS & "</vICMS>"
              
              'If bFINALIDADE_NFE_DEVOLUCAO = False Then
              '    ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              '    If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
              '        vBC_fcp = ValorBaseCalculoICMS
    
              '        vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
              '        strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
              '        strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
              '        vFCP_registroTotal = vFCP_registroTotal + CDbl(Replace(FormataValorTexto(vResultado, 2), ".", ","))
              '    End If
              'End If
              
              strLinha = strLinha & "</ICMS00>"
              AddLinha clcArquivo, strLinha
            
            
            ' ==============================================================================================
            ' 10
            Case "10"

              'ICMS Normal e ST
              strLinha = strLinha & "<ICMS10>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(strTributacao, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"

              ' ======================================================================
              ' Tratamento ValorBaseCalculoICMS
              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                  iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                  sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                  sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                  sTAG_vBC = FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2)
                  
                  ValorICMS = (ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2) * AliquotaICMS_Entrada) / 100
                  
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2)
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                  v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
              Else
                  sTAG_vBC = FormataValorTexto(ValorBaseCalculoICMS, 2)
                  
                  ValorICMS = (ValorBaseCalculoICMS * AliquotaICMS_Entrada) / 100
                 
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS, 2)
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                  v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
              End If
              ' ======================================================================
              
              If bFINALIDADE_NFE_DEVOLUCAO = True Then
                  AliquotaICMS_Entrada = .Fields("Percentual Icm Entrada").Value
                  sTAG_pICMS = FormataValorTexto(.Fields("Percentual Icm Entrada").Value, 2)
                  ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Entrada & "")) / 100
                  
                  If AliquotaICMS_Entrada > 0 Then
                      bTEM_ICMS_3179a = True
                  End If
              
                  If m_bln_tem_desconto = True Or bTemDescontoNaVenda = True Then
                      If sDescontoAuxZ_2 = "" Then
                          sDescontoAuxZ_2 = "0"
                      End If
                  
                      s_vlSOMA_DO_ICMS = ((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * AliquotaICMS_Entrada / 100)
                      s_vlSOMA_DO_ICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  Else
                      s_vlSOMA_DO_ICMS = FormataValorTexto(ValorICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  End If
              ElseIf bFINALIDADE_NFE_DEVOLUCAO = False Then
                  If bo_AliquotaICMS_interestadual = False Then
                      'VENDA ESTADUAL
                      sTAG_pICMS = FormataValorTexto(AliquotaICMS_Saida, 2)
                      
                      If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                          ValorICMS = (ValorBaseCalculoICMS - sDescontoAuxZ_2) * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      Else
                          ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      End If
                  Else
                      'VENDA INTERESTADUAL
                      sTAG_pICMS = FormataValorTexto(aliquotaICMS_tab_ICMS_PERC_ESTADOS, 2)
                      AliquotaICMS_Saida = aliquotaICMS_tab_ICMS_PERC_ESTADOS
                      
                      If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                          ValorICMS = (ValorBaseCalculoICMS - sDescontoAuxZ_2) * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      Else
                          ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      End If
                  End If

                  If AliquotaICMS_Saida > 0 Then
                      bTEM_ICMS_3179a = True
                  End If
    
                  If m_bln_tem_desconto = True Or bTemDescontoNaVenda = True Then
                      If sDescontoAuxZ_2 = "" Then
                          sDescontoAuxZ_2 = "0"
                      End If
    
                      s_vlSOMA_DO_ICMS = ((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * AliquotaICMS_Saida / 100)
                      s_vlSOMA_DO_ICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
    
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  Else
                      s_vlSOMA_DO_ICMS = FormataValorTexto(ValorICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  End If
              End If

              strLinha = strLinha & "<vBC>" & sTAG_vBC & "</vBC>"
              strLinha = strLinha & "<pICMS>" & sTAG_pICMS & "</pICMS>"
              strLinha = strLinha & "<vICMS>" & sTAG_vICMS & "</vICMS>"


              If bFINALIDADE_NFE_DEVOLUCAO = True Then
                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                      vAuxiliar = ValorBCICMSST_Entrada
    
                      If vAuxiliar = 0 Then
                          vBC_fcp = ValorBaseCalculoICMS
    
                          vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                          strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                          strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                          vFCP_registroTotal = vFCP_registroTotal + CDbl(Replace(FormataValorTexto(vResultado, 2), ".", ","))
                      End If
                  End If
  
                  strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
                    
                  ' pMVAST = Percentual da margem de valor Adicionado do ICMS ST da NF-e
                  ' Pode ser informado para os grupos de impostos: ICMS10, ICMS30, ICMS70, ICMS90, ICMSPart, ICMSSN201, ICMSSN202 e ICMSSN900.
                  ' Apenas para NFe que o pMVAST só deve ser preenchido...e quando a modBCST for igual à 4 = MVA.

                  ' Parâmetro definido na tela de Empresa-Filial, aba Nota Fiscal
                  ' Combo cboModDetBaseCalculoIcmsSt
                  If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                      If ValorBCICMSST_Entrada > 0 Then
                          s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Entrada, 2)
                          s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                          v_pMVAST_01 = CDbl(s_pMVAST_01)
                          s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                          s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                          v_pMVAST_02 = CDbl(s_pMVAST_02)
    
                          ' Formula é ((vBCST - vBC) / vBC) * 100
                          v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                          If v_pMVAST < 0 Then
                              v_pMVAST = 100 + v_pMVAST
                          End If
                          strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                      End If
                  End If

                  If bo_CalcularICMS_ST = True And ValorBCICMSST_Entrada > 0 Then
                      strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCST>"
                      strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Entrada, 4) & "</pICMSST>"
    
                      s_vlSOMA_DO_ICMS_ST_BaseDeCalculo = FormataValorTexto(ValorBCICMSST_Entrada, 2)
                      s_vlSOMA_DO_ICMS_ST_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_ST_BaseDeCalculo, ".", ",")
                      v_vlSOMA_DO_ICMS_ST_BaseDeCalculo = v_vlSOMA_DO_ICMS_ST_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_ST_BaseDeCalculo)
    
                      v_vlICMS_ST = (ValorBCICMSST_Entrada * (AliquotaICMSST_Entrada / 100))
                      strLinha = strLinha & "<vICMSST>" & FormataValorTexto(v_vlICMS_ST, 2) & "</vICMSST>"
                      v_vlSOMA_DO_ICMS_ST = v_vlSOMA_DO_ICMS_ST + v_vlICMS_ST
                  Else
                      strLinha = strLinha & "<vBCST>0.00</vBCST>"
                      strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                      strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                  End If

                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                      vAuxiliar = ValorBCICMSST_Entrada
    
                      If vAuxiliar <> 0 Then
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                          strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                          strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                          vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                      End If
                  End If
              Else
              
              
                  ' ===================================================
                  ' SAIDA NORMAL

                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                      vAuxiliar = ValorBCICMSST_Saida
    
                      If vAuxiliar = 0 Then
                          vBC_fcp = ValorBaseCalculoICMS
    
                          vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                          strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                          strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                          vFCP_registroTotal = vFCP_registroTotal + CDbl(Replace(FormataValorTexto(vResultado, 2), ".", ","))
                      End If
                  End If

                  strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"

                  ' pMVAST = Percentual da margem de valor Adicionado do ICMS ST da NF-e
                  ' Pode ser informado para os grupos de impostos: ICMS10, ICMS30, ICMS70, ICMS90, ICMSPart, ICMSSN201, ICMSSN202 e ICMSSN900.
                  ' Apenas para NFe que o pMVAST só deve ser preenchido...e quando a modBCST for igual à 4 = MVA.
    
                  ' Parâmetro definido na tela de Empresa-Filial, aba Nota Fiscal
                  ' Combo cboModDetBaseCalculoIcmsSt
                  If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                      If ValorBCICMSST_Saida > 0 Then
                          s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Saida, 2)
                          s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                          v_pMVAST_01 = CDbl(s_pMVAST_01)
                          s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                          s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                          v_pMVAST_02 = CDbl(s_pMVAST_02)
  
                          ' Formula é ((vBCST - vBC) / vBC) * 100
                          v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                          If v_pMVAST < 0 Then
                              v_pMVAST = 0
                          End If
                          strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                      End If
                  Else
                      strLinha = strLinha & "<pMVAST>0.0000</pMVAST>"
                  End If

                  If bo_CalcularICMS_ST = True And ValorBCICMSST_Saida > 0 Then
                      strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCST>"
                      strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Saida, 4) & "</pICMSST>"
  
                      s_vlSOMA_DO_ICMS_ST_BaseDeCalculo = FormataValorTexto(ValorBCICMSST_Saida, 2)
                      s_vlSOMA_DO_ICMS_ST_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_ST_BaseDeCalculo, ".", ",")
                      v_vlSOMA_DO_ICMS_ST_BaseDeCalculo = v_vlSOMA_DO_ICMS_ST_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_ST_BaseDeCalculo)
  
                      Dim sVlAux As String
                      sVlAux = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                      sVlAux = Replace(sVlAux, ".", ",")
                      v_vlICMS_ST = (ValorBCICMSST_Saida * (AliquotaICMSST_Saida / 100)) ' - CDbl(sVlAux)
                      strLinha = strLinha & "<vICMSST>" & FormataValorTexto(v_vlICMS_ST, 2) & "</vICMSST>"
                      v_vlSOMA_DO_ICMS_ST = v_vlSOMA_DO_ICMS_ST + v_vlICMS_ST
                  Else
                      strLinha = strLinha & "<vBCST>0.00</vBCST>"
                      strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                      strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                  End If

                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                      vAuxiliar = ValorBCICMSST_Saida
    
                      ' COLOCAR GRUPO 3111 ou 3112
                      If vAuxiliar <> 0 Then
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          ' [3112;vBCFCPST;pFCPST;vFCPST]
                          strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPST>"
                          strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                          strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                          vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                      End If
                  End If
              End If

              strLinha = strLinha & "</ICMS10>"
              AddLinha clcArquivo, strLinha

            ' ==============================================================================================
            ' 20
            Case "20"
                'Dados do ICMS da operação
                strLinha = strLinha & "<ICMS20>"
                strLinha = strLinha & "<orig>" & Left(rsSaidasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
                strLinha = strLinha & "<CST>" & Right(rsSaidasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
                strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
                
                If Left(.Fields("CodigoNBM").Value, 4) = "8703" Then
                    'NCM para venda de carro usado
                    strLinha = strLinha & "<pRedBC>95.0000</pRedBC>"
                    ValorBaseCalculoICMS = ValorBaseCalculoICMS * 0.05
                    strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS, 2) & "</vBC>"
                Else
                    strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBC>"
                    strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS, 2) & "</vBC>"
                End If

                If bFINALIDADE_NFE_DEVOLUCAO = True Then
                    strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS_Entrada, 4) & "</pICMS>"
                Else
                
                    ' Verificar UF_destinatario
                    If str_FCP_UF_destinatario = "PR" Then
                        'Dentro do Estado
                        strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS_Saida, 4) & "</pICMS>"
                        ValorICMS = ((.Fields("Qtde").Value * ValorBaseCalculoICMS) * AliquotaICMS_Saida) / 100
                    Else
                        'Fora do Estado
                        If isIEIsento = True Then
                            strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS_Saida, 4) & "</pICMS>"
                            ValorICMS = ((.Fields("Qtde").Value * ValorBaseCalculoICMS) * AliquotaICMS_Saida) / 100
                        Else
                            If str_FCP_UF_destinatario = "MG" Or str_FCP_UF_destinatario = "RS" Or str_FCP_UF_destinatario = "RJ" Or _
                                str_FCP_UF_destinatario = "SC" Or str_FCP_UF_destinatario = "SP" Then
                                
                                strLinha = strLinha & "<pICMS>12.0000</pICMS>"
                                ValorICMS = ((.Fields("Qtde").Value * ValorBaseCalculoICMS) * 12) / 100
                            Else
                                strLinha = strLinha & "<pICMS>7.0000</pICMS>"
                                ValorICMS = ((.Fields("Qtde").Value * ValorBaseCalculoICMS) * 7) / 100
                            End If
                        End If
                    End If
                End If

                strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
  
                bTEM_ICMS_3179a = True
                
                s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto((.Fields("Qtde").Value * ValorBaseCalculoICMS), 2)
                s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)

                s_vlSOMA_DO_ICMS = ValorICMS
                s_vlSOMA_DO_ICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
  
                ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                    vBC_fcp = ValorBaseCalculoICMS
                    vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
  
                    ' [3121;vBCFCP;pFCP;vFCP]
                    strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                    strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                    strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                    vFCP_registroTotal = vFCP_registroTotal + CDbl(Replace(FormataValorTexto(vResultado, 2), ".", ","))
                End If
  
                If rsSaidasProdutos.Fields("MotivoDesoneracaoICMS").Value = 3 Or rsSaidasProdutos.Fields("MotivoDesoneracaoICMS").Value = 9 Or rsSaidasProdutos.Fields("MotivoDesoneracaoICMS").Value = 12 Then
                    strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(rsSaidasProdutos.Fields("ValorICMSDesonerado").Value, 2) & "</vICMSDeson>" 'Valor ICms Desonerado
                    strLinha = strLinha & "<motDesICMS>" & rsSaidasProdutos.Fields("MotivoDesoneracaoICMS").Value & "</motDesICMS>" 'Código do motivo Icms Desonerado
                Else
                    strLinha = strLinha & "<vICMSDeson></vICMSDeson><motDesICMS></motDesICMS>"
                End If
                strLinha = strLinha & "</ICMS20>"
                AddLinha clcArquivo, strLinha
            
            ' ==============================================================================================
            ' 30
            Case "30"
              'Dados do icms da operação
              strLinha = strLinha & "<ICMS30>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(strTributacao, 2) & "</CST>"
              strLinha = strLinha & "</ICMS30>"
              AddLinha clcArquivo, strLinha

            ' ==============================================================================================
            ' 40, 41 e 50
            Case "40", "41", "50"
              strLinha = strLinha & "<ICMS40>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(strTributacao, 2) & "</CST>"
              strLinha = strLinha & "</ICMS40>"
              AddLinha clcArquivo, strLinha

            ' ==============================================================================================
            ' 60
            Case "60"
              'Dados do icms da operação
              strLinha = strLinha & "<ICMS60>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(strTributacao, 2) & "</CST>"
              If str_CFOP = "6409" Then
                strLinha = strLinha & "<vBCSTRet>0.00</vBCSTRet>"
                strLinha = strLinha & "<pST>0.00</pST>"
                strLinha = strLinha & "<vICMSSubstituto>0.00</vICMSSubstituto>"
                strLinha = strLinha & "<vICMSSTRet>0.00</vICMSSTRet>"
              End If
              strLinha = strLinha & "</ICMS60>"
              AddLinha clcArquivo, strLinha

            Case "90"
              'Dados do icms da operação
              strLinha = strLinha & "<ICMS90>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(strTributacao, 2) & "</CST>"
              strLinha = strLinha & "</ICMS90>"
              AddLinha clcArquivo, strLinha
              
            Case Else
              'gerar erro
              strLinha = strLinha & "<ICMS_ERRO>SEM CST</ICMS_ERRO>"
              AddLinha clcArquivo, strLinha

          End Select
      Else
          ' ========================================================================================
          ' Empresa SIMPLES NACIONAL
          
          If dblPercentualReducaoBC_SN <> 0 Then
             dblBaseCalculoSimplesItem = ((.Fields("Qtde").Value * .Fields("Preço").Value) * dblPercentualReducaoBC_SN)
          Else
             dblBaseCalculoSimplesItem = (.Fields("Qtde").Value * .Fields("Preço").Value)
          End If
          
          If bFINALIDADE_NFE_DEVOLUCAO = True Then
              If str_CSO_Produto = "" And CSO_NaturezaOperacao <> "" Then
                  str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
              End If
          End If
          
          ' ***************************************************************
          ' Tratamento de inconsistência no cadastro de CSO
          If IsNull(str_CSO_Produto) Or str_CSO_Produto = "" Then
              While str_CSO_Produto = ""
                  sResp = InputBox("O cadastro do produto " & vbCrLf & "CÓDIGO: " & rsSaidasProdutos.Fields("Código").Value & vbCrLf & "NOME: " & nomeProduto & vbCrLf & " está INCOMPLETO ! " & vbCrLf & vbCrLf & "Pontualmente lhe ajudaremos a seguir com a emissão desta NF-e. Porém, você deverá COMPLETAR este cadastro posteriormente." & vbCrLf & vbCrLf & "Agora, Digite o CSO correto vinculado a este produto " & vbCrLf & "Ex: 101, 102, 103, 400, 900, entre outros", "Cadastro INCOMPLETO - Vincular CSO")
      
                  If IsNull(sResp) Then
                      str_CSO_Produto = ""
                  ElseIf Not IsNumeric(sResp) Then
                      str_CSO_Produto = ""
                  ElseIf Len(Trim(sResp)) <> 3 Then
                      str_CSO_Produto = ""
                  Else
                      str_CSO_Produto = sResp
                  End If
              Wend
          End If
          ' ***************************************************************
          
          Select Case str_CSO_Produto
          
            Case "101" 'Tributada pelo Simples Nacional com permissão de crédito
              strLinha = strLinha & "<ICMSSN101>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 4) & "</pCredSN>"
              strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
              strLinha = strLinha & "</ICMSSN101>"
              AddLinha clcArquivo, strLinha

            Case "102", "103", "300", "400"
              strLinha = strLinha & "<ICMSSN102>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "</ICMSSN102>"
              AddLinha clcArquivo, strLinha
              
            Case "201" 'Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por Substituição
              strLinha = strLinha & "<ICMSSN201>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
              
              If bFINALIDADE_NFE_DEVOLUCAO = True Then
                  If iObterTributosProduto_EntradaOuSaida = 1 Then
                      ' pMVAST = Percentual da margem de valor Adicionado do ICMS ST da NF-e
                      ' Pode ser informado para os grupos de impostos: ICMS10, ICMS30, ICMS70, ICMS90, ICMSPart, ICMSSN201, ICMSSN202 e ICMSSN900.
                      ' Apenas para NFe que o pMVAST só deve ser preenchido...e quando a modBCST for igual à 4 = MVA.
        
                      ' Parâmetro definido na tela de Empresa-Filial, aba Nota Fiscal
                      ' Combo cboModDetBaseCalculoIcmsSt
                      If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                          If ValorBCICMSST_Entrada > 0 Then
                              s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Entrada, 2)
                              s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                              v_pMVAST_01 = CDbl(s_pMVAST_01)
                              s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                              s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                              v_pMVAST_02 = CDbl(s_pMVAST_02)
        
                              ' Formula é ((vBCST - vBC) / vBC) * 100
                              v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                              strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                          End If
                      Else
                          strLinha = strLinha & "<pMVAST>0.0000</pMVAST>"
                      End If
                  
                      strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
                      
                      If bo_CalcularICMS_ST = True Then
                          strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCST>"
                          strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Entrada, 4) & "</pICMSST>"
                          strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
                      Else
                          strLinha = strLinha & "<vBCST>0.00</vBCST>"
                          strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                          strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                      End If
                      
                      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                      If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                          vAuxiliar = ValorBCICMSST_Entrada
        
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCFCPST>"
                          strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                          strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                          vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                      End If
                  Else
                      If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                          If ValorBCICMSST_Saida > 0 Then
                              s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Saida, 2)
                              s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                              v_pMVAST_01 = CDbl(s_pMVAST_01)
                              s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                              s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                              v_pMVAST_02 = CDbl(s_pMVAST_02)
        
                              ' Formula é ((vBCST - vBC) / vBC) * 100
                              v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                              strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                          End If
                      Else
                          strLinha = strLinha & "<pMVAST>0.0000</pMVAST>"
                      End If
                  
                      strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
                      
                      If bo_CalcularICMS_ST = True Then
                          strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCST>"
                          strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Saida, 4) & "</pICMSST>"
                          strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
                      Else
                          strLinha = strLinha & "<vBCST>0.00</vBCST>"
                          strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                          strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                      End If
                      
                      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                      If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                          vAuxiliar = ValorBCICMSST_Saida
        
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPST>"
                          strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                          strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                          vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                      End If
                  End If
              Else
                  If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                      If ValorBCICMSST_Saida > 0 Then
                          s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Saida, 2)
                          s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                          v_pMVAST_01 = CDbl(s_pMVAST_01)
                          s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                          s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                          v_pMVAST_02 = CDbl(s_pMVAST_02)
    
                          ' Formula é ((vBCST - vBC) / vBC) * 100
                          v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                          strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                      End If
                  Else
                      strLinha = strLinha & "<pMVAST>0.0000</pMVAST>"
                  End If
              
                  strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
                  
                  If bo_CalcularICMS_ST = True Then
                      strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCST>"
                      strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Saida, 4) & "</pICMSST>"
                      strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
                  Else
                      strLinha = strLinha & "<vBCST>0.00</vBCST>"
                      strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                      strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                  End If
                  
                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                      vAuxiliar = ValorBCICMSST_Saida
    
                      vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                      strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPST>"
                      strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                      strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                      vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                  End If
              End If
                
              strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 4) & "</pCredSN>"
              strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
              strLinha = strLinha & "</ICMSSN201>"
              AddLinha clcArquivo, strLinha
              
            Case "202", "203"
              '202- Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por Substituição Tributária
              '203- Isenção do ICMS nos Simples Nacional para faixa de receita bruta e com cobrança do ICMS por Substituição Tributária
              strLinha = strLinha & "<ICMSSN202>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
              
              If bFINALIDADE_NFE_DEVOLUCAO = True Then
                  If iObterTributosProduto_EntradaOuSaida = 1 Then
                      ' pMVAST = Percentual da margem de valor Adicionado do ICMS ST da NF-e
                      ' Pode ser informado para os grupos de impostos: ICMS10, ICMS30, ICMS70, ICMS90, ICMSPart, ICMSSN201, ICMSSN202 e ICMSSN900.
                      ' Apenas para NFe que o pMVAST só deve ser preenchido...e quando a modBCST for igual à 4 = MVA.
        
                      ' Parâmetro definido na tela de Empresa-Filial, aba Nota Fiscal
                      ' Combo cboModDetBaseCalculoIcmsSt
                      If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                          If ValorBCICMSST_Entrada > 0 Then
                              s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Entrada, 2)
                              s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                              v_pMVAST_01 = CDbl(s_pMVAST_01)
                              s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                              s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                              v_pMVAST_02 = CDbl(s_pMVAST_02)
        
                              ' Formula é ((vBCST - vBC) / vBC) * 100
                              v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                              strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                          End If
                      Else
                          strLinha = strLinha & "<pMVAST>0.0000</pMVAST>"
                      End If
                  
                      strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
                      
                      If bo_CalcularICMS_ST = True Then
                          strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCST>"
                          strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Entrada, 4) & "</pICMSST>"
                          strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
                      Else
                          strLinha = strLinha & "<vBCST>0.00</vBCST>"
                          strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                          strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                      End If
                      
                      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                      If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                          vAuxiliar = ValorBCICMSST_Entrada
        
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCFCPST>"
                          strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                          strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                          vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                      End If
                  Else
                      If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                          If ValorBCICMSST_Saida > 0 Then
                              s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Saida, 2)
                              s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                              v_pMVAST_01 = CDbl(s_pMVAST_01)
                              s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                              s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                              v_pMVAST_02 = CDbl(s_pMVAST_02)
        
                              ' Formula é ((vBCST - vBC) / vBC) * 100
                              v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                              strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                          End If
                      Else
                          strLinha = strLinha & "<pMVAST>0.0000</pMVAST>"
                      End If
                  
                      strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
                      
                      If bo_CalcularICMS_ST = True Then
                          strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCST>"
                          strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Saida, 4) & "</pICMSST>"
                          strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
                      Else
                          strLinha = strLinha & "<vBCST>0.00</vBCST>"
                          strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                          strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                      End If
                      
                      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                      If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                          vAuxiliar = ValorBCICMSST_Saida
        
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPST>"
                          strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                          strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                          vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                      End If
                  End If
              Else
                  If rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value = 4 Then
                      If ValorBCICMSST_Saida > 0 Then
                          s_pMVAST_01 = FormataValorTexto(ValorBCICMSST_Saida, 2)
                          s_pMVAST_01 = Replace(s_pMVAST_01, ".", ",")
                          v_pMVAST_01 = CDbl(s_pMVAST_01)
                          s_pMVAST_02 = FormataValorTexto(ValorBaseCalculoICMS, 2)
                          s_pMVAST_02 = Replace(s_pMVAST_02, ".", ",")
                          v_pMVAST_02 = CDbl(s_pMVAST_02)
    
                          ' Formula é ((vBCST - vBC) / vBC) * 100
                          v_pMVAST = ((v_pMVAST_01 - s_pMVAST_02) / s_pMVAST_02) * 100
                          strLinha = strLinha & "<pMVAST>" & FormataValorTexto(v_pMVAST, 4) & "</pMVAST>"
                      End If
                  Else
                      strLinha = strLinha & "<pMVAST>0.0000</pMVAST>"
                  End If
              
                  strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
                  
                  If bo_CalcularICMS_ST = True Then
                      strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCST>"
                      strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Saida, 4) & "</pICMSST>"
                      strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
                  Else
                      strLinha = strLinha & "<vBCST>0.00</vBCST>"
                      strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
                      strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
                  End If
                  
                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                      vAuxiliar = ValorBCICMSST_Saida
    
                      vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                      strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPST>"
                      strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                      strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                      vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                  End If
              
              End If
              
              strLinha = strLinha & "</ICMSSN202>"
              AddLinha clcArquivo, strLinha
              
            Case "500" 'ICMS cobrado anteriormente por substituição tributária (substituído) ou por antecipação
              strLinha = strLinha & "<ICMSSN500>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              
              If bo_CalcularICMS_ST = False Then
                  ValorBCICMSST_Entrada = 0
                  ValorBCICMSST_Saida = 0
              End If
              
              If bFINALIDADE_NFE_DEVOLUCAO = True Then
                  If iObterTributosProduto_EntradaOuSaida = 1 Then
                      strLinha = strLinha & "<vBCSTRet>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCSTRet>"
                      
                      vTesteAux = ValorBCICMSST_Entrada
                      If vTesteAux > 0 Then
                          'pST Exemplo: alíquota da mercadoria na venda ao consumidor final = 18% e 2% de FCP
                          vResultado = AliquotaAdesaoUFDestFCP + AliquotaFCP_Destinatario
                          strLinha = strLinha & "<pST>" & FormataValorTexto(vResultado, 4) & "</pST>"
                      Else
                          strLinha = strLinha & "<pST>0.00</pST>"
                      End If
                      
                      strLinha = strLinha & "<vICMSSTRet>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSSTRet>"
                      
                      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                      If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
        
                          vAuxiliar = ValorBCICMSST_Entrada
        
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCPSTRet>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCFCPSTRet>"
                          strLinha = strLinha & "<pFCPSTRet>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPSTRet>"
                          strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vResultado, 2) & "</vFCPSTRet>"
                          vFCPSTRet_registroTotal = vFCPSTRet_registroTotal + vResultado
                      End If
                  Else
                      strLinha = strLinha & "<vBCSTRet>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCSTRet>"
                      
                      vTesteAux = ValorBCICMSST_Saida
                      If vTesteAux > 0 Then
                          'pST Exemplo: alíquota da mercadoria na venda ao consumidor final = 18% e 2% de FCP
                          vResultado = AliquotaAdesaoUFDestFCP + AliquotaFCP_Destinatario
                          strLinha = strLinha & "<pST>" & FormataValorTexto(vResultado, 4) & "</pST>"
                      Else
                          strLinha = strLinha & "<pST>0.00</pST>"
                      End If
                      
                      strLinha = strLinha & "<vICMSSTRet>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSSTRet>"
                      
                      ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                      If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
        
                          vAuxiliar = ValorBCICMSST_Saida
        
                          vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                          strLinha = strLinha & "<vBCFCPSTRet>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPSTRet>"
                          strLinha = strLinha & "<pFCPSTRet>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPSTRet>"
                          strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vResultado, 2) & "</vFCPSTRet>"
                          vFCPSTRet_registroTotal = vFCPSTRet_registroTotal + vResultado
                      End If
                  End If
              Else
                  strLinha = strLinha & "<vBCSTRet>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCSTRet>"
                  
                  vTesteAux = ValorBCICMSST_Saida
                  If vTesteAux > 0 Then
                      'pST Exemplo: alíquota da mercadoria na venda ao consumidor final = 18% e 2% de FCP
                      vResultado = AliquotaAdesaoUFDestFCP + AliquotaFCP_Destinatario
                      strLinha = strLinha & "<pST>" & FormataValorTexto(vResultado, 4) & "</pST>"
                  Else
                      strLinha = strLinha & "<pST>0.00</pST>"
                  End If
                  
                  strLinha = strLinha & "<vICMSSTRet>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSSTRet>"
                  
                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
    
                      vAuxiliar = ValorBCICMSST_Saida
    
                      vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                      strLinha = strLinha & "<vBCFCPSTRet>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPSTRet>"
                      strLinha = strLinha & "<pFCPSTRet>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPSTRet>"
                      strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vResultado, 2) & "</vFCPSTRet>"
                      vFCPSTRet_registroTotal = vFCPSTRet_registroTotal + vResultado
                  End If
              
              End If
              
              strLinha = strLinha & "</ICMSSN500>"
              AddLinha clcArquivo, strLinha
          
            Case "900" 'Isento
              strLinha = strLinha & "<ICMSSN900>"
              strLinha = strLinha & "<orig>" & Left(strTributacao, 1) & "</orig>"
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"

              ' ==============================================================================
              If bo_CalcularICMS = True And bFINALIDADE_NFE_DEVOLUCAO = True Then
                  ' ======================================================================
                  ' Tratamento ValorBaseCalculoICMS
                  If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                      iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                      sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                      sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                      sTAG_vBC = FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2)
                      
                      '10/03/2023 - Pablo
                      'Procura corrigir o problema do arredondamento com centavos
                      If iObterTributosProduto_EntradaOuSaida = 1 Then
                          ValorICMS = Truncate((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2) * AliquotaICMS_Entrada) / 100, 2)
                      Else
                          ValorICMS = Truncate((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2) * AliquotaICMS_Saida) / 100, 2)
                      End If
                      
                      s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2)
                      s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                      v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
                  Else
                      sTAG_vBC = FormataValorTexto(ValorBaseCalculoICMS, 2)
                      
                      If iObterTributosProduto_EntradaOuSaida = 1 Then
                          ValorICMS = (ValorBaseCalculoICMS * AliquotaICMS_Entrada) / 100
                      Else
                          ValorICMS = (ValorBaseCalculoICMS * AliquotaICMS_Saida) / 100
                      End If
                      ValorICMS = CCur(Round(ValorICMS, 2))
                                           
                      s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS, 2)
                      s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                      v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
                  End If
                  ' ======================================================================
                  
                  If iObterTributosProduto_EntradaOuSaida = 1 Then
                      'Tributos de Entrada
                      If Not IsNull(.Fields("Percentual Icm Entrada").Value) Then
                           AliquotaICMS_Entrada = .Fields("Percentual Icm Entrada").Value
                      Else
                           AliquotaICMS_Entrada = 0
                      End If

                      sTAG_pICMS = FormataValorTexto(AliquotaICMS_Entrada, 2)
                      ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Entrada & "")) / 100
                      
                      If AliquotaICMS_Entrada > 0 Then
                          bTEM_ICMS_3179a = True
                      End If
                  Else
                      'Tributos de Saída
                      If Not IsNull(.Fields("Percentual Icm Saida").Value) Then
                           AliquotaICMS_Saida = .Fields("Percentual Icm Saida").Value
                      Else
                           AliquotaICMS_Saida = 0
                      End If

                      sTAG_pICMS = FormataValorTexto(AliquotaICMS_Saida, 2)
                      ValorICMS = ValorBaseCalculoICMS * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
                      
                      If AliquotaICMS_Saida > 0 Then
                          bTEM_ICMS_3179a = True
                      End If
                  End If
                  ValorICMS = CCur(Round(ValorICMS, 2))
              
                  If m_bln_tem_desconto = True Or bTemDescontoNaVenda = True Then
                      If sDescontoAuxZ_2 = "" Then
                          sDescontoAuxZ_2 = "0"
                      End If
                  
                      If iObterTributosProduto_EntradaOuSaida = 1 Then
                          s_vlSOMA_DO_ICMS = ((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * AliquotaICMS_Entrada / 100)
                      Else
                          s_vlSOMA_DO_ICMS = ((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * AliquotaICMS_Saida / 100)
                      End If
                      
                      s_vlSOMA_DO_ICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  Else
                      s_vlSOMA_DO_ICMS = FormataValorTexto(ValorICMS, 2)
                      s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                      vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                      sTAG_vICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  End If
              
                  strLinha = strLinha & "<vBC>" & sTAG_vBC & "</vBC>"
                  
                  If PercentualReducaoBaseICMS > 0 Then
                      strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBC>"
                  End If
                  
                  strLinha = strLinha & "<pICMS>" & sTAG_pICMS & "</pICMS>"
                  strLinha = strLinha & "<vICMS>" & sTAG_vICMS & "</vICMS>"
              Else
              
                  strLinha = strLinha & "<vBC>0.00</vBC>"
                  strLinha = strLinha & "<pRedBC>0.0000</pRedBC>"
                  strLinha = strLinha & "<pICMS>0.0000</pICMS>"
                  strLinha = strLinha & "<vICMS>0.00</vICMS>"
              End If
              
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
              If MargemValorAdicionado_ICMS > 0 Then
                strLinha = strLinha & "<pMVAST>" & FormataValorTexto(MargemValorAdicionado_ICMS, 4) & "</pMVAST>"
              End If
              
              If bo_CalcularICMS_ST = True And PercentualReducaoBaseICMS > 0 Then
                strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
              End If
              
              If bo_CalcularICMS_ST = True Then
                  If iObterTributosProduto_EntradaOuSaida = 1 Then
                      strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCST>"
                      strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Entrada, 4) & "</pICMSST>"
                  Else
                      strLinha = strLinha & "<vBCST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCST>"
                      strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMSST_Saida, 4) & "</pICMSST>"
                  End If
              Else
                  strLinha = strLinha & "<vBCST>0.00</vBCST>"
                  strLinha = strLinha & "<pICMSST>0.0000</pICMSST>"
              End If
              
              If iObterTributosProduto_EntradaOuSaida = 1 Then
                  s_vlSOMA_DO_ICMS_ST_BaseDeCalculo = FormataValorTexto(ValorBCICMSST_Entrada, 2)
              Else
                  s_vlSOMA_DO_ICMS_ST_BaseDeCalculo = FormataValorTexto(ValorBCICMSST_Saida, 2)
              End If
              s_vlSOMA_DO_ICMS_ST_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_ST_BaseDeCalculo, ".", ",")
              v_vlSOMA_DO_ICMS_ST_BaseDeCalculo = v_vlSOMA_DO_ICMS_ST_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_ST_BaseDeCalculo)
                      
                                    
              If iObterTributosProduto_EntradaOuSaida = 1 Then
                  If bo_CalcularICMS = True And bo_CalcularICMS_ST = True Then
                      If (ValorBCICMSST_Entrada * (AliquotaICMSST_Entrada / 100)) > 0 Then
                          'Calculo é vICMS_ST = vICMS_ST - vICMS
                          v_vlICMS_ST = (ValorBCICMSST_Entrada * (AliquotaICMSST_Entrada / 100)) - CDbl(Replace(sTAG_vICMS, ".", ","))
                      Else
                          v_vlICMS_ST = 0
                      End If
                  Else
                      v_vlICMS_ST = (ValorBCICMSST_Entrada * (AliquotaICMSST_Entrada / 100))
                  End If
              Else
                  If bo_CalcularICMS = True And bo_CalcularICMS_ST = True Then
                      If (ValorBCICMSST_Saida * (AliquotaICMSST_Saida / 100)) > 0 Then
                          'Calculo é vICMS_ST = vICMS_ST - vICMS
                          If sTAG_vICMS = "" Then sTAG_vICMS = "0.00"
                          v_vlICMS_ST = (ValorBCICMSST_Saida * (AliquotaICMSST_Saida / 100)) - CDbl(Replace(sTAG_vICMS, ".", ","))
                      Else
                          v_vlICMS_ST = 0
                      End If
                  Else
                      v_vlICMS_ST = (ValorBCICMSST_Saida * (AliquotaICMSST_Saida / 100))
                  End If
              End If
              
              If bo_CalcularICMS_ST = True Then
                  strLinha = strLinha & "<vICMSST>" & FormataValorTexto(v_vlICMS_ST, 2) & "</vICMSST>"
              Else
                  strLinha = strLinha & "<vICMSST>0.00</vICMSST>"
              End If
              
              v_vlSOMA_DO_ICMS_ST = v_vlSOMA_DO_ICMS_ST + v_vlICMS_ST
              
              If iObterTributosProduto_EntradaOuSaida = 1 Then
                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
    
                      vAuxiliar = ValorBCICMSST_Entrada
    
                      vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                      strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Entrada, 2) & "</vBCFCPST>"
                      strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                      strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                      vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                  End If
              Else
                  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                  If bo_CalcularICMS_ST = True And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
    
                      vAuxiliar = ValorBCICMSST_Saida
    
                      vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                      strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(ValorBCICMSST_Saida, 2) & "</vBCFCPST>"
                      strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                      strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                      vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                  End If
              End If
              
              ' ==============================================================================

              strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 4) & "</pCredSN>"
              strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
              strLinha = strLinha & "</ICMSSN900>"
              AddLinha clcArquivo, strLinha
          End Select
        
      End If 'do Simples Nacional
      
      strLinha = "</ICMS>"
      AddLinha clcArquivo, strLinha


      ' ======================================================================================================
      ' IPI
      If (blnSimplesNacional = False) Or (blnSimplesNacional = True And bFINALIDADE_NFE_DEVOLUCAO = True) Then
          
          ' Destacar IPI para OPERAÇÃO DE DEVOLUÇÃO
          If bo_CalcularIPI = True And bFINALIDADE_NFE_DEVOLUCAO = True And .Fields("Percentual_IPI_Entrada").Value > 0 Then
              strLinha = "<IPI>"
              
              strLinha = strLinha & "<cEnq>999</cEnq>"
              
              strLinha = strLinha & "<IPITrib>"
              strLinha = strLinha & "<CST>99</CST>" ' 00-Entrada com recuperação de IPI crédito 49-Outras entradas 50-Saída tributada 99-Outras saídas
              
              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
                  iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                  sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                  sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                  strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Qtde").Value * .Fields("Preço").Value - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
              Else
                  strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Qtde").Value * .Fields("Preço").Value, 2) & "</vBC>" 'Valor da base de cálculo do ipi
              End If
          
              strLinha = strLinha & "<pIPI>" & Replace(FormataValorTexto(.Fields("Percentual_IPI_Entrada").Value, 2), ",", ".") & "</pIPI>" 'Aliquota do IPI
              
              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
                  strLinha = strLinha & "<vIPI>" & FormataValorTexto(Truncate((.Fields("Percentual_IPI_Entrada").Value * (.Fields("Qtde").Value * .Fields("Preço").Value - CDbl(sDescontoAuxZ_2))) / 100, 2), 2) & "</vIPI>" 'Valor do IPI
              Else
                  strLinha = strLinha & "<vIPI>" & FormataValorTexto(Truncate((.Fields("Percentual_IPI_Entrada").Value * (.Fields("Qtde").Value * .Fields("Preço").Value)) / 100, 2), 2) & "</vIPI>" 'Valor do IPI
              End If
              
              Dim sAuxIPI As String
              Dim lIndice01 As Long
              Dim lIndice02 As Long
          
              lIndice01 = InStr(1, strLinha, "<vIPI>")
              If lIndice01 > 0 Then
                  lIndice02 = InStr(lIndice01, strLinha, "</vIPI>")
                  sAuxIPI = Mid$(strLinha, lIndice01 + 6, lIndice02 - (lIndice01 + 6))
                  sAuxIPI = Replace(sAuxIPI, ".", ",")
                  vlSOMA_DO_IPI = vlSOMA_DO_IPI + CDbl(sAuxIPI)
              End If

              strLinha = strLinha & "</IPITrib>"
              strLinha = strLinha & "</IPI>"
              AddLinha clcArquivo, strLinha
            
          ' Destacar IPI para OPERAÇÃO DE SAIDA NORMAL
          ElseIf bo_CalcularIPI = True And bFINALIDADE_NFE_DEVOLUCAO = False And .Fields("Percentual IPI").Value > 0 Then
              strLinha = "<IPI>"
              
              strLinha = strLinha & "<cEnq>999</cEnq>"
              
              strLinha = strLinha & "<IPITrib>"
              strLinha = strLinha & "<CST>99</CST>" ' 00-Entrada com recuperação de IPI crédito 49-Outras entradas 50-Saída tributada 99-Outras saídas
              
              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
                  iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                  sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                  sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                  strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Qtde").Value * .Fields("Preço").Value - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
              Else
                  strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Qtde").Value * .Fields("Preço").Value, 2) & "</vBC>" 'Valor da base de cálculo do ipi
              End If
          
              strLinha = strLinha & "<pIPI>" & Replace(FormataValorTexto(.Fields("Percentual IPI").Value, 2), ",", ".") & "</pIPI>" 'Aliquota do IPI
              
              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" Then
                  strLinha = strLinha & "<vIPI>" & FormataValorTexto(Truncate((.Fields("Percentual IPI").Value * (.Fields("Qtde").Value * .Fields("Preço").Value - CDbl(sDescontoAuxZ_2))) / 100, 2), 2) & "</vIPI>" 'Valor do IPI
              Else
                  strLinha = strLinha & "<vIPI>" & FormataValorTexto(Truncate((.Fields("Percentual IPI").Value * (.Fields("Qtde").Value * .Fields("Preço").Value)) / 100, 2), 2) & "</vIPI>" 'Valor do IPI
              End If
          
              lIndice01 = InStr(1, strLinha, "<vIPI>")
              If lIndice01 > 0 Then
                  lIndice02 = InStr(lIndice01, strLinha, "</vIPI>")
                  sAuxIPI = Mid$(strLinha, lIndice01 + 6, lIndice02 - (lIndice01 + 6))
                  sAuxIPI = Replace(sAuxIPI, ".", ",")
                  vlSOMA_DO_IPI = vlSOMA_DO_IPI + CDbl(sAuxIPI)
              End If

              strLinha = strLinha & "</IPITrib>"
              strLinha = strLinha & "</IPI>"
              AddLinha clcArquivo, strLinha
          End If
      End If

      If blnSimplesNacional = False Then
            ' ******************
            ' Empresa LUCRO REAL
            
            If Right(strTributacao, 2) = "00" Or Right(strTributacao, 2) = "20" Or Right(strTributacao, 2) = "40" Or Right(strTributacao, 2) = "60" Then
                
                ' ******************************************************
                ' CST acima é "00" - SIM...Destacar ICMS no PIS E COFINS
                ' ******************************************************
                
                If localTipoSituacaoTributariaPIS = 99 Then
                    'Pis
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISOutr>"
                    strLinha = strLinha & "<CST>99</CST>"
                    
                    strLinha = strLinha & "<vBC>0.00</vBC>"
    
                    strLinha = strLinha & "<pPIS>0.0000</pPIS>"
                    strLinha = strLinha & "<vPIS>0.00</vPIS>"
                    strLinha = strLinha & "</PISOutr>"
                    strLinha = strLinha & "</PIS>"
    
                    'Cofins
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSOutr>"
                    strLinha = strLinha & "<CST>99</CST>"
                    strLinha = strLinha & "<vBC>0.00</vBC>"
                    
                    strLinha = strLinha & "<pCOFINS>0.0000</pCOFINS>"
                    strLinha = strLinha & "<vCOFINS>0.00</vCOFINS>"
                    strLinha = strLinha & "</COFINSOutr>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                    
                ElseIf localTipoSituacaoTributariaPIS = 4 Or localTipoSituacaoTributariaPIS = 6 Or _
                      localTipoSituacaoTributariaPIS = 7 Or localTipoSituacaoTributariaPIS = 8 Or _
                      localTipoSituacaoTributariaPIS = 9 Then
                      
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISNT>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    strLinha = strLinha & "</PISNT>"
                    strLinha = strLinha & "</PIS>"
          
                    'Cofins
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSNT>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    strLinha = strLinha & "</COFINSNT>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                
                ElseIf localTipoSituacaoTributariaPIS = 1 Or localTipoSituacaoTributariaPIS = 2 Or _
                    localTipoSituacaoTributariaPIS = 3 Or localTipoSituacaoTributariaPIS = 5 Then
                    
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISAliq>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto(.Fields("Preço Final").Value * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualPIS), 2)
                    End If
                    
                    strLinha = strLinha & "</PISAliq>"
                    strLinha = strLinha & "</PIS>"
    
                    
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSAliq>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualCOFINS), 2) & "</vCOFINS>"
                        
                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualCOFINS), 2)
                    End If

                    strLinha = strLinha & "</COFINSAliq>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                    
                Else
                
                    'Tratando o PIS/COFINS 50, 73...por enquanto
                                       
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISAliq>"
                    strLinha = strLinha & "<CST>01</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto(.Fields("Preço Final").Value * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualPIS), 2)
                    End If
                    
                    strLinha = strLinha & "</PISAliq>"
                    strLinha = strLinha & "</PIS>"
    
                    
                    
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSAliq>"
                    strLinha = strLinha & "<CST>01</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualCOFINS), 2) & "</vCOFINS>"
                        
                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualCOFINS), 2)
                    End If

                    strLinha = strLinha & "</COFINSAliq>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                    
                End If
         
            Else
            
                ' ****************************************************
                ' CST acima é "10"
                ' CST acima é "30"
                ' CST acima é "50"
                ' CST acima é "70"
                ' CST acima é "80"
                ' CST acima é "90"
                ' Cai nesta condição ainda tratada como genérica....
                ' ****************************************************
            
                If localTipoSituacaoTributariaPIS = 99 Then
                    'Pis
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISOutr>"
                    strLinha = strLinha & "<CST>99</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                        sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                        sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS, 2) & "</vBC>"
                    End If
    
                    strLinha = strLinha & "<pPIS>0.0000</pPIS>"
                    strLinha = strLinha & "<vPIS>0.00</vPIS>"
                    strLinha = strLinha & "</PISOutr>"
                    strLinha = strLinha & "</PIS>"
    
                    'Cofins
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSOutr>"
                    strLinha = strLinha & "<CST>99</CST>"
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                        sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                        sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS, 2) & "</vBC>"
                    End If
                    
                    strLinha = strLinha & "<pCOFINS>0.0000</pCOFINS>"
                    strLinha = strLinha & "<vCOFINS>0.00</vCOFINS>"
                    strLinha = strLinha & "</COFINSOutr>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                    
                ElseIf localTipoSituacaoTributariaPIS = 4 Or localTipoSituacaoTributariaPIS = 6 Or _
                      localTipoSituacaoTributariaPIS = 7 Or localTipoSituacaoTributariaPIS = 8 Or _
                      localTipoSituacaoTributariaPIS = 9 Then
                      
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISNT>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    strLinha = strLinha & "</PISNT>"
                    strLinha = strLinha & "</PIS>"
          
                    'Cofins
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSNT>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    strLinha = strLinha & "</COFINSNT>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                
                ElseIf localTipoSituacaoTributariaPIS = 1 Or localTipoSituacaoTributariaPIS = 2 Or _
                    localTipoSituacaoTributariaPIS = 3 Or localTipoSituacaoTributariaPIS = 5 Then
                    
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISAliq>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto(.Fields("Preço Final").Value * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualPIS), 2)
                    End If
                    
                    strLinha = strLinha & "</PISAliq>"
                    strLinha = strLinha & "</PIS>"
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSAliq>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualCOFINS), 2) & "</vCOFINS>"
                        
                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualCOFINS), 2)
                    End If

                    strLinha = strLinha & "</COFINSAliq>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha

                Else
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISAliq>"
                    strLinha = strLinha & "<CST>01</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto(.Fields("Preço Final").Value * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualPIS), 2)
                    End If
                    
                    strLinha = strLinha & "</PISAliq>"
                    strLinha = strLinha & "</PIS>"
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSAliq>"
                    strLinha = strLinha & "<CST>01</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And ValorBaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualCOFINS), 2) & "</vCOFINS>"
                        
                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualCOFINS), 2)
                    End If

                    strLinha = strLinha & "</COFINSAliq>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                  
                End If
            End If
      Else

            'PISAliq
            '01 – Operação Tributável - Base de Cálculo = Valor da Operação Alíquota Normal (Cumulativo/Não Cumulativo) = PISAliq
            '02 - Operação Tributável - Base de Calculo = Valor da Operação (Alíquota Diferenciada) = PISAliq
            
            'PISQtde
            '03 - Operação Tributável - Base de Calculo = Quantidade Vendida x Alíquota por Unidade de Produto = PISQtde
            
            'PISNT
            '04 - Operação Tributável - Tributação Monofásica - (Alíquota Zero) = PISNT
            '06 - Operação Tributável - Alíquota Zero = PISNT
            '07 - Operação Isenta da contribuição = PISNT
            '08 - Operação Sem Incidência da contribuição = PISNT
            '09 - Operação com suspensão da contribuição = PISNT
            
            'PISOutr
            '99 - Outras Operações = PISOutr
  
            If localTipoSituacaoTributariaPIS = 99 Then
                'Pis
                strLinha = "<PIS>"
                strLinha = strLinha & "<PISOutr>"
                strLinha = strLinha & "<CST>99</CST>"
                strLinha = strLinha & "<vBC>0.00</vBC>"
                strLinha = strLinha & "<pPIS>0.0000</pPIS>"
                strLinha = strLinha & "<vPIS>0.00</vPIS>"
                strLinha = strLinha & "</PISOutr>"
                strLinha = strLinha & "</PIS>"

                'Cofins
                strLinha = strLinha & "<COFINS>"
                strLinha = strLinha & "<COFINSOutr>"
                strLinha = strLinha & "<CST>99</CST>"
                strLinha = strLinha & "<vBC>0.00</vBC>"
                strLinha = strLinha & "<pCOFINS>0.0000</pCOFINS>"
                strLinha = strLinha & "<vCOFINS>0.00</vCOFINS>"
                strLinha = strLinha & "</COFINSOutr>"
                strLinha = strLinha & "</COFINS>"
                AddLinha clcArquivo, strLinha
            Else
                'Pis
                strLinha = "<PIS>"
                strLinha = strLinha & "<PISNT>"
                strLinha = strLinha & "<CST>" & "07" & "</CST>"
                strLinha = strLinha & "</PISNT>"
                strLinha = strLinha & "</PIS>"
      
                'Cofins
                strLinha = strLinha & "<COFINS>"
                strLinha = strLinha & "<COFINSNT>"
                strLinha = strLinha & "<CST>" & "07" & "</CST>"
                strLinha = strLinha & "</COFINSNT>"
                strLinha = strLinha & "</COFINS>"
                AddLinha clcArquivo, strLinha
            End If
      End If
        
      If (isConsumidorFinal And isInterestadual And (IndicadorIE = 2 Or IndicadorIE = 9)) Then
            Dim vBCUFDest As Double
            Dim pFCPUFDest As Double
            Dim pICMSUFDest As Double
            Dim pICMSInter As Double
            Dim pICMSInterPart As Double
            Dim vFCPUFDest As Double
            Dim vICMSUFDest As Double
            Dim vICMSUFRemet As Double
            Dim difal As Double
          
            If BaseCalculoICMS_ST > 0 Then
              vBCUFDest = BaseCalculoICMS_ST
            Else
              vBCUFDest = ValorBaseCalculoICMS
            End If
          
            pFCPUFDest = 2 'valor padrão colocado no QuickStore, mas cada UF pode ter uma regra diferente
            pICMSUFDest = rsSaidas.Fields("aliquota_destino").Value
            
            'Alíquota interestadual das UF envolvidas:
            '- 4% alíquota interestadual para produtos importados;
            '- 7% para os Estados de origem do Sul e Sudeste (exceto
            '   ES), destinado para os Estados do Norte, Nordeste, CentroOeste e Espírito Santo;
            '- 12% para os demais casos.
            If InStr("RS SC PR SP MG RJ", sEstadoFilial) <> 0 And InStr("AC AL AP AM BA CE ES GO MA MT MS PA PB PE PI RN RO RR SE TO", UF_destino) <> 0 Then
              pICMSInter = 7
            Else
              pICMSInter = 12
            End If
            
            pICMSInterPart = 100  ' 100% para 2019 em diante
            vFCPUFDest = 0
            vICMSUFDest = vBCUFDest * (pICMSUFDest / 100)
            vICMSUFRemet = ValorICMS
          
            'calcular o Fundo de Combate à Pobreza
            'FCP = Base do ICMS * (%FCP / 100)
            'FCP = 1.000,00 * (2,00% / 100)
            'FCP = 1.000,00 * 0,02
            'FCP = 20,00
            vFCPUFDest = ((vBCUFDest * pFCPUFDest) / 100)
          
            'calcular o DIFAL
            'DIFAL = Base do ICMS * ((%Alíquota do ICMS Intra - %Alíquota do ICMS Inter) / 100)
            'DIFAL = 1.000,00 * ((18,00% - 12,00%) / 100)
            'DIFAL = 1.000,00 * (6,00% / 100)
            'DIFAL = 1.000,00 * 0,06
            'DIFAL = 60,00
             difal = vBCUFDest * ((pICMSUFDest - pICMSInter) / 100)
           
            'Parte que compete ao estado de destino:
            'vICMSUFDest = vBCUFDest * (pICMSUFDest - pICMSInter) * pICMSInterPart
            'Parte UF Destino = Valor do DIFAL * (%Destino / 100) 80 - 2018 | 100 - 2019
            'Parte MA = 60,00 * (40,00% / 100)
            'Parte MA = 60, 0 * 0, 40
            'Parte MA = 24, 0
            vICMSUFDest = difal
           
            'efetuar a partilha do DIFAL
            'vICMSUFRemet = (vBCUFDest * (pICMSUFDest - pICMSInter)) - vICMSUFDest
            'Parte UF Origem = Valor do DIFAL * (%Origem / 100) 20 - 2018|0 - 2019
            'Parte SC = 60,00 * (60,00% / 100)
            'Parte SC = 60, 0 * 0, 60
            'Parte SC = 36, 0
            vICMSUFRemet = 0

            If blnSimplesNacionalGeral = False Then
                ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                If isInterestadual And (IndicadorIE = 9 Or IndicadorIE = 2) And (str_FCP_RG_INS_EST_destinatario = "") Then

                    ' Novo grupo ICMSUFDest
                    ' [3179b;vBCUFDest;vBCFCPUFDest;pFCPUFDest;pICMSUFDest;pICMSInter;pICMSInterPart;vFCPUFDest;vICMSUFDest;vICMSUFRemet]

                    strLinha = "<ICMSUFDest>"
                    
                    'vBCUFDest ------------------------
                    strLinha = strLinha & "<vBCUFDest>" & FormataValorTexto(vBCUFDest, 2) & "</vBCUFDest>"
                    'vBCFCPUFDest ------------------------
                    strLinha = strLinha & "<vBCFCPUFDest>" & FormataValorTexto(vBCUFDest, 2) & "</vBCFCPUFDest>"
                    'pFCPUFDest ------------------------
                    strLinha = strLinha & "<pFCPUFDest>" & FormataValorTexto(pFCPUFDest, 2) & "</pFCPUFDest>"
                    'pICMSUFDest ------------------------
                    strLinha = strLinha & "<pICMSUFDest>" & FormataValorTexto(pICMSUFDest, 2) & "</pICMSUFDest>"
                    'pICMSInter ------------------------
                    strLinha = strLinha & "<pICMSInter>" & FormataValorTexto(pICMSInter, 2) & "</pICMSInter>"
                    'pICMSInterPart ------------------------
                    strLinha = strLinha & "<pICMSInterPart>" & FormataValorTexto(pICMSInterPart, 2) & "</pICMSInterPart>"
                    'vFCPUFDest ------------------------
                    strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(vFCPUFDest, 2) & "</vFCPUFDest>"
                    'vICMSUFDest --------------------------
                    ' DIFAL
                    strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(vICMSUFDest, 2) & "</vICMSUFDest>"
                    ' vICMSUFRemet ------------------------
                    strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(vICMSUFRemet, 2) & "</vICMSUFRemet>"
                    strLinha = strLinha & "</ICMSUFDest>"
                    AddLinha clcArquivo, strLinha
                    
                    vFCPUFDest_registroTotal = vFCPUFDest_registroTotal + CDbl(Replace(FormataValorTexto(vFCPUFDest, 2), ".", ","))
                    vICMSUFDest_registroTotal = vICMSUFDest_registroTotal + CDbl(Replace(FormataValorTexto(vICMSUFDest, 2), ".", ","))
                End If
            End If
      End If
        
      strLinha = strLinha & "</imposto>"
        
      If bFINALIDADE_NFE_DEVOLUCAO = True And nfeDevolucao_impostoDevol = True Then
            'Colocar neste condição o novo grupo tag <impostoDevol>
            strLinha = strLinha & "<impostoDevol>"
            strLinha = strLinha & "<pDevol>" & Replace(Format(.Fields("IPI").Value), ",", ".") & "</pDevol>" 'Aliquota do IPI
            strLinha = strLinha & "<IPI>"
            strLinha = strLinha & "<vIPIDevol>" & FormataValorTexto((.Fields("IPI").Value * (.Fields("Qtde").Value * .Fields("Preço").Value)) / 100, 2) & "</vIPIDevol>" 'Valor do IPI
            strLinha = strLinha & "</IPI>"
            strLinha = strLinha & "</impostoDevol>"
      End If
        
      ' ===============================================================================
      ' TAG: infAdProd
      If nfeInfAdProd = True Then
          If Not IsNull(.Fields("Obs").Value) Then
              If Len(Trim(.Fields("Obs").Value)) > 0 Then
                  strLinha = strLinha & "<infAdProd>"
                  
                  If Len(Trim(.Fields("Obs").Value)) > 500 Then
                      strLinha = strLinha & Mid(trataCaracteresEspeciaisNFE(.Fields("Obs").Value), 1, 499)
                  Else
                      strLinha = strLinha & trataCaracteresEspeciaisNFE(.Fields("Obs").Value)
                  End If
                  strLinha = strLinha & "</infAdProd>"
              End If
          End If
      End If
      ' ===============================================================================
      
      ' Pilatti Outubro/2017 - linha incluida
      strLinha = strLinha & "</det>"
      ' Pilatti Fim

      AddLinha clcArquivo, strLinha
    
      rsSaidasProdutos_CFOP.Close
      
      strSQL = ""
      .MoveNext
    Loop
    .Close
  End With
  Set rsSaidasProdutos = Nothing
  
  If Not (rsSaidasAUXILIAR_ICMS Is Nothing) Then
      rsSaidasAUXILIAR_ICMS.Close
      Set rsSaidasAUXILIAR_ICMS = Nothing
  End If
  
End Function

'18/01/2010 - Andrea
Private Function AddProdutoEntrada(ByVal Filial As String, ByVal Sequencia As Long, ByRef clcArquivo As Collection, ByRef rsEntradas As Recordset, ByRef rsParametros As Recordset, PercentualPIS As Double, PercentualCOFINS As Double, ByVal blnSimplesNacional As Boolean, ByVal dblAliquotaSimples As Double, ByVal dblPercentualReducaoBC_SN As Double, ByVal CSO_NaturezaOperacao As String, ByVal strEstadoCliente As String, ByVal PadraoArquivoIntegracao As String)
  Dim strSQL As String
  Dim strLinha As String
  Dim strTributacao As String
  Dim strCFOP As String
  Dim rsEntradasAUXILIAR_ICMS As Recordset
  Dim sResp As String
  
  Dim nomeProdutoNota As String
  Dim nomeProduto As String
  Dim iContaRS As Integer
  Dim iContaLoop As Integer
  Dim s_vlSOMA_DO_ICMS As String
  vlSOMA_DO_ICMS = 0
  s_vlSOMA_DO_ICMS = ""
  iContaRS = 0
  iContaLoop = 0
  Dim strCest As String
  Dim strSQLCest As String
  
  'Valores Totais da NF-e
  strSQL = "Select * "
  strSQL = strSQL & " FROM [Entradas] WHERE Filial = " & Filial
  strSQL = strSQL & "AND Sequência = " & Sequencia
  Set rsEntradasAUXILIAR_ICMS = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  Dim dblValor As Double
  Dim lngCasasDecimais As Long
  Dim rsEntradasProdutos As Recordset
  Dim rsProdutoCFOP As Recordset
  Dim rsOp_Entradas As Recordset
  
  Dim rsEntradasProdutos_CFOP As Recordset
  
  'Rateio dos valores de frete e seguro pelos itens da movimentacao
  Dim dbl_percentual_frete_rateado As Double
  Dim dbl_percentual_seguro_rateado As Double
  Dim dbl_valor_frete_rateado As Double
  Dim dbl_valor_seguro_rateado As Double
   
  'Add Produtos
  strSQL = "SELECT Produtos.Código, Produtos.Nome, Produtos.CodigoNBM,Produtos.[Nome Nota], " ' , [Saídas - Produtos].CFOP
  strSQL = strSQL & "Produtos.[Tipo ICM], Produtos.[Base Cálculo], Produtos.[Redução ICM], "
  strSQL = strSQL & "Produtos.[Unidade Venda], [Entradas - Produtos].Qtde, [Entradas - Produtos].Preço, "
  strSQL = strSQL & "[Entradas - Produtos].[Preço Final],[Entradas - Produtos].Desconto as DescontoProdutos, "
  strSQL = strSQL & "Entradas.Produtos, Entradas.Frete, Entradas.Desconto, Entradas.Operação, " ' , Entradas.Seguro
  strSQL = strSQL & "[Entradas - Produtos].ICM, Produtos.[Situação Tributária], [Entradas - Produtos].IPI, Produtos.[IPI_Reduzido], [Entradas - Produtos].Linha, "
  strSQL = strSQL & "[Entradas - Produtos].Valor_Aprox_Impostos, [Entradas - Produtos].MotivoDesoneracaoICMS,"
  strSQL = strSQL & "[Entradas - Produtos].ValorICMSDesonerado, [Entradas - Produtos].Percentual_Diferimento"
  strSQL = strSQL & " FROM [Entradas - Produtos], Produtos, Entradas "
  strSQL = strSQL & " WHERE [Entradas - Produtos].Filial = " & Filial
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = " & Sequencia
  strSQL = strSQL & " AND Produtos.Código = [Entradas - Produtos].[Código Sem Grade] "
  strSQL = strSQL & " AND Entradas.Filial = " & Filial
  strSQL = strSQL & " AND Entradas.Sequência = " & Sequencia
  
  Set rsEntradasProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  Set rsProdutoCFOP = db.OpenRecordset("ProdutoCFOP", dbOpenTable, dbReadOnly)
  Set rsOp_Entradas = db.OpenRecordset("Operações Entrada", dbOpenTable, dbReadOnly)
  
  Dim BaseCalculoICMS As Double
  Dim AliquotaICMS As Double
  Dim ValorICMS As Double
  Dim BaseCalculoICMS_ST As Double
  Dim AliquotaICMS_ST As Double
  Dim ValorICMS_ST As Double
  Dim MargemValorAdicionado_ICMS As Double
  Dim PercentualReducaoBaseICMS As Double
  
  '21/01/2010 - mpdea
  Dim dbl_total_produtos As Double
    
  '31/03/2011 - Andrea
  Dim str_CSO_Produto As String
  Dim strValorItem As String

  With rsEntradasProdutos
    '21/01/2010 - mpdea
    'Total de produtos para rateio do frete
    dbl_total_produtos = 0
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        dbl_total_produtos = dbl_total_produtos + CDbl(rsEntradasProdutos.Fields("Preço Final").Value)
        .MoveNext
      Loop
      .MoveFirst
    End If
    
    Do Until .EOF
      strValorItem = Round(.Fields("Preço").Value, 2)

      '21/01/2010 - mpdea
      'Corrigido percentual de rateio do item
      
      'Entrada Produtos CFOP
      strSQL = "SELECT CFOP, CSO  "
      strSQL = strSQL & " FROM ProdutoCFOP "
      strSQL = strSQL & " WHERE CodProduto = '" & rsEntradasProdutos.Fields("Código").Value & "'"
      strSQL = strSQL & " And CodOperacao = " & .Fields("Operação").Value

      Set rsEntradasProdutos_CFOP = db.OpenRecordset(strSQL, dbOpenDynaset)
      '
      'Rateio do Frete por item
      If rsEntradas.Fields("Frete").Value > 0 And dbl_total_produtos > 0 Then
        dbl_percentual_frete_rateado = ((rsEntradasProdutos.Fields("Preço Final").Value * 100) / dbl_total_produtos)
        If dbl_percentual_frete_rateado > 0 Then
          dbl_valor_frete_rateado = ((rsEntradas.Fields("Frete").Value * dbl_percentual_frete_rateado) / 100)
        End If
      End If
    'Loop
      'Procura o CFOP do produto.
      rsProdutoCFOP.Index = "PrimaryKey"
      rsProdutoCFOP.Seek "=", rsEntradasProdutos.Fields("Código").Value, rsEntradasProdutos.Fields("Operação").Value
      If rsProdutoCFOP.NoMatch Then
        rsOp_Entradas.Index = "Código"
        rsOp_Entradas.Seek "=", rsEntradasProdutos.Fields("Operação").Value
        If Not rsOp_Entradas.NoMatch Then
          strCFOP = rsOp_Entradas("Código Fiscal") & ""
          str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
        End If
      Else
        strCFOP = rsProdutoCFOP("CFOP") & ""
      End If
      
      If Len(.Fields("Nome Nota").Value) > 0 Then
        nomeProduto = .Fields("Nome Nota").Value
      Else
        nomeProduto = .Fields("Nome").Value
      End If

      '21/11/2011 - Andrea
      Dim str_NCM As String
      Dim str_CFOP As String
      Dim Linha As String

      If PadraoArquivoIntegracao = "XML" Then

          ' ------------------------------
          ' XML NFe 4.00
        
          iDetNumItemXML = iDetNumItemXML + 1
          strLinha = "<det nItem=""" & iDetNumItemXML & """>"
          strLinha = strLinha & "<prod>"
          strLinha = strLinha & "<cProd>" & .Fields("Código").Value & "</cProd>"
          strLinha = strLinha & "<cEAN>SEM GTIN</cEAN>"
          nomeProdutoNota = RemoveCaracteresEspeciaisParaNFE(.Fields("Nome").Value)
          strLinha = strLinha & "<xProd>" & nomeProdutoNota & "</xProd>"
          strLinha = strLinha & "<NCM>" & .Fields("CodigoNBM").Value & "</NCM>"
          
          ' PILATTI/2018 nfe4.0 ---- vai para o REGISTRO 3000b
          'inclui código CEST
          strCest = ""
          Dim rsCest3 As Recordset
          strSQLCest = "select cest from Ref_CEST_NCM where ncm = '" & .Fields("CodigoNBM").Value & "';"
          Set rsCest3 = db.OpenRecordset(strSQLCest, dbOpenDynaset)
          If Not (rsCest3.BOF And rsCest3.EOF) Then
            Do Until rsCest3.EOF
              strCest = CStr(rsCest3.Fields("cest").Value)
              rsCest3.MoveNext
            Loop
            rsCest3.MoveFirst
            
            strLinha = strLinha & "<CEST>" & strCest & "</CEST>"
          End If
          rsCest3.Close
          Set rsCest3 = Nothing
          ' FIM PILATTI
          
          If Not rsEntradasProdutos_CFOP.EOF Then
              If Not IsNull(rsEntradasProdutos_CFOP.Fields("CFOP").Value) And rsEntradasProdutos_CFOP.Fields("CFOP").Value <> 0 Then
                  str_CFOP = Format(rsEntradasProdutos_CFOP.Fields("CFOP").Value, String(4, "0")) 'CFOP do cadastro de produto X CFOP
              Else
                  str_CFOP = Format(.Fields("CFOP").Value, String(4, "0")) 'CFOP dos itens da venda
              End If
              
              If Not IsNull(rsEntradasProdutos_CFOP.Fields("CSO").Value) And rsEntradasProdutos_CFOP.Fields("CSO").Value <> 0 Then
                  str_CSO_Produto = Format(rsEntradasProdutos_CFOP.Fields("CSO").Value, String(3, "0")) 'CSO do produto X Natureza de operacao
              Else
                  str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
              End If
          Else
              str_CFOP = Format(strCFOP, String(4, "0")) 'CFOP do Produto
              str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
          End If


          ' ***************************************************************
          ' Tratamento de inconsistência no cadastro de CFOP
          If IsNull(str_CFOP) Or str_CFOP = "" Then
              While str_CFOP = ""
                  sResp = InputBox("O cadastro do produto " & vbCrLf & "CÓDIGO: " & rsEntradasProdutos.Fields("Código").Value & vbCrLf & "NOME: " & nomeProduto & vbCrLf & " está INCOMPLETO ! " & vbCrLf & vbCrLf & "Pontualmente lhe ajudaremos a seguir com a emissão desta NF-e. Porém, você deverá COMPLETAR este cadastro posteriormente." & vbCrLf & vbCrLf & "Agora, Digite o CFOP correto vinculado a este produto " & vbCrLf & "Ex: 1202, 1411, entre outros", "Cadastro INCOMPLETO - Vincular CFOP")
                  
                  If IsNull(sResp) Then
                      str_CFOP = ""
                  ElseIf Not IsNumeric(sResp) Then
                      str_CFOP = ""
                  ElseIf Len(Trim(sResp)) <> 4 Then
                      str_CFOP = ""
                  Else
                      str_CFOP = sResp
                  End If
              Wend
          End If
          ' ***************************************************************

          
          strLinha = strLinha & "<CFOP>" & str_CFOP & "</CFOP>"
          strLinha = strLinha & "<uCom>" & .Fields("Unidade Venda").Value & "</uCom>"
          strLinha = strLinha & "<qCom>" & FormataValorTexto(.Fields("Qtde").Value, 4) & "</qCom>"
          strLinha = strLinha & "<vUnCom>" & FormataValorTexto(strValorItem, 4) & "</vUnCom>"
          strLinha = strLinha & "<vProd>" & FormataValorTexto((.Fields("Qtde").Value * .Fields("Preço").Value), 2) & "</vProd>"
          strLinha = strLinha & "<cEANTrib>SEM GTIN</cEANTrib>"
          strLinha = strLinha & "<uTrib>" & .Fields("Unidade Venda").Value & "</uTrib>"
          strLinha = strLinha & "<qTrib>" & FormataValorTexto(.Fields("Qtde").Value, 4) & "</qTrib>"
          strLinha = strLinha & "<vUnTrib>" & FormataValorTexto(strValorItem, 4) & "</vUnTrib>"
          
          If dbl_valor_frete_rateado <> 0 Then
            strLinha = strLinha & "<vFrete>" & FormataValorTexto(dbl_valor_frete_rateado, 2) & "</vFrete>"
          End If
          If dbl_valor_seguro_rateado <> 0 Then
            strLinha = strLinha & "<vSeg>" & FormataValorTexto(dbl_valor_seguro_rateado, 2) & "</vSeg>"
          End If
          
          If ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) <= 0.009 Then
          Else
            m_bln_tem_desconto = True
            strLinha = strLinha & "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value), 2) & "</vDesc>"
          End If
  
          strLinha = strLinha & "<indTot>" & "1" & "</indTot>"
          strLinha = strLinha & "</prod>"
          AddLinha clcArquivo, strLinha
      End If

      'Calculo dos encargos por item
      If rsEntradasProdutos.Fields("Tipo ICM").Value = "N" Then
          'ICMS Normal
          BaseCalculoICMS = rsEntradasProdutos.Fields("Preço Final").Value
          AliquotaICMS = rsEntradasProdutos.Fields("ICM").Value
          ValorICMS = rsEntradasProdutos.Fields("Preço Final").Value * CSng(gsHandleNull(rsEntradasProdutos.Fields("ICM").Value & "")) / 100
      ElseIf rsEntradasProdutos.Fields("Tipo ICM").Value = "R" Then
          'ICMS Retido
          If CDbl(rsEntradasProdutos.Fields("Base Cálculo").Value) <> 0 Then
            'Base Fixa
            BaseCalculoICMS_ST = CDbl(rsEntradasProdutos.Fields("Qtde").Value) * CDbl(rsEntradasProdutos.Fields("Base Cálculo").Value)
            AliquotaICMS_ST = rsEntradasProdutos.Fields("ICM").Value
            ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsEntradasProdutos.Fields("ICM").Value & "")) / 100
          ElseIf CDbl(rsEntradasProdutos.Fields("Redução ICM").Value) <> 0 Then
            'Base Reduzida
            BaseCalculoICMS_ST = CDbl(rsEntradasProdutos.Fields("Preço Final").Value) * CDbl(rsEntradasProdutos.Fields("Redução ICM").Value) / 100
            AliquotaICMS_ST = rsEntradasProdutos.Fields("ICM").Value
            ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsEntradasProdutos.Fields("ICM").Value & "")) / 100
            PercentualReducaoBaseICMS = rsEntradasProdutos.Fields("Redução ICM").Value
          End If
      ElseIf rsEntradasProdutos.Fields("Tipo ICM").Value = "Z" Then
          'Tratamento do
          'ICMS Reduzido
          If CDbl(rsEntradasProdutos.Fields("Base Cálculo").Value) <> 0 Then
            'Base Fixa
            BaseCalculoICMS_ST = CDbl(rsEntradasProdutos.Fields("Qtde").Value) * CDbl(rsEntradasProdutos.Fields("Base Cálculo").Value)
            AliquotaICMS_ST = rsEntradasProdutos.Fields("ICM").Value
            ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsEntradasProdutos.Fields("ICM").Value & "")) / 100
          ElseIf CDbl(rsEntradasProdutos.Fields("Redução ICM").Value) <> 0 Then
            'Base Reduzida
            BaseCalculoICMS_ST = CDbl(rsEntradasProdutos.Fields("Preço Final").Value) * CDbl(rsEntradasProdutos.Fields("Redução ICM").Value) / 100
            AliquotaICMS_ST = rsEntradasProdutos.Fields("ICM").Value
            ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsEntradasProdutos.Fields("ICM").Value & "")) / 100
            PercentualReducaoBaseICMS = rsEntradasProdutos.Fields("Redução ICM").Value
          End If
      End If
      
      strTributacao = Right(.Fields("Situação Tributária").Value, 2)
      
      Dim dblBaseCalculoSimplesItem As Double
      If PadraoArquivoIntegracao = "XML" Then 'XML

        ' ------------------------------------
        'XML NFe 4.0

        strLinha = "<imposto>"
        
        If Not IsNull(.Fields("Valor_Aprox_Impostos").Value) And .Fields("Valor_Aprox_Impostos").Value <> 0 Then
          strLinha = strLinha & "<vTotTrib>" & FormataValorTexto(.Fields("Valor_Aprox_Impostos").Value, 2) & "</vTotTrib>"
        End If
        AddLinha clcArquivo, strLinha
        
        strLinha = "<ICMS>"
        
        If blnSimplesNacional = False Then

          'Select Case strTributacao
          Select Case Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2)
            Case "00"
              'Tributos Incidentes no Produto ou no serviço
              strLinha = strLinha & "<ICMS00>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
              strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
              strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
              strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"

              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              ' IDEM [3101;pFCP;vFCP]
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  vBC_fcp = BaseCalculoICMS
        
                  vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                  strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                  strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                  vFCP_registroTotal = vFCP_registroTotal + vResultado
              End If
              
              strLinha = strLinha & "</ICMS00>"
              AddLinha clcArquivo, strLinha
            
            Case "10"
              'ICMS Normal e ST
              strLinha = strLinha & "<ICMS10>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
              strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
              strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
              strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  Dim vAuxiliar As Double
                  vAuxiliar = BaseCalculoICMS_ST
        
                  ' COLOCAR GRUPO 3111 ou 3112
                  If vAuxiliar = 0 Then
                      vBC_fcp = BaseCalculoICMS
        
                      vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                      ' IDEM [3111;vBCFCP;pFCP;vFCP]
                      strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                      strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                      strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                      vFCP_registroTotal = vFCP_registroTotal + vResultado
                  End If
              End If
              
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
              strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
              strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
              strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 2) & "</pICMSST>"
              strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"

              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  vAuxiliar = BaseCalculoICMS_ST
        
                  ' COLOCAR GRUPO 3111 ou 3112
                  If vAuxiliar <> 0 Then
                      vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                      ' [3112;vBCFCPST;pFCPST;vFCPST]
                      strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                      strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                      strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                      vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                  End If
              End If
              
              strLinha = strLinha & "</ICMS10>"
              AddLinha clcArquivo, strLinha

            Case "20"
              'Dados do ICMS da operação
              strLinha = strLinha & "<ICMS20>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
              strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBC>"
              strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
              strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 4) & "</pICMS>"
              strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  vBC_fcp = BaseCalculoICMS
                  vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
        
                  ' [3121;vBCFCP;pFCP;vFCP]
                  strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                  strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                  strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                  vFCP_registroTotal = vFCP_registroTotal + vResultado
              End If
              
              If rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value = 3 Or rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value = 9 Or rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value = 12 Then
                strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(rsEntradasProdutos.Fields("ValorICMSDesonerado").Value, 2) & "</vICMSDeson>" 'Valor ICms Desonerado
                strLinha = strLinha & "<motDesICMS>" & rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value & "</motDesICMS>" 'Código do motivo Icms Desonerado
              Else
                strLinha = strLinha & "<vICMSDeson></vICMSDeson><motDesICMS></motDesICMS>"
              End If
              strLinha = strLinha & "</ICMS20>"
              AddLinha clcArquivo, strLinha

            Case "30"
              'Dados do icms da operação
              strLinha = strLinha & "<ICMS30>" & vbCrLf
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>" & vbCrLf
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>" & vbCrLf
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBCST>" & vbCrLf
              strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>" & vbCrLf
              strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBCST>" & vbCrLf
              strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMSST>" & vbCrLf
              strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS, 2) & "</vICMSST>" & vbCrLf
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  vAuxiliar = BaseCalculoICMS_ST
        
                  vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                  ' [3131;vBCFCPST;pFCPST;vFCPST]
                  strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                  strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                  strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                  vFCPST_registroTotal = vFCPST_registroTotal + vResultado
              End If
              
              If rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value = 3 Or rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value = 9 Or rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value = 12 Then
                strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(rsEntradasProdutos.Fields("ValorICMSDesonerado").Value, 2) & "</vICMSDeson>" 'Valor ICms Desonerado
                strLinha = strLinha & "<motDesICMS>" & rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value & "</motDesICMS>" 'Código do motivo Icms Desonerado
              Else
                strLinha = strLinha & "<vICMSDeson></vICMSDeson><motDesICMS></motDesICMS>"
              End If
              strLinha = strLinha & "</ICMS30>"
              AddLinha clcArquivo, strLinha
            
            Case "40", "41", "50"
              strLinha = strLinha & "<ICMS40>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
              If Len(rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value) >= 1 Then
                strLinha = strLinha & "<motDesICMS>" & rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value & "</motDesICMS>"
              Else
                strLinha = strLinha & "<motDesICMS>" & "1" & "</motDesICMS>" '  1 - Taxi
              End If
              strLinha = strLinha & "</ICMS40>"
              AddLinha clcArquivo, strLinha

            Case "51"
              'Dados do icms da operação
              strLinha = strLinha & "<ICMS51>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
              strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBC>"
              strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
              strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
              If rsEntradasProdutos.Fields("Percentual_Diferimento").Value > 0 Then
                strLinha = strLinha & "<vICMSOp>" & FormataValorTexto(ValorICMS, 2) & "</vICMSOp>"
                strLinha = strLinha & "<pDif>" & FormataValorTexto(rsEntradasProdutos.Fields("Percentual_Diferimento").Value, 2) & "</pDif>"
                strLinha = strLinha & "<vICMS>" & FormataValorTexto(((ValorICMS / 100) * rsEntradasProdutos.Fields("Percentual_Diferimento").Value), 2) & "</vICMS>"
              Else
                strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
              End If
              
              If BaseCalculoICMS <= 0 Then
                ' Entra aqui se não é ST
                ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
        
                    vBC_fcp = BaseCalculoICMS
                    vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                    ' [3146;vBCFCP;pFCP;vFCP]
                    strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                    strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                    strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                    vFCP_registroTotal = vFCP_registroTotal + vResultado
                End If
              End If
              
              strLinha = strLinha & "</ICMS51>"
              AddLinha clcArquivo, strLinha
            
            Case "60"
              'Dados do icms da operação
              strLinha = strLinha & "<ICMS60>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "<vBCSTRet>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCSTRet>"
             
              Dim vTesteAux As Double
              vTesteAux = BaseCalculoICMS_ST
              If vTesteAux > 0 Then
                  'pST Exemplo: alíquota da mercadoria na venda ao consumidor final = 18% e 2% de FCP
                  vResultado = AliquotaAdesaoUFDestFCP + AliquotaFCP_Destinatario
                  strLinha = strLinha & "<pST>" & FormataValorTexto(vResultado, 2) & "</pST>"
              Else
                  strLinha = strLinha & "<pST>0.00</pST>"
              End If
              
              strLinha = strLinha & "<vICMSSTRet>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSSTRet>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  vAuxiliar = BaseCalculoICMS_ST
        
                  vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                  ' [3151;vBCFCPSTRet;pFCPSTRet;vFCPSTRet]
                  strLinha = strLinha & "<vBCFCPSTRet>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPSTRet>"
                  strLinha = strLinha & "<pFCPSTRet>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPSTRet>"
                  strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vResultado, 2) & "</vFCPSTRet>"
                  vFCPSTRet_registroTotal = vFCPSTRet_registroTotal + vResultado
              End If
              
              strLinha = strLinha & "</ICMS60>"
              AddLinha clcArquivo, strLinha
            
            Case "70"
              strLinha = strLinha & "<ICMS70>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
              strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBC>" 'Percentual da Redução de BC do ICMS
              strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
              strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
              strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  vAuxiliar = BaseCalculoICMS_ST
        
                  ' COLOCAR GRUPO 3161           ou           3162
                  If vAuxiliar = 0 Then
                      vBC_fcp = BaseCalculoICMS
                      vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                      ' [3161;vBCFCP;pFCP;vFCP]
                      strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                      strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
                      strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                      vFCP_registroTotal = vFCP_registroTotal + vResultado
                  End If
              End If
              
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
              strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
              strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
              strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 2) & "</pICMSST>"
              strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                  'Dim vAuxiliar As Double
                  vAuxiliar = BaseCalculoICMS_ST
        
                  ' COLOCAR GRUPO 3161           ou           3162
                  If vAuxiliar <> 0 Then
                      vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                      ' [3162;vBCFCPST;pFCPST;vFCPST]
                      strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                      strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                      strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                      vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                  End If
              End If
              
              If Len(rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value) >= 1 Then
                strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(rsEntradasProdutos.Fields("ValorICMSDesonerado").Value, 2) & "</vICMSDeson>" 'Valor ICms Desonerado
                strLinha = strLinha & "<motDesICMS>" & rsEntradasProdutos.Fields("MotivoDesoneracaoICMS").Value & "</motDesICMS>"
              Else
                strLinha = strLinha & "<motDesICMS></motDesICMS>" '  1 - Taxi
              End If
              
              strLinha = strLinha & "</ICMS70>"
              AddLinha clcArquivo, strLinha
           
            Case "90"
              'Dados do icms da operação
              strLinha = strLinha & "<ICMS90>"
              strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(rsEntradasProdutos.Fields("Situação Tributária").Value, 2) & "</CST>"
              strLinha = strLinha & "</ICMS90>"
              AddLinha clcArquivo, strLinha

          End Select
        Else
          
          If dblPercentualReducaoBC_SN <> 0 Then
             dblBaseCalculoSimplesItem = ((.Fields("Qtde").Value * .Fields("Preço").Value) * dblPercentualReducaoBC_SN)
          Else
             dblBaseCalculoSimplesItem = (.Fields("Qtde").Value * .Fields("Preço").Value)
          End If
          
          If bFINALIDADE_NFE_DEVOLUCAO = True Then
              If str_CSO_Produto = "" And CSO_NaturezaOperacao <> "" Then
                  str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
              End If
          End If
          
          
          ' ***************************************************************
          ' Tratamento de inconsistência no cadastro de CSO
          If IsNull(str_CSO_Produto) Or str_CSO_Produto = "" Then
              While str_CSO_Produto = ""
                  sResp = InputBox("O cadastro do produto " & vbCrLf & "CÓDIGO: " & rsEntradasProdutos.Fields("Código").Value & vbCrLf & "NOME: " & nomeProduto & vbCrLf & " está INCOMPLETO ! " & vbCrLf & vbCrLf & "Pontualmente lhe ajudaremos a seguir com a emissão desta NF-e. Porém, você deverá COMPLETAR este cadastro posteriormente." & vbCrLf & vbCrLf & "Agora, Digite o CSO correto vinculado a este produto " & vbCrLf & "Ex: 101, 102, 103, 400, 900, entre outros", "Cadastro INCOMPLETO - Vincular CSO")
      
                  If IsNull(sResp) Then
                      str_CSO_Produto = ""
                  ElseIf Not IsNumeric(sResp) Then
                      str_CSO_Produto = ""
                  ElseIf Len(Trim(sResp)) <> 3 Then
                      str_CSO_Produto = ""
                  Else
                      str_CSO_Produto = sResp
                  End If
              Wend
          End If
          ' ***************************************************************
          
          Select Case str_CSO_Produto
  
            Case "101" 'Tributada pelo Simples Nacional com permissão de crédito
              strLinha = strLinha & "<ICMSSN101>"
              
              If IsNull(rsEntradasProdutos.Fields("Situação Tributária").Value) Or LTrim(RTrim(rsEntradasProdutos.Fields("Situação Tributária").Value)) = "" Then
                  strLinha = strLinha & "<orig>0</orig>"
              Else
                  strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              End If
              
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 4) & "</pCredSN>"
              strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
              strLinha = strLinha & "</ICMSSN101>"
              AddLinha clcArquivo, strLinha

            Case "102", "103", "300", "400"
              strLinha = strLinha & "<ICMSSN102>"
              
              If IsNull(rsEntradasProdutos.Fields("Situação Tributária").Value) Or LTrim(RTrim(rsEntradasProdutos.Fields("Situação Tributária").Value)) = "" Then
                  strLinha = strLinha & "<orig>0</orig>"
              Else
                  strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              End If
              
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "</ICMSSN102>"
              AddLinha clcArquivo, strLinha

            Case "201" 'Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por Substituição
              strLinha = strLinha & "<ICMSSN201>"
              
              If IsNull(rsEntradasProdutos.Fields("Situação Tributária").Value) Or LTrim(RTrim(rsEntradasProdutos.Fields("Situação Tributária").Value)) = "" Then
                  strLinha = strLinha & "<orig>0</orig>"
              Else
                  strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              End If
              
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBCST>"
              strLinha = strLinha & "<pMVAST>" & FormataValorTexto(MargemValorAdicionado_ICMS, 4) & "</pMVAST>"
              strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
              strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
              strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 4) & "</pICMSST>"
              strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                  vAuxiliar = BaseCalculoICMS_ST

                  vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                  ' [3177a;vBCFCPST;pFCPST;vFCPST]
                  strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                  strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                  strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                  vFCPST_registroTotal = vFCPST_registroTotal + vResultado
              End If
              
              strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 4) & "</pCredSN>"
              strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
              strLinha = strLinha & "</ICMSSN201>"
              AddLinha clcArquivo, strLinha

            Case "202", "203"
              '202- Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por Substituição Tributária
              '203- Isenção do ICMS nos Simples Nacional para faixa de receita bruta e com cobrança do ICMS por Substituição Tributária
              strLinha = strLinha & "<ICMSSN202>"
              
              If IsNull(rsEntradasProdutos.Fields("Situação Tributária").Value) Or LTrim(RTrim(rsEntradasProdutos.Fields("Situação Tributária").Value)) = "" Then
                  strLinha = strLinha & "<orig>0</orig>"
              Else
                  strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              End If
              
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBCST>"
              strLinha = strLinha & "<pMVAST>" & FormataValorTexto(MargemValorAdicionado_ICMS, 4) & "</pMVAST>"
              strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
              strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
              strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 4) & "</pICMSST>"
              strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then
                  vAuxiliar = BaseCalculoICMS_ST

                  vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                  ' [3178a;vBCFCPST;pFCPST;vFCPST]
                  strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                  strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                  strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                  vFCPST_registroTotal = vFCPST_registroTotal + vResultado
              End If
              
              strLinha = strLinha & "</ICMSSN202>"
              AddLinha clcArquivo, strLinha

            Case "500" 'ICMS cobrado anteriormente por substituição tributária (substituído) ou por antecipação
              strLinha = strLinha & "<ICMSSN500>"
              
              If IsNull(rsEntradasProdutos.Fields("Situação Tributária").Value) Or LTrim(RTrim(rsEntradasProdutos.Fields("Situação Tributária").Value)) = "" Then
                  strLinha = strLinha & "<orig>0</orig>"
              Else
                  strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              End If
              
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<vBCSTRet>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCSTRet>"
              
              'Dim vTesteAux As Double
              vTesteAux = BaseCalculoICMS_ST
              If vTesteAux > 0 Then
                  'pST Exemplo: alíquota da mercadoria na venda ao consumidor final = 18% e 2% de FCP
                  vResultado = AliquotaAdesaoUFDestFCP + AliquotaFCP_Destinatario
                  strLinha = strLinha & "<pST>" & FormataValorTexto(vResultado, 4) & "</pST>"
              Else
                  strLinha = strLinha & "<pST>0.00</pST>"
              End If
              
              strLinha = strLinha & "<vICMSSTRet>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSSTRet>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then

                  vAuxiliar = BaseCalculoICMS_ST

                  vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                  ' [3179a0;vBCFCPSTRet;pFCPSTRet;vFCPSTRet]
                  strLinha = strLinha & "<vBCFCPSTRet>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPSTRet>"
                  strLinha = strLinha & "<pFCPSTRet>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPSTRet>"
                  strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vResultado, 2) & "</vFCPSTRet>"
                  vFCPSTRet_registroTotal = vFCPSTRet_registroTotal + vResultado
              End If
              strLinha = strLinha & "</ICMSSN500>"
              AddLinha clcArquivo, strLinha

            Case "900" 'Isento
              strLinha = strLinha & "<ICMSSN900>"
              
              If IsNull(rsEntradasProdutos.Fields("Situação Tributária").Value) Or LTrim(RTrim(rsEntradasProdutos.Fields("Situação Tributária").Value)) = "" Then
                  strLinha = strLinha & "<orig>0</orig>"
              Else
                  strLinha = strLinha & "<orig>" & Left(rsEntradasProdutos.Fields("Situação Tributária").Value, 1) & "</orig>"
              End If
              
              strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
              strLinha = strLinha & "<modBC>" & rsParametros.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
              strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
              If PercentualReducaoBaseICMS > 0 Then
                strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBC>"
              End If
              strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 4) & "</pICMS>"
              
              If AliquotaICMS > 0 Then
                bTEM_ICMS_3179a = True
              End If
              
              iContaLoop = iContaLoop + 1
              If m_bln_tem_desconto = True Then
                s_vlSOMA_DO_ICMS = FormataValorTexto((BaseCalculoICMS - ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value)) * AliquotaICMS / 100, 2)
                s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                
                If iContaRS = iContaLoop Then
                    Dim s_total_icms_campo As String
                    Dim v_total_icms_campo As Double
                    Dim v_diferenca As Double
                    s_total_icms_campo = FormataValorTexto(rsEntradasAUXILIAR_ICMS.Fields("Valor ICM").Value)
                    s_total_icms_campo = Replace(s_total_icms_campo, ".", ",")
                    v_total_icms_campo = CDbl(s_total_icms_campo)
                    
                    If vlSOMA_DO_ICMS = v_total_icms_campo Then
                        strLinha = strLinha & "<vICMS>" & FormataValorTexto((BaseCalculoICMS - ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value)) * AliquotaICMS / 100, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
                    ElseIf vlSOMA_DO_ICMS < v_total_icms_campo Then
                        v_diferenca = v_total_icms_campo - vlSOMA_DO_ICMS
                        strLinha = strLinha & "<vICMS>" & FormataValorTexto(((BaseCalculoICMS - ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value)) * AliquotaICMS / 100) + v_diferenca, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
                    ElseIf vlSOMA_DO_ICMS > v_total_icms_campo Then
                        v_diferenca = vlSOMA_DO_ICMS - v_total_icms_campo
                        strLinha = strLinha & "<vICMS>" & FormataValorTexto(((BaseCalculoICMS - ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value)) * AliquotaICMS / 100) - v_diferenca, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
                    End If
                    
                Else
                    strLinha = strLinha & "<vICMS>" & FormataValorTexto((BaseCalculoICMS - ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value)) * AliquotaICMS / 100, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
                End If
              Else
                  s_vlSOMA_DO_ICMS = FormataValorTexto(ValorICMS, 2)
                  s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                  vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                  strLinha = strLinha & "<vICMS>" & FormataValorTexto(s_vlSOMA_DO_ICMS, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
              End If

              strLinha = strLinha & "<modBCST>" & rsParametros.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
              If MargemValorAdicionado_ICMS > 0 Then
                strLinha = strLinha & "<pMVAST>" & FormataValorTexto(MargemValorAdicionado_ICMS, 4) & "</pMVAST>"
              End If
              If PercentualReducaoBaseICMS > 0 Then
                strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 4) & "</pRedBCST>"
              End If
              
              strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
              strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 4) & "</pICMSST>"
              strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
              
              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
              If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") And blnSimplesNacionalGeral = False Then ' Or isIEIsento = True) Then

                  vAuxiliar = BaseCalculoICMS_ST

                  vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                  ' [3179a2;vBCFCPST;pFCPST;vFCPST]
                  strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                  strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPST>"
                  strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                  vFCPST_registroTotal = vFCPST_registroTotal + vResultado
              End If
              
              strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 4) & "</pCredSN>"
              strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
              strLinha = strLinha & "</ICMSSN900>"
              AddLinha clcArquivo, strLinha
          
          End Select
        
        End If 'do Simples Nacional
        
        strLinha = "</ICMS>"
        AddLinha clcArquivo, strLinha
        
        If .Fields("IPI").Value > 0 Then
            strLinha = "<IPI>"
            
            strLinha = strLinha & "<cEnq>999</cEnq>"
            
            strLinha = strLinha & "<IPITrib>"
            strLinha = strLinha & "<CST>99</CST>" ' 00-Entrada com recuperação de IPI crédito 49-Outras entradas 50-Saída tributada 99-Outras saídas
            strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Qtde").Value * .Fields("Preço").Value, 2) & "</vBC>" 'Valor da base de cálculo do ipi
            strLinha = strLinha & "<pIPI>" & Replace(FormataValorTexto(.Fields("IPI").Value, 2), ",", ".") & "</pIPI>" 'Aliquota do IPI
            strLinha = strLinha & "<vIPI>" & FormataValorTexto((.Fields("IPI").Value * (.Fields("Qtde").Value * .Fields("Preço").Value)) / 100, 2) & "</vIPI>" 'Valor do IPI
            strLinha = strLinha & "</IPITrib>"
            strLinha = strLinha & "</IPI>"
            AddLinha clcArquivo, strLinha
        End If
        
        If blnSimplesNacional = False Then
          strLinha = "<PIS>"
          strLinha = strLinha & "<PISAliq>"
          strLinha = strLinha & "<CST>" & "01" & "</CST>"
          strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
          strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
          strLinha = strLinha & "<vPIS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualPIS), 2) & "</vPIS>"
          strLinha = strLinha & "</PISAliq>"
          strLinha = strLinha & "</PIS>"

          strLinha = strLinha & "<COFINS>"
          strLinha = strLinha & "<COFINSAliq>"
          strLinha = strLinha & "<CST>" & "01" & "</CST>"
          strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
          strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
          strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualCOFINS), 2) & "</vCOFINS>"
          strLinha = strLinha & "</COFINSAliq>"
          strLinha = strLinha & "</COFINS>"
          AddLinha clcArquivo, strLinha

        Else
          'Pis
          strLinha = "<PIS>"
          strLinha = strLinha & "<PISNT>"
          strLinha = strLinha & "<CST>" & "07" & "</CST>"
          strLinha = strLinha & "</PISNT>"
          strLinha = strLinha & "</PIS>"
          
          'Cofins
          strLinha = strLinha & "<COFINS>"
          strLinha = strLinha & "<COFINSNT>"
          strLinha = strLinha & "<CST>" & "07" & "</CST>"
          strLinha = strLinha & "</COFINSNT>"
          strLinha = strLinha & "</COFINS>"
          AddLinha clcArquivo, strLinha
        End If
        
        If (isConsumidorFinal And isInterestadual And isIEIsento) Then
            Dim vBCUFDest As Double
            Dim pFCPUFDest As Double
            Dim pICMSUFDest As Double
            Dim pICMSInter As Double
            Dim pICMSInterPart As Double
            Dim vFCPUFDest As Double
            Dim vICMSUFDest As Double
            Dim vICMSUFRemet As Double
            Dim difal As Double
          
            If BaseCalculoICMS_ST > 0 Then
              vBCUFDest = BaseCalculoICMS_ST
            Else
              vBCUFDest = .Fields("Qtde").Value * .Fields("Preço").Value
            End If
          
            'pICMSInterPart = 40 - 2016 | 60 - 2017 |80 - 2018 | 100 - 2019 ----
            pFCPUFDest = 2
            pICMSUFDest = rsSaidas.Fields("aliquota_destino").Value
            pICMSInter = rsSaidas.Fields("aliquota_origem").Value
            pICMSInterPart = 60
            vFCPUFDest = 0
            vICMSUFDest = 0
            vICMSUFRemet = 0
          
            'calcular o Fundo de Combate à Pobreza
            'FCP = Base do ICMS * (%FCP / 100)
            'FCP = 1.000,00 * (2,00% / 100)
            'FCP = 1.000,00 * 0,02
            'FCP = 20,00
            vFCPUFDest = ((vBCUFDest * pFCPUFDest) / 100)
          
            'calcular o DIFAL
            'DIFAL = Base do ICMS * ((%Alíquota do ICMS Intra - %Alíquota do ICMS Inter) / 100)
            'DIFAL = 1.000,00 * ((18,00% - 12,00%) / 100)
            'DIFAL = 1.000,00 * (6,00% / 100)
            'DIFAL = 1.000,00 * 0,06
            'DIFAL = 60,00
             difal = vBCUFDest * ((pICMSUFDest - pICMSInter) / 100)
           
            'efetuar a partilha do DIFAL
            'Parte UF Origem = Valor do DIFAL * (%Origem / 100) 20 - 2018|0 - 2019
            'Parte SC = 60,00 * (60,00% / 100)
            'Parte SC = 60, 0 * 0, 60
            'Parte SC = 36, 0
            vICMSUFRemet = difal * (40 / 100)
           
            'Parte que compete ao estado de destino:
            'Parte UF Destino = Valor do DIFAL * (%Destino / 100) 80 - 2018 | 100 - 2019
            'Parte MA = 60,00 * (40,00% / 100)
            'Parte MA = 60, 0 * 0, 40
            'Parte MA = 24, 0
            vICMSUFDest = difal * (60 / 100)
            
            If blnSimplesNacionalGeral = False Then
                ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
                If str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then

                    ' Novo grupo ICMSUFDest
                    ' [3179b;vBCUFDest;vBCFCPUFDest;pFCPUFDest;pICMSUFDest;pICMSInter;pICMSInterPart;vFCPUFDest;vICMSUFDest;vICMSUFRemet]

                    strLinha = "<ICMSUFDest>"
                    
                    vBC_fcp = .Fields("Qtde").Value * .Fields("Preço").Value
                    'vBCUFDest ------------------------
                    strLinha = strLinha & "<vBCUFDest>" & FormataValorTexto(vBC_fcp, 2) & "</vBCUFDest>"
                    'vBCFCPUFDest ------------------------
                    strLinha = strLinha & "<vBCFCPUFDest>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCPUFDest>"
                    'pFCPUFDest ------------------------
                    strLinha = strLinha & "<pFCPUFDest>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCPUFDest>"
                    'pICMSUFDest ------------------------
                    strLinha = strLinha & "<pICMSUFDest>" & FormataValorTexto(AliquotaFCP_Destinatario, 4) & "</pICMSUFDest>"
                    'pICMSInter ------------------------
                    strLinha = strLinha & "<pICMSInter>" & FormataValorTexto(AliquotaFCP_EMIT_DEST, 4) & "</pICMSInter>"

                    'pICMSInterPart ------------------------
                    If iAnoFCP = 2018 Then
                        strLinha = strLinha & "<pICMSInterPart>" & FormataValorTexto("80.00", 2) & "</pICMSInterPart>"   ' 80%
                    Else
                        strLinha = strLinha & "<pICMSInterPart>" & FormataValorTexto("100.00", 2) & "</pICMSInterPart>"  ' 100% para 2019 em diante
                    End If

                    Dim dAux As Double

                    'vFCPUFDest ------------------------
                    dAux = (vFCP_registroTotal * AliquotaFCP_Destinatario) / 100
                    strLinha = strLinha & "<vFCPUFDest>" & FormataValorTexto(dAux, 2) & "</vFCPUFDest>"

                    'vICMSUFDest --------------------------
                    ' DIFAL
                    CalculoFCPEstadosDiferentes = (vFCP_registroTotal * AliquotaFCP_EMIT_DEST) / 100
                    CalculoFCPEstadosIguais = (vFCP_registroTotal * AliquotaFCP_Emitente) / 100
                    CalculoFCP_diferenca = CalculoFCPEstadosIguais - CalculoFCPEstadosDiferentes

                    ' Ano 2018 80% para o Destino   e   20% para o Emitente
                    ' Ano maior que 2018 100% para o Destino
                    If (iAnoFCP = 2018) Then
                        CalculoFCP_participacaoDest = (CalculoFCP_diferenca * 80) / 100   ' 80%
                    Else
                        CalculoFCP_participacaoDest = CalculoFCP_diferenca   ' 100% para 2019 em diante
                    End If
                    strLinha = strLinha & "<vICMSUFDest>" & FormataValorTexto(CalculoFCP_participacaoDest, 2) & "</vICMSUFDest>"

                    ' vICMSUFRemet ------------------------
                    If (iAnoFCP = 2018) Then
                        CalculoFCP_participacaoEmitente = (CalculoFCP_diferenca * 20) / 100   ' 20%
                    Else
                        CalculoFCP_participacaoEmitente = 0   ' 100% para 2019 em diante
                    End If
                    strLinha = strLinha & "<vICMSUFRemet>" & FormataValorTexto(CalculoFCP_participacaoEmitente, 2) & "</vICMSUFRemet>"
                    strLinha = strLinha & "</ICMSUFDest>"
                    AddLinha clcArquivo, strLinha
                End If
            End If
        End If
        
        strLinha = strLinha & "</imposto>"
      
        ' Pilatti Outubro/2017 - linha incluida
        strLinha = strLinha & "</det>"
        ' Pilatti Fim
  
        AddLinha clcArquivo, strLinha
      
      End If 'do tipo de arquivo de integracao
      
      strSQL = ""
      .MoveNext
    Loop
    .Close
  End With
  Set rsEntradasProdutos = Nothing

  rsEntradasAUXILIAR_ICMS.Close
  Set rsEntradasAUXILIAR_ICMS = Nothing
  
  rsEntradasProdutos_CFOP.Close
  Set rsEntradasProdutos_CFOP = Nothing

End Function

'Adiciona a linha na coleção
Private Function AddLinha(ByRef clcArquivo As Collection, ByRef strLinha As String)
  'Insere linha
  
  If iIndex = 0 Then
    clcArquivo.Add strLinha
    iIndex = clcArquivo.Count
  Else
    If (CStr(clcArquivo.Item(iIndex)) = Replace(Replace(strLinha, "/", ""), " ", "")) Then
      clcArquivo.Remove (iIndex)
      iIndex = iIndex - 1
    Else
      clcArquivo.Add strLinha
      iIndex = clcArquivo.Count
    End If
  End If

  If sPadraoArquivoIntegracao = "XML" Then
    sXML_40 = sXML_40 + strLinha
  End If
  
  'Limpa linha
  strLinha = ""

End Function

'24/08/2009 - mpdea
'Formata o campo do tipo CNPJ/CPF com os seus dígitos não significativos
Private Function FormataCnpjCpf(ByVal strDocumento As String) As String
  Dim lngFormato As Long
  
  'Retorna somente os números
  strDocumento = RetornaSomenteNumero(strDocumento)
  'Verifica tipo de documento
  Select Case Len(strDocumento)
    Case 11 'CPF
      lngFormato = 11
    Case Else 'CNPJ
      lngFormato = 14
  End Select
  'Retorna documento formatado
  FormataCnpjCpf = Format(strDocumento, String(lngFormato, "0"))
End Function

'25/08/2009 - mpdea
'Formata o campo do tipo IE verificando se é Isento ou não contribuinte
Private Function FormataIE(ByVal strDocumento As String, ByVal blnPessoaFisica As Boolean) As String
  Dim strReturn As String
  Dim intX As Integer
  Const SPECIAL_CHARS As String = " '!@#$%¨&*()-_|\/?,.;<>:[]{}"
  
  'Verifica se não é contribuinte do ICMS
  If Trim(strDocumento) = "" Or blnPessoaFisica Then
    strReturn = ""
  'Verifica se é isento
  ElseIf InStr(UCase(strDocumento), "ISENTO") > 0 Or InStr(UCase(strDocumento), "ISENTA") > 0 Then
    strReturn = "ISENTO"
  'Formata IE
  Else
    'Remove caracteres especiais
    For intX = 1 To Len(SPECIAL_CHARS)
      strDocumento = Replace(strDocumento, Mid(SPECIAL_CHARS, intX, 1), "")
    Next
    strReturn = strDocumento
  End If
  
  'Retorno
  FormataIE = strReturn
End Function

'25/08/2009 - mpdea
'Formata o número de acordo com a quantidade de dígitos informada
Private Function FormataNumero(ByVal strNumero As String, ByVal lngDigitos As Long) As String
  FormataNumero = Format(RetornaSomenteNumero(strNumero), String(lngDigitos, "0"))
End Function

'Formata o valor de acordo com o número de casas decimais e substitui separador decimal por ponto
Private Function FormataValorTexto(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTexto = Replace(Format(dblValor, "#0." & String(lngCasasDecimais, "0")), ",", ".")
End Function

Private Function FormataValorTextoComVirgula(ByVal dblValor As Double, Optional ByVal lngCasasDecimais As Long = 4) As String
  FormataValorTextoComVirgula = Format(dblValor, "#0." & String(lngCasasDecimais, "0"))
End Function


'Retorna somente a parte numérica de um texto
Private Function RetornaSomenteNumero(ByVal strTexto) As String
  Dim intX As Integer
  Dim strPart As String
  Dim strReturn As String
  
  strReturn = ""
  For intX = 1 To Len(strTexto)
    strPart = Mid(strTexto, intX, 1)
    If IsNumeric(strPart) Then
      strReturn = strReturn & strPart
    End If
  Next

  RetornaSomenteNumero = strReturn
End Function

'Grava o arquivo
Private Function GravarArquivo(ByRef clcArquivo As Collection, ByVal strNomeArquivo As String)
  Dim intFreeFile As Integer
  Dim varLinha As Variant
  
  'Gravação em arquivo
  'Abre arquivo
  intFreeFile = FreeFile
  Open strNomeArquivo For Output As #intFreeFile
  For Each varLinha In clcArquivo
    'Grava linhas
    Print #intFreeFile, CStr(varLinha)
  Next
  'Fecha arquivo
  Close #intFreeFile

End Function

'20/09/2009 - mpdea
'Obtém a chave de acesso para NFe gerada
Private Function GetChaveAcesso(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal TipoMovimento As Integer) As String
  Dim rstNFe As Recordset
  Dim strSQL As String
  Dim strRet As String
  
  strSQL = "SELECT ChaveAcesso "
  strSQL = strSQL & "FROM NFe "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequencia = " & Sequencia & " AND TipoMovimento = " & TipoMovimento
  Set rstNFe = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstNFe
    If Not (.BOF And .EOF) Then
      strRet = .Fields("ChaveAcesso").Value & ""
    End If
    .Close
  End With
  Set rstNFe = Nothing
  
  'Retorno
  GetChaveAcesso = strRet

End Function

'29/11/2010 - Andrea
'Obtém o tipo de emissao para NFe gerada
Private Function GetTipoEmissao(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal TipoMovimento As Integer) As Integer
  Dim rstNFe As Recordset
  Dim strSQL As String
  Dim intRet As Integer
  
  strSQL = "SELECT FormaEmissao "
  strSQL = strSQL & "FROM NFe "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequencia = " & Sequencia & " AND TipoMovimento = " & TipoMovimento
  Set rstNFe = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstNFe
    If Not (.BOF And .EOF) Then
      intRet = .Fields("FormaEmissao").Value & ""
    End If
    .Close
  End With
  Set rstNFe = Nothing
  
  'Retorno
  GetTipoEmissao = intRet

End Function


'20/09/2009 - mpdea
'Salva registro da NFe gerada
Private Sub SaveEnvioNFe_4_0(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal TipoMovimento As Integer, _
  ByVal DataHoraEnvio As Date, ByVal Status As Integer, ByVal Ambiente As Byte, ByVal FormaEmissao As Byte, _
  ByVal Numero As Long, ByVal Serie As String, ByVal Modelo As String, ByVal ChaveAcesso As String, ByVal TpEmiss As String)
  
  Dim rstNFe As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT * "
  strSQL = strSQL & "FROM NFe "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequencia = " & Sequencia & " AND TipoMovimento = " & TipoMovimento
  Set rstNFe = db.OpenRecordset(strSQL, dbOpenDynaset)
  With rstNFe
    If (.BOF And .EOF) Then
      .AddNew
      .Fields("Filial").Value = Filial
      .Fields("Sequencia").Value = Sequencia
      .Fields("TipoMovimento").Value = TipoMovimento
      .Fields("Status").Value = Status
      .Fields("Ambiente").Value = Ambiente
      .Fields("FormaEmissao").Value = FormaEmissao
      .Fields("Numero").Value = Numero
      .Fields("Serie").Value = Serie
      .Fields("Modelo").Value = Modelo
      .Fields("ChaveAcesso").Value = ChaveAcesso
      .Fields("ProtocoloAutorizacao").Value = ""
      .Fields("DataHoraAutorizacao").Value = Nothing
      .Fields("ProtocoloCancelamento").Value = ""
      .Fields("DataHoraCancelamento").Value = Nothing
    Else
      .Edit
      .Fields("FormaEmissao").Value = TpEmiss
      .Fields("ChaveAcesso").Value = ChaveAcesso
    End If
    .Fields("DataHoraEnvio").Value = DataHoraEnvio
    .Update
    .Close
  End With
  Set rstNFe = Nothing

End Sub


Private Function CalculaDV(ByRef strChaveNFe As String) As String
  Dim intResultado As Integer
  Dim intDigito As Integer
  Dim intDV As Integer
  Dim i As Integer
  Dim F As Integer
  
  intDigito = 0
  i = 4
  
  For F = 1 To 43
    intResultado = intResultado + i * Mid(strChaveNFe, F, 1)
    i = i - 1
    If i = 1 Then
      i = 9
    End If
  Next F
  
  intDV = (intResultado Mod 11)
  
  If intDV = 0 Or intDV = 1 Then
    intDV = 0
  Else
    intDV = 11 - intDV
  End If
  
  CalculaDV = intDV
    
End Function

