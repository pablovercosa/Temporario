VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const VERSAO_LAYOUT_CANCELAMENTO As String = "1.07"
'Private Const VERSAO_PROCESSO_EMISSAO = "3.10"
Private Const VERSAO_PROCESSO_EMISSAO = "4.00"
Private m_bln_tem_desconto As Boolean
Private isDescontoItem As Boolean
Private m_clcMensagens As Collection
Dim ServicesDescription As String
Private CFOP_Generico As String
Private dbl_Desconto_ProdutoACUMULADO As Double

Public total_produto As Double
Public TotalAproxImpostos As Double

Private sEstadoFilial As String

Private str_FCP_UF_emitente As String
Private str_FCP_INS_EST_emitente As String
Private str_FCP_CNPJ_CPF_destinatario As String
Private str_FCP_RG_INS_EST_destinatario As String
Private str_FCP_UF_destinatario As String

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
Dim vFCP_registroTotal As Double
Dim vFCPST_registroTotal As Double
Dim vFCPSTRet_registroTotal As Double
Dim vResultado As Double
Dim blnSimplesNacionalGeral As Boolean
'Dim rsSaidas As Recordset
Dim str_sql As String
Dim bDescontoNaVendaRapida_por_valorOUperc As Boolean

Dim bTemDescontoNaVenda As Boolean
Dim vTemDescontoNaVendaPercentual As Double
Dim vAuxTotalSaida As Double
Dim vSOMA_TOTAL_DESCONTO As Double

Dim localTipoSituacaoTributariaPIS As Integer
Dim bTEM_ICMS_3179a As Boolean
Dim vlSOMA_DO_ICMS As Double
Dim s_vlSOMA_DO_ICMS_BaseDeCalculo As String
Dim v_vlSOMA_DO_ICMS_BaseDeCalculo As Double

Private sCNPJ_EmitenteXML As String
Private sXML_40 As String
Private sRetFazenda As String

Private iDetNumItemXML As Long

Private inst_dblPercentualPIS As Double
Private inst_dblPercentualCofins As Double

Private aliquotaICMS_tab_ICMS_PERC_ESTADOS As String
Private bo_AliquotaICMS_interestadual As Boolean

'Private Declare Function Autorizar_Xml Lib "NfceInterface.dll" (ByVal cnpj As String, ByVal xml As String) As String

Public Function EnviarXML_SEFAZ(ByVal pSequencia As Long, ByVal pFilial As Integer, ByVal pCNPJEmitente As String) As Integer
On Error GoTo ErroTratamento

    gRetOID = 0

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

    Dim sRetorno As String
    Dim rsSaidaSEFAZ As Recordset
    
    ' Via SOAP
    If bSoapClient_MSSoapInit_NFCe = False Then
      Set soapclient_NFCe = New SoapClient30
      soapclient_NFCe.MSSoapInit sSoapClient_MSSoapInit_NFCe
      soapclient_NFCe.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL_NFCe
      bSoapClient_MSSoapInit_NFCe = True
    End If
    
    sRetorno = soapclient_NFCe.Autoriza_Xml(pCNPJEmitente, sXML_40)

    Dim iIndice01 As Integer
    Dim iIndice02 As Integer
    Dim strLinha As String

    sRetorno = Replace(sRetorno, vbCrLf, "")
    sRetorno = Replace(sRetorno, "|", "|")
    sRetFazenda = ""
    sRetFazenda = sRetorno
    
    Dim sMensagemTratadaNFCe As String
    sMensagemTratadaNFCe = ""

    If InStr(1, sRetorno, "<emContingencia>false</emContingencia>") > 0 Then
        If InStr(sRetorno, "<statusAutorizacao>OK</statusAutorizacao>") Then
            'tudo ok com o CUPOM FISCAL
            Set rsSaidaSEFAZ = db.OpenRecordset("Select * from [Saídas] where Filial = " & pFilial & " And Sequência = " & pSequencia & "")

            rsSaidaSEFAZ.Edit
            rsSaidaSEFAZ!retNFCe = sRetorno
            rsSaidaSEFAZ.Update
            rsSaidaSEFAZ.Close
            Set rsSaidaSEFAZ = Nothing
            
            'sSql = "UPDATE Saídas Set retNFCe='" & sRetorno & "', NFCe_contingencia_num=0 Where Sequência=" & pSequencia & " And Filial=" & pFilial
            'db.Execute sSql

            EnviarXML_SEFAZ = 1 'Enviar para impressão do CUPOM
        Else
            iIndice01 = InStr(1, sRetorno, "<exMessage>")
            
            If iIndice01 > 0 Then
                iIndice02 = InStr(1, sRetorno, "</exMessage>")
                
                If iIndice02 > 0 Then
                    sMensagemTratadaNFCe = TratarMsgRetornoNFCe_Sefaz(sRetorno)
                    If sMensagemTratadaNFCe <> "" Then
                        sMensagemTratadaNFCe = sMensagemTratadaNFCe & vbCrLf & vbCrLf & "Mensagem Original:" & vbCrLf
                    End If
                End If
                
                strLinha = Mid$(sRetorno, iIndice01 + 11, iIndice02 - (iIndice01 + 11))
                MsgBox sMensagemTratadaNFCe & strLinha, vbInformation, "Atenção: CUPOM FISCAL inconsistente"
            Else
                iIndice01 = InStr(1, sRetorno, "<detalheAutorizacao>")
                If iIndice01 > 0 Then
                    iIndice02 = InStr(1, sRetorno, "</detalheAutorizacao>")
                    
                    If iIndice02 > 0 Then
                        sMensagemTratadaNFCe = TratarMsgRetornoNFCe_Sefaz(sRetorno)
                        If sMensagemTratadaNFCe <> "" Then
                            sMensagemTratadaNFCe = sMensagemTratadaNFCe & vbCrLf & vbCrLf & "Mensagem Original:" & vbCrLf
                        End If
                    End If
                
                    strLinha = Mid$(sRetorno, iIndice01 + 20, iIndice02 - (iIndice01 + 20))
                    MsgBox sMensagemTratadaNFCe & strLinha, vbInformation, "Atenção: CUPOM FISCAL inconsistente"
                Else
                    MsgBox sRetorno, vbInformation, "Atenção: CUPOM FISCAL inconsistente"
                End If
            End If
            gStringRetorno = sRetorno
            EnviarXML_SEFAZ = 0 'Não enviar para impressão do CUPOM
        End If

    Else
        '***********************
        'Em contingencia de NFCe
        
        MsgBox sRetorno, vbInformation, "CUPOM FISCAL EM CONTINGÊNCIA"
        
        Dim sNumNFCe As String
        Dim sSerieNFCe As String
        Dim sNumNFCeRetCont As String
        Dim sStatusNFCeRetCont As String
        Dim sSerieNFCeRetCont As String
        Dim sChaveNFCeRetCont As String
        
        iIndice01 = InStr(1, sXML_40, "<nNF>")
        If iIndice01 > 0 Then
            iIndice02 = InStr(1, sXML_40, "</nNF>")
            strLinha = Mid$(sXML_40, iIndice01 + 5, iIndice02 - (iIndice01 + 5))
        End If
        sNumNFCe = strLinha
        
        iIndice01 = InStr(1, sXML_40, "<serie>")
        If iIndice01 > 0 Then
            iIndice02 = InStr(1, sXML_40, "</serie>")
            strLinha = Mid$(sXML_40, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
        End If
        sSerieNFCe = strLinha
        
        
        iIndice01 = InStr(1, sRetorno, "<numero>")
        If iIndice01 > 0 Then
            iIndice02 = InStr(1, sRetorno, "</numero>")
            strLinha = Mid$(sRetorno, iIndice01 + 8, iIndice02 - (iIndice01 + 8))
        Else
            strLinha = ""
        End If
        sNumNFCeRetCont = strLinha
        
        iIndice01 = InStr(1, sRetorno, "<statusAutorizacao>")
        If iIndice01 > 0 Then
            iIndice02 = InStr(1, sRetorno, "</statusAutorizacao>")
            strLinha = Mid$(sRetorno, iIndice01 + 19, iIndice02 - (iIndice01 + 19))
        Else
            strLinha = ""
        End If
        sStatusNFCeRetCont = strLinha
        
        iIndice01 = InStr(1, sRetorno, "<serie>")
        If iIndice01 > 0 Then
            iIndice02 = InStr(1, sRetorno, "</serie>")
            strLinha = Mid$(sRetorno, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
        Else
            strLinha = ""
        End If
        sSerieNFCeRetCont = strLinha
        
        iIndice01 = InStr(1, sRetorno, "<chave>")
        If iIndice01 > 0 Then
            iIndice02 = InStr(1, sRetorno, "</chave>")
            strLinha = Mid$(sRetorno, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
        Else
            strLinha = ""
        End If
        sChaveNFCeRetCont = strLinha

        'Dim X As Variant
        'X = InputBox$("Informe o tamanho: ", "Tamanho")

        If InStr(1, sRetorno, "<statusAutorizacao>OK</statusAutorizacao>") > 0 Or InStr(1, sRetorno, "<statusAutorizacao>Pendente</statusAutorizacao>") > 0 Then
            'tudo ok com o CUPOM FISCAL ou pelo menos encaminhado via contingencia e recebido OK pela fazenda
            'sSql = "UPDATE Saídas Set retNFCe='" & sRetorno & "' Where Sequência=" & pSequencia & " And Filial=" & pFilial
            'MsgBox "Aqui 4: " & sSql, vbInformation, "Teste"
            'db.Execute sSql
            
            Set rsSaidaSEFAZ = db.OpenRecordset("Select * from [Saídas] where Filial = " & pFilial & " And Sequência = " & pSequencia & "")
      
            rsSaidaSEFAZ.Edit
            rsSaidaSEFAZ!retNFCe_contingencia = sRetorno
            rsSaidaSEFAZ!NFCe_contingencia_status = sStatusNFCeRetCont
            rsSaidaSEFAZ!NFCe_contingencia_num = sNumNFCeRetCont
            rsSaidaSEFAZ!NFCe_contingencia_serie = sSerieNFCeRetCont
            rsSaidaSEFAZ!NFCe_contingencia_chave = sChaveNFCeRetCont
            rsSaidaSEFAZ.Update
            rsSaidaSEFAZ.Close
            Set rsSaidaSEFAZ = Nothing

            EnviarXML_SEFAZ = 1 'Enviar para impressão do CUPOM
        Else
            iIndice01 = InStr(1, sRetorno, "<exMessage>")
            If iIndice01 > 0 Then
                iIndice02 = InStr(1, sRetorno, "</exMessage>")
                strLinha = Mid$(sRetorno, iIndice01 + 11, iIndice02 - (iIndice01 + 11))
                MsgBox strLinha, vbInformation, "Atenção: CUPOM FISCAL"
            Else
                MsgBox sRetorno, vbInformation, "Atenção: CUPOM FISCAL"
            End If
            EnviarXML_SEFAZ = 0 'Não enviar para impressão do CUPOM

            'sSql = "UPDATE Saídas Set retNFCe='" & sRetorno & "' Where Sequência=" & pSequencia & " And Filial=" & pFilial
            'db.Execute sSql
            Set rsSaidaSEFAZ = db.OpenRecordset("Select * from [Saídas] where Filial = " & pFilial & " And Sequência = " & pSequencia & "")
      
            rsSaidaSEFAZ.Edit
            rsSaidaSEFAZ!retNFCe_contingencia = sRetorno
            rsSaidaSEFAZ!NFCe_contingencia_status = sStatusNFCeRetCont
            rsSaidaSEFAZ!NFCe_contingencia_num = sNumNFCeRetCont
            rsSaidaSEFAZ!NFCe_contingencia_serie = sSerieNFCeRetCont
            rsSaidaSEFAZ!NFCe_contingencia_chave = sChaveNFCeRetCont
            rsSaidaSEFAZ.Update
            rsSaidaSEFAZ.Close
            Set rsSaidaSEFAZ = Nothing
        End If

        Dim lnumNFCe As Long
        Dim lnumNFCeRetCont As Long

        lnumNFCe = CLng(sNumNFCe)
        lnumNFCeRetCont = CLng(sNumNFCeRetCont)

        If InStr(1, sRetorno, "<indicaNovo>false</indicaNovo>") > 0 Then
            'indicaNovoRegistro é false...então o que vale agora é a série, númeroNFCe e Chave que veio no retorno da benefix
            
            If sSerieNFCe = sSerieNFCeRetCont Then
                If lnumNFCeRetCont = lnumNFCe + 1 Then
                    db.Execute "UPDATE [Parâmetros Filial] set UltimaNFCe=UltimaNFCe+1 Where Filial = " & gnCodFilial
                End If
            Else
                db.Execute "UPDATE [Parâmetros Filial] set UltimaNFCe=UltimaNFCe-1 Where Filial = " & gnCodFilial
            End If
        Else
            'indicaNovoRegistro é true...então a serie e numero original foi utilizado...

            If sSerieNFCe = sSerieNFCeRetCont Then
                If lnumNFCeRetCont = lnumNFCe + 1 Then
                    db.Execute "UPDATE [Parâmetros Filial] set UltimaNFCe=UltimaNFCe+1 Where Filial = " & gnCodFilial
                End If
            End If
        End If
    End If
    
    Exit Function

ErroTratamento:

    MsgBox "Erro na função EnviarXML_SEFAZ " & Err.Number & " - " & Err.Description, vbInformation, "Erro"
End Function

Public Function TratarMsgRetornoNFCe_Sefaz(ByVal pMensagem As String) As String
On Error GoTo ErroTratamento
    Dim sMensagemTratada As String
    Dim iIndice As Long
    Dim iIndice01 As Long
    Dim iIndice02 As Long
    Dim sCampo As String
    Dim sCampo2 As String
    
    TratarMsgRetornoNFCe_Sefaz = ""
        
    iIndice = InStr(1, pMensagem, "element is invalid - The value")
    If iIndice > 0 Then
        iIndice = InStr(1, pMensagem, "is invalid according to its datatype")

        If iIndice > 0 Then
            iIndice01 = InStr(1, pMensagem, "www.portalfiscal.inf.br/nfe:")
            If iIndice01 > 0 Then
                iIndice02 = InStr(iIndice01 + 15, pMensagem, "'")
                'Obter nome do campo...
                sCampo = Mid(pMensagem, iIndice01 + 28, iIndice02 - (iIndice01 + 28))

                'Obter o valor do campo inconsistente
                iIndice01 = InStr(1, pMensagem, "element is invalid - The value '")
                If iIndice01 > 0 Then
                    iIndice02 = InStr(iIndice01 + 33, pMensagem, "'")
                    sCampo2 = Mid(pMensagem, iIndice01 + 33, iIndice02 - (iIndice01 + 33))
                End If

                If UCase(sCampo) = "NCM" Then
                    TratarMsgRetornoNFCe_Sefaz = sCampo + " inválido ou inexistente em algum dos produtos. **Revise na tela Cadastro de Produtos, aba Impostos, botão Tributos do produto, campo [NCM]."
                ElseIf UCase(sCampo) = "ORIG" Then
                    TratarMsgRetornoNFCe_Sefaz = sCampo + " inválido ou inexistente em algum dos produtos. **Revise na tela Cadastro de Produtos, aba Impostos, botão Tributos do produto, campo [Origem + Situação Tributária Saída]."
                Else
                    TratarMsgRetornoNFCe_Sefaz = sCampo + " " + sCampo2 + " inválido ou inexistente em algum dos produtos"
                End If
            End If
        End If
        Exit Function
    End If
        
    iIndice = InStr(1, pMensagem, "Schema XML: The element 'ICMS'")
    If iIndice > 0 Then
        iIndice = InStr(1, pMensagem, "ICMS00, ICMS10, ICMS20")
        If iIndice > 0 Then
            TratarMsgRetornoNFCe_Sefaz = "Algum dos produtos não esta destacada a Situação Tributária. **Revise na tela Cadastro de Produtos, aba Impostos, botão Tributos do produto, campo [Origem + Situação Tributária Saída]."
        Else
            TratarMsgRetornoNFCe_Sefaz = "**Revise em algum dos produtos algo relacionado a Situação Tributária. Na tela Cadastro de Produtos, aba Impostos, botão Tributos do produto."
        End If
        
        Exit Function
    End If
    
    Exit Function

ErroTratamento:
  MsgBox "Alerta de erro (porém sem interferência na Nota Fiscal) na função TratarMsgRetornoNFCe " & Err.Number & " - " & Err.Description, vbInformation, "Atenção"
End Function

Public Function CancelarXML_SEFAZ(ByVal pSequencia As Long, ByVal pCNPJEmitente As String, ByVal pChave As String, ByVal pJustificativa As String)
On Error GoTo ErroTratamento

    Dim sRetorno As String

    gRetOID = 0

    ' Via SOAP
    If bSoapClient_MSSoapInit_NFCe = False Then
      Set soapclient_NFCe = New SoapClient30
      soapclient_NFCe.MSSoapInit sSoapClient_MSSoapInit_NFCe
      soapclient_NFCe.ConnectorProperty("EndPointURL") = sSoapClient_ConnectorProperty_EndPointURL_NFCe
      bSoapClient_MSSoapInit_NFCe = True
    End If
    
    sRetorno = soapclient_NFCe.Cancela_Xml(pCNPJEmitente, pChave, pJustificativa)

    Dim iIndice01 As Integer
    Dim iIndice02 As Integer
    Dim strLinha As String
    Dim sSql As String
    
    sRetorno = Replace(sRetorno, vbCrLf, "")
    sRetorno = Replace(sRetorno, "|", "|")

    sSql = "UPDATE Saídas Set retNFCe='" & sRetorno & "' Where Sequência=" & pSequencia & " And Filial=" & gnCodFilial

    db.Execute sSql
    
    If InStr(sRetorno, "<statusCancelamento>OK</statusCancelamento>") Then
        gRetOID = 0
    Else
        Dim sAuxCanc As String
        iIndice01 = InStr(1, sRetorno, "<detalheCancelamento>")
        If iIndice01 > 0 Then
            iIndice02 = InStr(1, sRetorno, "</detalheCancelamento>")
            sAuxCanc = Mid$(sRetorno, iIndice01 + 21, iIndice02 - (iIndice01 + 21))
        End If
    
        MsgBox "CANCELAMENTO NÃO REALIZADO! " + sAuxCanc, vbInformation, "Atenção: Cancelamento de Cupom Fiscal"
        gRetOID = -1
    End If

Exit Function

ErroTratamento:

    MsgBox "Erro na função CancelarXML_SEFAZ " & Err.Number & " - " & Err.Description, vbInformation, "Erro"

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

Public Function EnviaNFCe(ByVal Sequencia As Long, ByVal Vdesc As Double) As Boolean
On Error GoTo TrataErro:

  Dim clcArquivo As New Collection
  Dim strSQL As String
  Dim dblPeso As Double
  Dim rsSaida As Recordset
  Dim rsFilial As Recordset
  Dim rsCliFor As Recordset
  Dim rsOpSaida As Recordset
  Dim rsTransportadora As Recordset
  Dim strLinha As String
  Dim strChaveNFCe As String
  Dim strCodigoAleatorioChaveNFCe As String
  Dim intDV As Integer
  Dim lngNumNFCe As Long
  Dim blnShowObs As Boolean
  Dim frmX As Form
  Dim intX As Integer
  Dim intSerieNFC As Integer
  Dim dblPercentualPIS As Double
  Dim dblPercentualCofins As Double
  Dim dblAliquotaSimples As Double
  Dim dblPercentualReducaoSimplesNacional As Double
  Dim TpEmis As Integer
  Dim mostraCliente As Boolean
  Dim Resposta As Integer
  mostraCliente = False
  Dim mostrouCliente As Boolean
  Dim somenteCPF_CNPJ As Boolean
  somenteCPF_CNPJ = False
  Dim mostrouSomenteCpf_CNPJ As Boolean
  Dim CPF_CnpjCliente As String
  Dim IE_Cliente As String
  Dim PessoaJuridica As String
  PessoaJuridica = False
  Dim IndicadorIE As String
  Dim IE As String
  Dim strNomeCliente As String
  Dim sTrataDados As String
  Dim Endereco_Emitente As String
  Dim sUFDestinatario As String
  
  Dim strPastaEnvio As String
  
  gsCPF_Cnpj = ""
  
  bDescontoNaVendaRapida_por_valorOUperc = False
  
  sRetornoEnvioNFCe = ""

  dbl_Desconto_ProdutoACUMULADO = 0
  str_FCP_UF_destinatario = ""

  bTEM_ICMS_3179a = False
  
  inst_dblPercentualPIS = 0
  inst_dblPercentualCofins = 0

  TotalAproxImpostos = 0
  
  strNomeCliente = ""
  
  aliquotaICMS_tab_ICMS_PERC_ESTADOS = ""
  bo_AliquotaICMS_interestadual = False
  
  
  TpEmis = 1
  
  Set rsSaida = db.OpenRecordset("Select * from [Saídas] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
  If rsSaida("Efetivada") = False Then
    MsgBox ("NFC-e só pode ser emitido a partir de uma venda efetivada. Encontre uma venda efetivada.")
    Exit Function
  End If
  
  
  '************************************************************************************
  'TRATAMENTO APENAS PARA REIMPRESSAO DO CUPOM
  If InStr(rsSaida("retNFCe").Value, "<statusAutorizacao>OK</statusAutorizacao>") _
      Or InStr(rsSaida("retNFCe").Value, "<statusAutorizacao>Pendente</statusAutorizacao>") _
      Or InStr(rsSaida("retNFCe").Value, "<detalheAutorizacao>100") Then
      'Apenas enviar o CUPOM para a impressora, pois já foi AUTORIZADO pela Fazenda ou Enviado com sucesso via Contingencia

      'Obter o arquivo
      Dim sXMLImpr As String
      Dim sNomeArqXML As String

      Set rsFilial = db.OpenRecordset("Select * from [Parâmetros Filial] Where Filial = " & gnCodFilial & "")

      'Validações
      strPastaEnvio = Trim(rsFilial.Fields("PastaEnvioNfe").Value)
      If strPastaEnvio = "" Then
        MsgBox "Pasta de envio não configurada para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
        rsFilial.Close
        Set rsFilial = Nothing
        Exit Function
      Else
        If Right(strPastaEnvio, 1) <> "\" Then
          strPastaEnvio = strPastaEnvio & "\"
        End If
      End If

      'Abrir arquivo .xml
      Dim ff As Integer
      ff = FreeFile
      
      sNomeArqXML = strPastaEnvio & "NFCeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_SeqVenda" & Sequencia & ".xml"
      Open sNomeArqXML For Input As #ff

      Dim Linha As String
      While EOF(ff) = False
          Linha = ""
          Line Input #ff, Linha
          sXMLImpr = sXMLImpr + Linha
      Wend
      Close #ff
      sXMLImpr = Replace(sXMLImpr, vbCrLf, "")

      Dim rsTesteConexaoImpr As Recordset
On Error GoTo ApenasImprimir
      Set rsTesteConexaoImpr = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = ''")
ApenasImprimir:

      If Err.Number <> 0 Then
          Set BancoPDV = OpenDatabase(rsFilial("BancoPDV").Value & "\QuickStore.mdb", False, False, ";PWD=" & gsGetPValue())
      End If

      Dim rsNFCe_JobImpr As Recordset
      Set rsNFCe_JobImpr = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = 'XXXXXXX' And cnpj = '999999'")

      If Left(rsFilial.Fields("EnderecoComplemento").Value, 60) <> "" Then
        Endereco_Emitente = Left(rsFilial.Fields("Endereço").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoNumero").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoComplemento").Value, 60) & ", " & Left(rsFilial.Fields("Bairro").Value, 60) & ", " & Left(rsFilial.Fields("Cidade").Value, 60)
      Else
        Endereco_Emitente = Left(rsFilial.Fields("Endereço").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoNumero").Value, 60) & ", " & Left(rsFilial.Fields("Bairro").Value, 60) & ", " & Left(rsFilial.Fields("Cidade").Value, 60)
      End If

      Dim iIndice01 As Integer
      Dim iIndice02 As Integer
      iIndice01 = InStr(1, sXMLImpr, "<serie>")
      If iIndice01 > 0 Then
          iIndice02 = InStr(1, sXMLImpr, "</serie>")
          intSerieNFC = Mid$(sXMLImpr, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
      End If

      iIndice01 = InStr(1, sXMLImpr, "<nNF>")
      If iIndice01 > 0 Then
          iIndice02 = InStr(1, sXMLImpr, "</nNF>")
          lngNumNFCe = Mid$(sXMLImpr, iIndice01 + 5, iIndice02 - (iIndice01 + 5))
      End If

      iIndice01 = InStr(1, sXMLImpr, "Id=""NFe")
      If iIndice01 > 0 Then
          iIndice02 = InStr(iIndice01 + 10, sXMLImpr, """")
          strChaveNFCe = Mid$(sXMLImpr, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
      End If

      iIndice01 = InStr(1, sXMLImpr, "<dest>")
      If iIndice01 > 0 Then
          iIndice01 = InStr(iIndice01, sXMLImpr, "<CPF>")
          If iIndice01 > 0 Then
              iIndice02 = InStr(iIndice01, sXMLImpr, "</CPF>")
              CPF_CnpjCliente = Mid$(sXMLImpr, iIndice01 + 5, iIndice02 - (iIndice01 + 5))
          Else
              iIndice01 = InStr(iIndice01, sXMLImpr, "<CNPJ>")
              iIndice02 = InStr(iIndice01, sXMLImpr, "</CNPJ>")
              CPF_CnpjCliente = Mid$(sXMLImpr, iIndice01 + 6, iIndice02 - (iIndice01 + 6))
          End If

          iIndice01 = InStr(iIndice01, sXMLImpr, "<xNome>")
          If iIndice01 > 0 Then
              iIndice02 = InStr(iIndice01, sXMLImpr, "</xNome>")
              strNomeCliente = Mid$(sXMLImpr, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
          End If
      End If

      If rsNFCe_JobImpr.EOF Then
        rsNFCe_JobImpr.AddNew
        rsNFCe_JobImpr!CNPJ = FormataCnpjCpf(rsFilial.Fields("CGC").Value)
        rsNFCe_JobImpr!xml = sXMLImpr                                                           'ok
        rsNFCe_JobImpr!Tipo = 1                                                                 'ver o que tipo=1
        rsNFCe_JobImpr!Serie = intSerieNFC                                                      'ok
        rsNFCe_JobImpr!N_NF = lngNumNFCe                                                        'ok
        rsNFCe_JobImpr!Chave = strChaveNFCe                                                     'ok
        rsNFCe_JobImpr!CPF = CPF_CnpjCliente                                                    'ok
        rsNFCe_JobImpr!Nome_Consumidor = strNomeCliente                                         'ok
        rsNFCe_JobImpr!Data_Emissao = rsSaida.Fields("DataEmissaoNota").Value                   'ok
        rsNFCe_JobImpr!Total_Tributos = FormataValorTexto(rsSaida.Fields("TotalNCM").Value, 2)  'ok
        rsNFCe_JobImpr!Nome_Emitente = Left(rsFilial.Fields("Razão Social").Value, 60)          'ok
        rsNFCe_JobImpr!Endereco_Emitente = Endereco_Emitente                                    'ok
        rsNFCe_JobImpr!IE_Emitente = FormataIE(rsFilial.Fields("Inscrição").Value, False)       'ok
        rsNFCe_JobImpr!retFazenda = rsSaida("retNFCe").Value                                    'ok
        rsNFCe_JobImpr.Update
      End If
      rsNFCe_JobImpr.Close
      Set rsNFCe_JobImpr = Nothing

      MsgBox "Cupom Fiscal enviado novamente para a impressora!", vbInformation, "CUPOM FISCAL"

      Exit Function
  End If
  'FIM DO TRATAMENTO APENAS PARA REIMPRESSAO DO CUPOM
  '************************************************************************************
  
  
  If rsSaida("NFCe") > 0 Then
    Resposta = MsgBox("Um NFCe já foi emitido para essa movimentação. Deseja reenviar?", vbYesNo + vbQuestion, "Reenviar")
    If Resposta = vbNo Then Exit Function
  End If
  
  Set rsFilial = db.OpenRecordset("Select * from [Parâmetros Filial] Where Filial = " & gnCodFilial & "")
  Set rsCliFor = db.OpenRecordset("Select * from [Cli_For] Where Código = " & rsSaida("Cliente") & "")
  Set rsOpSaida = db.OpenRecordset("Select * from [Operações Saída] where Código = " & rsSaida("Operação"))
  
  If Not (rsCliFor.EOF And rsCliFor.BOF) Then
      If Not IsNull(rsCliFor.Fields("Estado").Value) Then
          sUFDestinatario = rsCliFor.Fields("Estado").Value
      Else
          sUFDestinatario = ""
      End If
  End If
  
  
  'Validações
  strPastaEnvio = Trim(rsFilial.Fields("PastaEnvioNfe").Value)
  If strPastaEnvio = "" Then
    MsgBox "Pasta de envio não configurada para a Filial: " & gnCodFilial, vbExclamation, "Atenção"
    rsFilial.Close
    Set rsFilial = Nothing
    Exit Function
  Else
    If Right(strPastaEnvio, 1) <> "\" Then
      strPastaEnvio = strPastaEnvio & "\"
    End If
  End If
  
  CFOP_Generico = rsOpSaida.Fields("Código Fiscal").Value
  CFOP_Generico = Trim(CFOP_Generico)
  
  If rsOpSaida("PermiteMostrarCliente") = True Then
    Resposta = MsgBox("Deseja mostrar os dados do cliente neste NFCe?", vbYesNo + vbQuestion, "Enviar")
    If Resposta = vbYes Then
      mostraCliente = True
'''      Resposta = MsgBox("Deseja mostrar somente o CPF/CNPJ e Inscrição Estadual do cliente neste NFCe?", vbYesNo + vbQuestion, "Enviar")
'''      If Resposta = vbYes Then somenteCPF_CNPJ = True
      somenteCPF_CNPJ = True
    End If
  End If
  
  If GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") = 0 Then
    MsgBox "Código IBGE inválido para a UF da Filial: " & rsFilial.Fields("Filial").Value, vbExclamation, "Atenção"
    rsFilial.Close
    Set rsFilial = Nothing
    Exit Function
  End If
  
  If Not IsNull(rsFilial.Fields("TipoSituacaoTributariaPIS").Value) Then
      gTipoSituacaoTributariaPIS = rsFilial.Fields("TipoSituacaoTributariaPIS").Value
  Else
      gTipoSituacaoTributariaPIS = 0
  End If
  
  ' Pilatti Outubro/17
  sEstadoFilial = rsFilial.Fields("Estado").Value
  
  ' *******************************************
  ' PILATTI Outubro/17
  ' Tratar horario de verão
  Dim sFuso As String
  sFuso = ""
  'bHorarioDeVerao = True    ' Sim, estamos em HORARIO DE VERÃO
  func_horarioFuso bHorarioDeVerao, sEstadoFilial, sFuso
  ' *******************************************
  
  'Valida CNPJ
  If (Trim(rsFilial.Fields("CGC").Value) & "") = "" Or Not Valid_CGC(rsFilial.Fields("CGC").Value & "") Then
    MsgBox "CNPJ inválido para a Filial: " & rsFilial.Fields("Filial").Value, vbExclamation, "Atenção"
    rsFilial.Close
    Set rsFilial = Nothing
    Exit Function
  End If
  
  If GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") = 0 Then
    MsgBox "Código IBGE inválido para a UF do Cliente: " & rsSaida.Fields("Cliente").Value, vbExclamation, "Atenção"
    rsFilial.Close
    Set rsCliFor = Nothing
    Exit Function
  End If
  
  If rsFilial.Fields("CodigoRegimeTributario").Value = 1 Then
    blnSimplesNacionalGeral = True
  Else
    blnSimplesNacionalGeral = False
  End If
  
  If mostraCliente = True Then
    If rsCliFor.Fields("Física_Jurídica").Value = "J" Then PessoaJuridica = True
    
    If (Len(rsSaida.Fields("CPF_CPNJ_Cliente").Value) > 0) Then
      CPF_CnpjCliente = FormataCnpjCpf(rsSaida.Fields("CPF_CPNJ_Cliente").Value)
    ElseIf (Len(rsCliFor.Fields("CGC").Value) > 0) Then
      CPF_CnpjCliente = FormataCnpjCpf(rsCliFor.Fields("CGC").Value)
    End If
    If (Len(rsCliFor.Fields("Inscrição").Value) > 0) Then
      IE_Cliente = FormataIE(rsCliFor.Fields("Inscrição").Value, PessoaJuridica)
    End If
    If somenteCPF_CNPJ = True Then
      Set frmX = New frmDadosClienteNFCe
      If Len(CPF_CnpjCliente) > 0 Then
        gsCPF_Cnpj = CPF_CnpjCliente
      End If
      If Len(IE_Cliente) > 0 Then
        gsIE = IE_Cliente
      End If
      If Len(rsSaida.Fields("Nome_Cliente_NFCe").Value) > 0 Then
        gsNomeCliente = Left(rsSaida.Fields("Nome_Cliente_NFCe").Value, 60)
      ElseIf Len(Left(rsCliFor.Fields("Nome").Value, 60)) > 0 Then
        gsNomeCliente = Left(rsCliFor.Fields("Nome").Value, 60)
      End If
      frmX.Show vbModal
      Set frmX = Nothing
      If gsRetornoDoc <> "OK" Then
        MsgBox "Nota Fiscal ao Consumidor não impressa."
        Exit Function
      Else
        CPF_CnpjCliente = FormataCnpjCpf(gsCPF_Cnpj)
        If Len(CPF_CnpjCliente) > 11 Then
          PessoaJuridica = True
        Else
          PessoaJuridica = False
        End If
        If Len(gsIE) > 0 And IsNumeric(gsIE) Then
            If PessoaJuridica = True Then
                IE_Cliente = FormataIE(gsIE, False) 'parametros são numIE, bolPessoaFisica
            Else
                IE_Cliente = FormataIE(gsIE, True)
            End If
        End If
        strNomeCliente = gsNomeCliente
        mostraCliente = False
      End If
    End If
    '05/03/2010 - mpdea
    'Incluído validação de CPF
    '
    'Valida CNPJ
    If PessoaJuridica = True Then
      If (Trim(CPF_CnpjCliente) & "") = "" Or Not Valid_CGC(CPF_CnpjCliente & "") Then
        MsgBox "CNPJ inválido para o Cliente: " & rsSaida.Fields("Cliente").Value, vbExclamation, "Atenção"
        rsCliFor.Close
        Set rsCliFor = Nothing
        Exit Function
      End If
    Else
      If (Trim(CPF_CnpjCliente) & "") = "" Or Not Valid_CPF(CPF_CnpjCliente & "") Then
        MsgBox "CPF inválido para o Cliente: " & rsSaida.Fields("Cliente").Value, vbExclamation, "Atenção"
        rsCliFor.Close
        Set rsCliFor = Nothing
        Exit Function
      End If
    End If
  End If
  
   blnShowObs = False
'''  blnShowObs = rsOpSaida.Fields("InTelaObsTransp").Value

'''  If blnShowObs Then
'''    Set frmX = New frmObsNota
'''    frmX.gsCliente = rsCliFor.Fields("Transportadora").Value
'''    frmX.lngSequencia = Sequencia
'''    frmX.bytTipoTabela = 1
'''    frmX.Show vbModal
'''    Set frmX = Nothing
'''    If gsRetornoDoc <> "OK" Then
'''      StatusMsg "Nota Fiscal ao Consumidor não impressa."
'''      Exit Function
'''    End If
'''  Else
    For intX = 0 To 7
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
'''  End If
  
  Call IsDataType(dtLong, rsSaida.Fields("NFCe").Value, lngNumNFCe)
  If lngNumNFCe = 0 Then
    lngNumNFCe = g_longNextNFCe(rsFilial("Filial"))
    rsSaida.Edit
    rsSaida("NFCe").Value = lngNumNFCe
    rsSaida.Update
  End If
  
  If blnShowObs Then
    rsSaida.Edit
    rsSaida.Fields("obs_Placa").Value = gsPlaca
    rsSaida.Fields("obs_UF").Value = gsUfrmPlaca
    rsSaida.Fields("obs_Qtde").Value = gsQtdeTrans
    rsSaida.Fields("obs_Marca").Value = gsMarcaTrans
    rsSaida.Fields("obs_Especie").Value = gsEspecieTrans
    rsSaida.Fields("obs_PesoBruto").Value = gsPesoBruto
    rsSaida.Fields("obs_PesoLiquido").Value = gsPesoLiquido
    rsSaida.Fields("obs_Transportadora").Value = gsTransportadora
    
    rsSaida.Fields("obs_infCpl1").Value = gsObsDoc(0)
    rsSaida.Fields("obs_infCpl2").Value = gsObsDoc(1)
    
'    rsSaida.Fields("obs_Obs1").Value = gsObsDoc(0)
'    rsSaida.Fields("obs_Obs2").Value = gsObsDoc(1)
'    rsSaida.Fields("obs_Obs3").Value = gsObsDoc(2)
'    rsSaida.Fields("obs_Obs4").Value = gsObsDoc(3)
'    rsSaida.Fields("obs_Obs5").Value = gsObsDoc(4)
'    rsSaida.Fields("obs_Obs6").Value = gsObsDoc(5)
'    rsSaida.Fields("obs_Obs7").Value = gsObsDoc(6)
'    rsSaida.Fields("obs_Obs8").Value = gsObsDoc(7)
    rsSaida.Fields("obs_FretePago").Value = gsFretePago
    rsSaida.Update
  End If
  
  rsSaida.Edit
  rsSaida.Fields("DataEmissaoNota").Value = Now
  rsSaida.Update
    
  If gNrSerieNF = 0 Then
      Call IsDataType(dtInteger, rsSaida.Fields("SerieNF").Value, intSerieNFC)
      If intSerieNFC = 0 Then
        intSerieNFC = 1
        rsSaida.Edit
        rsSaida.Fields("SerieNF").Value = intSerieNFC
        rsSaida.Update
      End If
  Else
      Dim xConta As Integer
      Dim sCGC As String
      sCGC = FormataCnpjCpf(rsFilial.Fields("CGC").Value & "")
      For xConta = 0 To gNrSerieNF - 1
          If gArrayNrSerieNF(xConta, 0) = sCGC Then
              'NrCnpj1 , SerieNFe1, SerieNFCe1
              intSerieNFC = CInt(gArrayNrSerieNF(xConta, 2))
              Exit For
          End If
      Next
      
      If IsNull(intSerieNFC) Or intSerieNFC = 0 Then
        intSerieNFC = 1
      End If

      rsSaida.Edit
      rsSaida.Fields("SerieNF").Value = intSerieNFC
      rsSaida.Update
  End If

  Dim intTipoEmissao As Integer
  intTipoEmissao = GetTipoEmissao(rsFilial.Fields("Filial").Value, Sequencia, 1)
    
  strChaveNFCe = GetChaveAcesso(rsFilial.Fields("Filial").Value, Sequencia, 1)
  If strChaveNFCe <> "" Then    'Código Numérico (aleatório)
    strCodigoAleatorioChaveNFCe = Mid(strChaveNFCe, 36, 8)
    'Dígito verificador
    intDV = Right(strChaveNFCe, 1)
  
    If intTipoEmissao <> TpEmis Then 'Se a forma de emissao que está gravada na NFe for diferente da forma de emissao atual
      If TpEmis <> 1 Then 'Se for emissão em estado de contingência
        'Tem que refazer a chave de acesso, como o tipo de emissao de contingencia
        'Código Numérico (aleatório)
        Randomize
        strCodigoAleatorioChaveNFCe = Format(Int((99999999) * Rnd + 1), String(8, "0"))
        'Monta o Identificador da TAG a ser assinada
        'Exemplo (campos separados por espaço): 41 0908 79077343000104 55 001 000027025 000000001 4
        strChaveNFCe = GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") 'Código da UF
        strChaveNFCe = strChaveNFCe & Format(Date, "yyMM") 'AAMM da emissão
        strChaveNFCe = strChaveNFCe & FormataCnpjCpf(rsFilial.Fields("CGC").Value & "") 'CNPJ do Emitente
        strChaveNFCe = strChaveNFCe & "65" 'Modelo da NFCe
        strChaveNFCe = strChaveNFCe & Format(intSerieNFC, String(3, "0")) 'Serie da NFe
        strChaveNFCe = strChaveNFCe & Format(lngNumNFCe, String(9, "0")) 'Número da NFe
        strChaveNFCe = strChaveNFCe & TpEmis
        strChaveNFCe = strChaveNFCe & strCodigoAleatorioChaveNFCe 'Código Numérico (aleatório)
        'Dígito verificador
        intDV = CalculaDV(strChaveNFCe)
        strChaveNFCe = strChaveNFCe & intDV
      End If
    End If
    
  Else
    'Código Numérico (aleatório)
    Randomize
    '18/11/2010 - Andrea
    '3. Alterações layout 2.0
    'Alterado de 9 para 8 posicoes
    strCodigoAleatorioChaveNFCe = Format(Int((99999999) * Rnd + 1), String(8, "0"))
    'strCodigoAleatorioChaveNFCe = Format(Int((999999999) * Rnd + 1), String(9, "0"))
    'Monta o Identificador da TAG a ser assinada
    'Exemplo (campos separados por espaço): 41 0908 79077343000104 55 001 000027025 000000001 4
    strChaveNFCe = GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") 'Código da UF
    strChaveNFCe = strChaveNFCe & Format(Date, "yyMM") 'AAMM da emissão
    strChaveNFCe = strChaveNFCe & FormataCnpjCpf(rsFilial.Fields("CGC").Value & "") 'CNPJ do Emitente
    strChaveNFCe = strChaveNFCe & "65" 'Modelo da NFCe
    strChaveNFCe = strChaveNFCe & Format(intSerieNFC, String(3, "0")) 'Serie da NFe
    strChaveNFCe = strChaveNFCe & Format(lngNumNFCe, String(9, "0")) 'Número da NFe
    
    '18/11/2010 - Andrea
    '4. Alterações layout 2.0
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
    
    strChaveNFCe = strChaveNFCe & TpEmis
    
    strChaveNFCe = strChaveNFCe & strCodigoAleatorioChaveNFCe 'Código Numérico (aleatório)
    'Dígito verificador
    intDV = CalculaDV(strChaveNFCe)
    strChaveNFCe = strChaveNFCe & intDV
  End If
  
  '--------------------------------------------------------------------------------------------------
  ' A- Dados da Nota Fiscal Consumidor eletrônica
  '--------------------------------------------------------------------------------------------------
  
  strLinha = "<?xml version=""1.0"" encoding=""UTF-8""?>"
  strLinha = strLinha & "<NFe xmlns=""http://www.portalfiscal.inf.br/nfe"">"
  strLinha = strLinha & "<infNFe Id=""NFe" & strChaveNFCe & """ versao=""4.00"">"
  strLinha = strLinha & "<ide>"
  strLinha = strLinha & "<cUF>" & GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") & "</cUF>"
  strLinha = strLinha & "<cNF>" & strCodigoAleatorioChaveNFCe & "</cNF>"
  strLinha = strLinha & "<natOp>" & Left(rsOpSaida.Fields("Nome").Value, 60) & "</natOp>"
    
'  'Itentificação da forma de pagamento
'  If rsSaida.Fields("Total Vista").Value = rsSaida.Fields("Total").Value Then
'    strLinha = strLinha & "<indPag>0</indPag>"
'  ElseIf rsSaida.Fields("Total Prazo").Value = rsSaida.Fields("Total").Value Then
'    strLinha = strLinha & "<indPag>1</indPag>"
'  Else
'    strLinha = strLinha & "<indPag>2</indPag>"
'  End If
  
  strLinha = strLinha & "<mod>65</mod>"
  strLinha = strLinha & "<serie>" & intSerieNFC & "</serie>"
  strLinha = strLinha & "<nNF>" & lngNumNFCe & "</nNF>"
  
  'strLinha = strLinha & "<dhEmi>" & Format(rsSaida.Fields("DataEmissaoNota").Value, "yyyy-MM-dd") & "T" & Format(rsSaida.Fields("DataEmissaoNota").Value, "hh:mm:ss") & "-03:00</dhEmi>"
  strLinha = strLinha & "<dhEmi>" & Format(rsSaida.Fields("DataEmissaoNota").Value, "yyyy-MM-dd") & "T" & Format(rsSaida.Fields("DataEmissaoNota").Value, "hh:mm:ss") & sFuso & "</dhEmi>"
  
  'strLinha = strLinha & "<dhEmi>" & Format(rsSaida.Fields("DataEmissaoNota").Value, "yyyy-MM-dd") & "T" & Format(rsSaida.Fields("DataEmissaoNota").Value, "hh:mm:ss") & "-03:00</dhEmi>"
  strLinha = strLinha & "<tpNF>1</tpNF>"
  '25/01/2017 Jean
  'Alterado lógica de validação para identificar se a ID de destinatario é estadual ou interestadual baseado também no fato de mostrar ou não os dados do cliente
  If mostraCliente = False Or somenteCPF_CNPJ = True Then
    strLinha = strLinha & "<idDest>1</idDest>"
  Else
    If Len(Trim(rsCliFor.Fields("Estado").Value)) > 0 Then
      If rsFilial.Fields("Estado").Value = rsCliFor.Fields("Estado").Value Then
        strLinha = strLinha & "<idDest>1</idDest>"
      ElseIf rsFilial.Fields("Estado").Value = "EX" Then
        strLinha = strLinha & "<idDest>3</idDest>"
      Else
        strLinha = strLinha & "<idDest>2</idDest>"
      End If
    Else
      strLinha = strLinha & "<idDest>1</idDest>"
    End If
  End If
  strLinha = strLinha & "<cMunFG>" & GetTerritorioMunicipioCodigoIbge(rsFilial.Fields("Estado").Value & "", rsFilial.Fields("Cidade").Value & "") & "</cMunFG>"
  strLinha = strLinha & "<tpImp>4</tpImp>"
  strLinha = strLinha & "<tpEmis>1</tpEmis>"
  strLinha = strLinha & " <cDV>" & intDV & "</cDV>"
  strLinha = strLinha & "<tpAmb>" & rsFilial.Fields("AmbienteNFe").Value & "</tpAmb>"
  If rsSaida.Fields("FinalidadeNFe").Value > 0 Then
    strLinha = strLinha & "<finNFe>" & rsSaida.Fields("FinalidadeNFe").Value & "</finNFe>"  '& Finalidade de emissão da NF-e, 1- NF-e normal/2-NF-e complementar/3 – NF-e de ajuste NOT_FOUND
  Else
    strLinha = strLinha & "<finNFe>1</finNFe>"
  End If
  If IsNull(rsSaida.Fields("Consumidor_Final").Value) Then
    strLinha = strLinha & "<indFinal>0</indFinal>"
  Else
    strLinha = strLinha & "<indFinal>" & rsSaida.Fields("Consumidor_Final").Value & "</indFinal>"
  End If
  If IsNull(rsSaida.Fields("Presenca_Comprador").Value) Then
    strLinha = strLinha & "<indPres>1</indPres>"
  Else
    strLinha = strLinha & "<indPres>" & rsSaida.Fields("Presenca_Comprador").Value & "</indPres>"
  End If
  
  ' =====================================================================
  ' NT2020.006 v1.10 para produção em 2021
  ' Tratamento tag: indIntermed (Indicador de intermediador/marketplace)
  '                 0=Operação sem intermediador (em site ou plataforma própria)
  '                 1=Operação em site ou plataforma de terceiros (intermediadores/marketplace)
  If rsFilial.Fields("Estado").Value = "SP" Or rsFilial.Fields("Estado").Value = "MG" Or _
      rsFilial.Fields("Estado").Value = "GO" Or rsFilial.Fields("Estado").Value = "MA" Then
    
      If Not IsNull(rsSaida.Fields("Presenca_Comprador").Value) Then
          If rsSaida.Fields("Presenca_Comprador").Value = 2 Or rsSaida.Fields("Presenca_Comprador").Value = 3 Or _
          rsSaida.Fields("Presenca_Comprador").Value = 4 Or rsSaida.Fields("Presenca_Comprador").Value = 9 Then

              strLinha = strLinha & "<indIntermed>0</indIntermed>"
          'elseif rsSaidas.Fields("Presenca_Comprador").Value = 1 and procEmi = 1  Nota Fiscal Avulsa eletrônica then
              strLinha = strLinha & "<indIntermed>0</indIntermed>"
          End If
      Else
          strLinha = strLinha & "<indIntermed>0</indIntermed>"
      End If
  End If
  ' =====================================================================
  
  strLinha = strLinha & "<procEmi>0</procEmi>"
  strLinha = strLinha & "<verProc>" & rsFilial.Fields("VersaoLayoutEnvio").Value & "</verProc>"
  strLinha = strLinha & "</ide>"
  AddLinha clcArquivo, strLinha
  
  '--------------------------------------------------------------------------------------------------
  'C - Identificação do Emitente da Nota Fiscal Eletrônica
  '--------------------------------------------------------------------------------------------------
  With rsFilial
    dblPercentualPIS = (.Fields("PIS").Value / 100)
    dblPercentualCofins = (.Fields("COFINS").Value / 100)
    dblAliquotaSimples = .Fields("PercentualSimplesNacional").Value
    dblPercentualReducaoSimplesNacional = (1 - (.Fields("PercentualReducaoBCSimplesNacional").Value / 100))
    
    sCNPJ_EmitenteXML = FormataCnpjCpf(.Fields("CGC").Value & "")
    
    strLinha = "<emit>"
    strLinha = strLinha & "<CNPJ>" & FormataCnpjCpf(.Fields("CGC").Value & "") & "</CNPJ>"
    
    sTrataDados = trataCaracteresEspeciaisNFE(Left(.Fields("Razão Social").Value, 60))
    strLinha = strLinha & "<xNome>" & sTrataDados & "</xNome>"
    
    sTrataDados = trataCaracteresEspeciaisNFE(Left(.Fields("Nome").Value, 60))
    strLinha = strLinha & "<xFant>" & sTrataDados & "</xFant>"
    strLinha = strLinha & "<enderEmit>"
    
    sTrataDados = trataCaracteresEspeciaisNFE(Left(.Fields("Endereço").Value, 60))
    strLinha = strLinha & "<xLgr>" & sTrataDados & "</xLgr>"
    strLinha = strLinha & "<nro>" & Left(.Fields("EnderecoNumero").Value, 60) & "</nro>"
   
    If Left(.Fields("EnderecoComplemento").Value, 60) <> "" Then
      sTrataDados = trataCaracteresEspeciaisNFE(Left(.Fields("EnderecoComplemento").Value, 60))
      strLinha = strLinha & "<xCpl>" & sTrataDados & "</xCpl>"
    End If
   
    sTrataDados = trataCaracteresEspeciaisNFE(Left(.Fields("Bairro").Value, 60))
    strLinha = strLinha & "<xBairro>" & sTrataDados & "</xBairro>"
   
    If UCase(.Fields("Estado").Value) = "EX" Then
      strLinha = strLinha & "<cMun>" & "9999999" & "</cMun>"
      strLinha = strLinha & "<xMun>" & "EXTERIOR" & "</xMun>"
      strLinha = strLinha & "<UF>" & "EX" & "</UF>"
      str_FCP_UF_emitente = "EX"
    Else
      strLinha = strLinha & "<cMun>" & GetTerritorioMunicipioCodigoIbge(.Fields("Estado").Value, .Fields("Cidade").Value & "") & "</cMun>"
      strLinha = strLinha & "<xMun>" & Left(.Fields("Cidade").Value, 60) & "</xMun>"
      strLinha = strLinha & "<UF>" & .Fields("Estado").Value & "</UF>"
      str_FCP_UF_emitente = .Fields("Estado").Value
    End If

    strLinha = strLinha & "<CEP>" & FormataNumero(.Fields("CEP").Value & "", 8) & "</CEP>"
    If Len(.Fields("Pais").Value) > 0 Then
      strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(.Fields("Pais").Value & ""), 4) & "</cPais>"
      strLinha = strLinha & "<xPais>" & Left(.Fields("Pais").Value, 60) & "</xPais>"
    Else
      strLinha = strLinha & "<cPais>1058</cPais>"
      strLinha = strLinha & "<xPais>Brasil</xPais>"
    End If
   
    If Left(RetornaSomenteNumero(.Fields("Fone").Value), 10) <> "" Then
      strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(.Fields("Fone").Value), 10) & "</fone>"
    End If
   
    strLinha = strLinha & "</enderEmit>"
  
    strLinha = strLinha & "<IE>" & Left(FormataIE(.Fields("Inscrição").Value, False), 14) & "</IE>"
   
    If rsSaida.Fields("Produtos").Value > 0 And rsSaida.Fields("Serviços").Value > 0 Then
      strLinha = strLinha & "<IM>" & Left(.Fields("InscricaoMunicipal").Value, 15) & "</IM>"
      strLinha = strLinha & "<CNAE>" & Left(.Fields("CNAE").Value, 7) & "</CNAE>"
    End If
     
    strLinha = strLinha & "<CRT>" & rsFilial.Fields("CodigoRegimeTributario").Value & "</CRT>"
    strLinha = strLinha & "</emit>"
    AddLinha clcArquivo, strLinha
  End With
  
  '--------------------------------------------------------------------------------------------------
  'C - Identificação do Destinatario da Nota Fiscal Eletrônica
  '--------------------------------------------------------------------------------------------------
  If somenteCPF_CNPJ = True Then
    strLinha = strLinha & "<dest>"
    If PessoaJuridica = True Then
      strLinha = strLinha & "<CNPJ>" & CPF_CnpjCliente & "</CNPJ>"
    Else
      strLinha = strLinha & "<CPF>" & CPF_CnpjCliente & "</CPF>"
    End If
    If Len(strNomeCliente) > 0 Then
      strNomeCliente = trataCaracteresEspeciaisNFE(strNomeCliente)
      strLinha = strLinha & "<xNome>" & LTrim(RTrim(strNomeCliente)) & "</xNome>"
    End If
    IE = Left(IE_Cliente, 14)
    str_FCP_RG_INS_EST_destinatario = IE
    
    If Not IsNull(rsCliFor.Fields("IndicadorIE").Value) Then
        IndicadorIE = rsCliFor.Fields("IndicadorIE").Value
    End If
    
    If Not IsNull(IndicadorIE) And Trim(IndicadorIE) <> "" Then
    
      If Trim(IndicadorIE) = "0" Then
          If PessoaJuridica = True Then
              strLinha = strLinha & "<indIEDest>1</indIEDest>"
          Else
              strLinha = strLinha & "<indIEDest>9</indIEDest>"
          End If
      Else
          strLinha = strLinha & "<indIEDest>" & LTrim(RTrim(IndicadorIE)) & "</indIEDest>"
      End If
      
    Else
      If IsNumeric(IE) Then
        strLinha = strLinha & "<indIEDest>1</indIEDest>"
        IndicadorIE = "1"
      ElseIf IE = "ISENTO" Then
        strLinha = strLinha & "<indIEDest>2</indIEDest>"
        IndicadorIE = "2"
      Else
        strLinha = strLinha & "<indIEDest>9</indIEDest>"
        IndicadorIE = "9"
      End If
    End If
    
    If IndicadorIE = "1" Then
      strLinha = strLinha & "<IE>" & Left(IE_Cliente, 14) & "</IE>"
    End If
    
    strLinha = strLinha & "</dest>"
        
    'Adicionar linha
    AddLinha clcArquivo, strLinha
  
    '--------------------------------------------------------------------------------------------------
    ' TRATAMENTO PARA DESTAQUE DE ICMS PARA EMPRESAS <LUCRO REAL>....'NÃO SIMPLES'
    '--------------------------------------------------------------------------------------------------
    If blnSimplesNacionalGeral = False Then
    
        If sUFDestinatario = "" Then
            sUFDestinatario = rsFilial.Fields("Estado").Value
        End If
    
        Dim rsEstadosICMS As Recordset
        aliquotaICMS_tab_ICMS_PERC_ESTADOS = ""
  
        Set rsEstadosICMS = db.OpenRecordset("ICMS_Percentual_Estados", , dbReadOnly)
  
        If Not (rsEstadosICMS.EOF And rsEstadosICMS.BOF) Then
            rsEstadosICMS.MoveFirst
            While Not rsEstadosICMS.EOF
                If UCase(rsEstadosICMS.Fields("ESTADO_ORIGEM").Value) = UCase(gsEstadoOrigemEmpresaLogado) And _
                  UCase(rsEstadosICMS.Fields("ESTADO_DESTINO").Value) = UCase(sUFDestinatario) Then
                    aliquotaICMS_tab_ICMS_PERC_ESTADOS = rsEstadosICMS.Fields("ALIQUOTA").Value
                    rsEstadosICMS.MoveLast
                End If
                rsEstadosICMS.MoveNext
            Wend
        End If
        rsEstadosICMS.Close
        Set rsEstadosICMS = Nothing
        
        If UCase(gsEstadoOrigemEmpresaLogado) = UCase(sUFDestinatario) Then
            bo_AliquotaICMS_interestadual = False
        Else
            bo_AliquotaICMS_interestadual = True
        End If
    End If
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
  
  End If
  
  If mostraCliente = True Then
    'Validar se é para mostrar dados do cliente
    With rsCliFor
      strLinha = strLinha & "<dest>"
      If (.Fields("Física_Jurídica").Value) = "F" Then
        strLinha = strLinha & "<CPF>" & CPF_CnpjCliente & "</CPF>"
      Else
        strLinha = strLinha & "<CNPJ>" & CPF_CnpjCliente & "</CNPJ>"
      End If
      strNomeCliente = Left(.Fields("Nome").Value, 60)
      strNomeCliente = trataCaracteresEspeciaisNFE(strNomeCliente)
      strLinha = strLinha & "<xNome>" & LTrim(RTrim(strNomeCliente)) & "</xNome>"
      
      strLinha = strLinha & "<enderDest>"
      Dim sTrataDado As String
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Endereço").Value, 60))
      strLinha = strLinha & "<xLgr>" & LTrim(RTrim(sTrataDado)) & "</xLgr>"
      strLinha = strLinha & "<nro>" & Left(.Fields("Endereço Número").Value, 60) & "</nro>"
      
      If Len(.Fields("Complemento").Value) <> 0 Then
        sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Complemento").Value, 60))
        strLinha = strLinha & "<xCpl>" & LTrim(RTrim(sTrataDado)) & "</xCpl>"
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
        sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Cidade").Value, 60))
        strLinha = strLinha & "<xMun>" & sTrataDado & "</xMun>"
        strLinha = strLinha & "<UF>" & .Fields("Estado").Value & "</UF>"
        str_FCP_UF_destinatario = .Fields("Estado").Value
      End If
      
      strLinha = strLinha & "<CEP>" & FormataNumero(.Fields("CEP").Value & "", 8) & "</CEP>"
      If Len(.Fields("Pais").Value) > 0 Then
        strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(.Fields("Pais").Value & ""), 4) & "</cPais>"
        strLinha = strLinha & "<xPais>" & Left(.Fields("Pais").Value, 60) & "</xPais>"
      Else
        strLinha = strLinha & "<cPais>1058</cPais>"
        strLinha = strLinha & "<xPais>Brasil</xPais>"
      End If
      
      If Left(RetornaSomenteNumero(.Fields("Fone 1").Value), 10) <> "" Then
        strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(.Fields("Fone 1").Value), 10) & "</fone>"
      End If
      
      strLinha = strLinha & "</enderDest>"
      
      'Dim IndicadorIE As String
      'Dim IE As String
      IE = Left(IE_Cliente, 14)
      
      If Not IsNull(rsCliFor.Fields("IndicadorIE").Value) Then
          IndicadorIE = rsCliFor.Fields("IndicadorIE").Value
      End If
      
      If Not IsNull(IndicadorIE) And IndicadorIE <> "" Then
        strLinha = strLinha & "<indIEDest>" & IndicadorIE & "</indIEDest>"
      Else
        If IsNumeric(IE) Then
          strLinha = strLinha & "<indIEDest>1</indIEDest>"
          IndicadorIE = "1"
        ElseIf IE = "ISENTO" Then
          strLinha = strLinha & "<indIEDest>2</indIEDest>"
          IndicadorIE = "2"
        Else
          strLinha = strLinha & "<indIEDest>9</indIEDest>"
          IndicadorIE = "9"
        End If
      End If
      
'''      If IsNumeric(IE) Then
'''        strLinha = strLinha & "<indIEDest>1</indIEDest>"
'''        IndicadorIE = "1"
'''      ElseIf IE = "ISENTO" Then
'''        strLinha = strLinha & "<indIEDest>2</indIEDest>"
'''        IndicadorIE = "2"
'''      Else
'''        strLinha = strLinha & "<indIEDest>9</indIEDest>"
'''        IndicadorIE = "9"
'''      End If
    '      If IsNull(.Fields("IndicadorIE").Value) Then
    '        strLinha = strLinha & "9;"
    '      Else
    '        strLinha = strLinha & .Fields("IndicadorIE").Value & ";"
    '      End If
      If IndicadorIE = "1" Then
        strLinha = strLinha & "<IE>" & Left(IE_Cliente, 14) & "</IE>"
      End If
     
      'strLinha = strLinha & "        <IE>" & Left(FormataIE(.Fields("Inscrição").Value, UCase(.Fields("Física_Jurídica").Value) = "F"), 14) & "</IE>"
      
      If Len(.Fields("email").Value) <> 0 Then
        sTrataDado = trataCaracteresEspeciaisNFE(.Fields("email").Value)
        strLinha = strLinha & "<email>" & sTrataDado & "</email>"
      End If
    
      strLinha = strLinha & "</dest>"
        
      'Adicionar linha
      AddLinha clcArquivo, strLinha
    End With
  End If
  '--------------------------------------------------------------------------------------------------
  'I - Produtos e Serviços da NF-e
  '--------------------------------------------------------------------------------------------------
  isDescontoItem = False
  
''  If (Vdesc <= 0) Then
''    If rsSaida.Fields("Desconto").Value <= 0 Then
''      Vdesc = Round(rsSaida.Fields("DescontoSubTotal").Value, 2) 'alteracao desconto 11-05-2017
''       bDescontoNaVendaRapida_por_valorOUperc = True
''    Else
''      Vdesc = Round(rsSaida.Fields("Desconto").Value, 2) 'alteracao desconto 11-05-2017
''    End If
''
''    isDescontoItem = True
''  End If
  
  '************************
  'Tratamento para desconto na venda
  Dim vColuna_Desconto As Double          ' Coluna da tabela de Saídas
  Dim vColuna_DescontoSubTotal As Double  ' Coluna da tabela de Saídas
  Dim vAuxDesc As Double
  Dim vAuxTotalSaidaProdutos As Double
  bTemDescontoNaVenda = False
  
  vAuxTotalSaida = rsSaida.Fields("Produtos").Value
  vColuna_DescontoSubTotal = rsSaida.Fields("DescontoSubTotal").Value
  vColuna_Desconto = rsSaida.Fields("Desconto").Value

  If vColuna_DescontoSubTotal > 0 Then
      bTemDescontoNaVenda = True
      vTemDescontoNaVendaPercentual = (vColuna_DescontoSubTotal * 100) / vAuxTotalSaida
      
  ElseIf vColuna_Desconto > 0 Then
      Dim rsSaidasProdAux As Recordset
      strSQL = "SELECT SUM([Preço Final]) From [Saídas - Produtos] "
      strSQL = strSQL & "WHERE Filial = " & gnCodFilial & " AND Sequência = " & Sequencia
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

  'bTemDescontoNaVenda = False   'se FALSE, Desliga a rotina de tratamento
  '************************
  
   
  iDetNumItemXML = 0
  'AddProdutoSevicosSaida rsFilial.Fields("Filial").Value, Sequencia, clcArquivo, rsSaida, rsFilial, dblPercentualPIS, dblPercentualCofins, dblAliquotaSimples, dblPercentualReducaoSimplesNacional, rsOpSaida.Fields("CSO").Value, Vdesc
  AddProdutoSevicosSaida rsFilial.Fields("Filial").Value, Sequencia, clcArquivo, rsSaida, rsFilial, dblPercentualPIS, dblPercentualCofins, dblAliquotaSimples, dblPercentualReducaoSimplesNacional, rsOpSaida.Fields("CSO").Value, vSOMA_TOTAL_DESCONTO
  '--------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------
  'W - Valores Totais da NF-e
  '--------------------------------------------------------------------------------------------------
  Dim str_volumes_quantidade As String
  strLinha = "<total>"
  strLinha = strLinha & "<ICMSTot>"
  
  If (rsFilial.Fields("CodigoRegimeTributario").Value) <> 1 Then
  
    '===================================================
    ' Regime tributário LUCRO REAL
  
    If vlSOMA_DO_ICMS > 0 And v_vlSOMA_DO_ICMS_BaseDeCalculo > 0 Then
        strLinha = strLinha & "<vBC>" & FormataValorTexto(v_vlSOMA_DO_ICMS_BaseDeCalculo, 2) & "</vBC>"
        strLinha = strLinha & " <vICMS>" & FormataValorTexto(vlSOMA_DO_ICMS, 2) & "</vICMS>"
    Else
        strLinha = strLinha & "<vBC>0.00</vBC>"
        strLinha = strLinha & "<vICMS>0.00</vICMS>"
    End If
    
    strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(rsSaida.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
    
    ' FCP tratamento NFe 4.0    Pilatti/2018
    ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
    If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
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

    strLinha = strLinha & "<vBCST>" & FormataValorTexto(rsSaida.Fields("Base ICM Subs").Value, 2) & "</vBCST>"
    strLinha = strLinha & "<vST>" & FormataValorTexto(rsSaida.Fields("Valor ICM Subs").Value, 2) & "</vST>"
  Else
    strLinha = strLinha & "<vBC>" & "0.00" & "</vBC>"
    strLinha = strLinha & "<vICMS>" & "0.00" & "</vICMS>"
    strLinha = strLinha & "<vICMSDeson>" & "0.00" & "</vICMSDeson>"
    strLinha = strLinha & "<vFCP>0.00</vFCP>"
    strLinha = strLinha & "<vBCST>" & "0.00" & "</vBCST>"
    strLinha = strLinha & "<vST>" & "0.00" & "</vST>"
  End If
  
  ' FCP tratamento NFe 4.0    Pilatti/2018
  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
  If blnSimplesNacionalGeral = False Then
    If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
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
  
  'strLinha = strLinha & "<vProd>" & FormataValorTexto(rsSaida.Fields("Produtos").Value + rsSaida.Fields("Serviços").Value, 2) & "</vProd>"
  strLinha = strLinha & "<vProd>" & FormataValorTexto(total_produto + rsSaida.Fields("Serviços").Value, 2) & "</vProd>"
  strLinha = strLinha & "<vFrete>" & FormataValorTexto(rsSaida.Fields("Frete").Value, 2) & "</vFrete>"
  strLinha = strLinha & "<vSeg>" & FormataValorTexto(rsSaida.Fields("Seguro").Value, 2) & "</vSeg>"
  
''  If m_bln_tem_desconto = True Or Vdesc > 0 Then
''    If rsSaida.Fields("Desconto").Value <= 0 Then
''      strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsSaida.Fields("DescontoSubTotal").Value, 2) & "</vDesc>"
''    Else
''      strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsSaida.Fields("Desconto").Value, 2) & "</vDesc>"
''    End If
''  Else
''    strLinha = strLinha & "<vDesc>" & "0.00" & "</vDesc>"
''  End If
''  m_bln_tem_desconto = False
      If m_bln_tem_desconto = True Then
        
        If bTemDescontoNaVenda = True Then 'Tratamento para DESCONTO DADO DIRETO NA VENDA (Não desconto por produto)
          strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
        Else
          strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsSaida.Fields("Desconto").Value, 2) & "</vDesc>"
        End If
      
      Else
        If bTemDescontoNaVenda = True Then
            strLinha = strLinha & "<vDesc>" & FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2) & "</vDesc>"
        Else
            strLinha = strLinha & "<vDesc>0.00</vDesc>"
        End If
      End If
      m_bln_tem_desconto = False
  
  strLinha = strLinha & "<vII>0.00</vII>"
  '''strLinha = strLinha & "<vIPI>" & FormataValorTexto(rsSaida.Fields("IPI").Value, 2) & "</vIPI>"
  strLinha = strLinha & "<vIPI>0.00</vIPI>"
  
  ' NFe 4.0 Pilatti 2018
  'vIPIDevol
  strLinha = strLinha & "<vIPIDevol>0.00</vIPIDevol>"

'  strLinha = strLinha & "<vPIS>" & FormataValorTexto((rsSaida.Fields("Total").Value * dblPercentualPIS), 2) & "</vPIS>"
'  strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((rsSaida.Fields("Total").Value * dblPercentualPIS), 2) & "</vCOFINS>"
  If inst_dblPercentualPIS > 0 Then
      strLinha = strLinha & "<vPIS>" & FormataValorTexto(inst_dblPercentualPIS, 2) & "</vPIS>"
      strLinha = strLinha & "<vCOFINS>" & FormataValorTexto(inst_dblPercentualCofins, 2) & "</vCOFINS>"
  Else
      strLinha = strLinha & "<vPIS>0.00</vPIS>"
      strLinha = strLinha & "<vCOFINS>0.00</vCOFINS>"
  End If
  
  strLinha = strLinha & "<vOutro>" & "0.00" & "</vOutro>"
  strLinha = strLinha & "<vNF>" & FormataValorTexto(rsSaida.Fields("Total").Value, 2) & "</vNF>"
  If IsNull(rsSaida.Fields("TotalNCM").Value) Or TotalAproxImpostos = 0 Then
  Else
    strLinha = strLinha & "<vTotTrib>" & FormataValorTexto(TotalAproxImpostos, 2) & "</vTotTrib>"
  End If
  strLinha = strLinha & "</ICMSTot>"
  strLinha = strLinha & "</total>"
  AddLinha clcArquivo, strLinha
  strSQL = ""
  
  '--------------------------------------------------------------------------------------------------
  'X-Informações do Transporte da NF-e
  '--------------------------------------------------------------------------------------------------
  strSQL = "SELECT [obs_FretePago] FROM [Saídas] WHERE Filial = " & gnCodFilial & " AND Sequência = " & Sequencia
  Set rsTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
'  If rsTransportadora.EOF Then
    strLinha = "<transp>"
    strLinha = strLinha & "<modFrete>" & "9" & "</modFrete>"
    strLinha = strLinha & "</transp>"
    AddLinha clcArquivo, strLinha
'  Else
'    strLinha = "<transp>"
'
'
'    '16ªlinha [5000;0]
'    ' Modalidade do frete, 0 – por conta do emitente; 1 – por conta do destinatário
'    If rsTransportadora.Fields("obs_FretePago").Value = 1 Then
'      strLinha = strLinha & "<modFrete>" & "0" & "</modFrete>"
'    ElseIf rsTransportadora.Fields("obs_FretePago").Value = 2 Then
'      strLinha = strLinha & "<modFrete>" & "1" & "</modFrete>"
'    Else
'      strLinha = strLinha & "<modFrete>" & "9" & "</modFrete>"
'    End If
'
'    AddLinha clcArquivo, strLinha
'
'    strSQL = ""
'    rsTransportadora.Close
'
'    Set rsTransportadora = Nothing
'
'    'Transporta
'    strSQL = "SELECT [Saídas].obs_Transportadora, Transportadoras.CGC, Transportadoras.Nome, "
'    strSQL = strSQL & "Transportadoras.[inscrição], Transportadoras.[Endereço], Transportadoras.Cidade, "
'    strSQL = strSQL & "Transportadoras.Estado FROM [Saídas], Transportadoras WHERE [Saídas].Filial = " & gnCodFilial & "  AND "
'    strSQL = strSQL & "[Saídas].Sequência = " & Sequencia & " AND Transportadoras.Nome = [Saídas].obs_Transportadora"
'
'    Set rsTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
'
'    If Not (rsTransportadora.BOF And rsTransportadora.EOF) Then
'
'      strLinha = strLinha & "<transporta>"
'      If (rsTransportadora.Fields("CGC").Value <> "") Then
'      strLinha = strLinha & "<CNPJ>" & RetornaSomenteNumero(rsTransportadora.Fields("CGC").Value & "") & "</CNPJ>"
'      End If
'      If (rsTransportadora.Fields("Nome").Value <> "") Then
'      strLinha = strLinha & "<xNome>" & rsTransportadora.Fields("Nome").Value & "</xNome>"
'      End If
'      If (rsTransportadora.Fields("inscrição").Value <> "") Then
'      strLinha = strLinha & "<IE>" & RetornaSomenteNumero(rsTransportadora.Fields("inscrição").Value & "") & "</IE>"
'      End If
'      If (rsTransportadora.Fields("Endereço").Value <> "") Then
'      strLinha = strLinha & "<xEnder>" & rsTransportadora.Fields("Endereço").Value & "</xEnder>"
'      End If
'      If (rsTransportadora.Fields("Cidade").Value <> "") Then
'      strLinha = strLinha & "<xMun>" & rsTransportadora.Fields("Cidade").Value & "</xMun>" & vbCrLf
'      End If
'      If (rsTransportadora.Fields("Estado").Value <> "") Then
'      strLinha = strLinha & "<UF>" & rsTransportadora.Fields("Estado").Value & "</UF>" & vbCrLf
'      End If
'      strLinha = strLinha & "</transporta>"
'      AddLinha clcArquivo, strLinha
'
'      If Not IsNull(rsSaida.Fields("obs_Placa").Value) Then
'        If rsSaida.Fields("obs_Placa").Value <> "" Then
'          strLinha = "<veicTransp>" & vbCrLf
'          strLinha = strLinha & "<placa>" & rsSaida.Fields("obs_Placa").Value & "</placa>" & vbCrLf
'          strLinha = strLinha & "<UF>" & rsSaida.Fields("obs_UF").Value & "</UF>" & vbCrLf
'          strLinha = strLinha & "</veicTransp>"
'          AddLinha clcArquivo, strLinha
'        End If
'      End If
'
'      strLinha = "<vol>" & vbCrLf
'
'      If IsNull(rsSaida.Fields("obs_Qtde").Value) Then
'        str_volumes_quantidade = "0"
'      Else
'        str_volumes_quantidade = rsSaida.Fields("obs_Qtde").Value
'      End If
'
'      str_volumes_quantidade = Format(str_volumes_quantidade, String(15, "0"))
'
'      If (str_volumes_quantidade <> "") Then
'      strLinha = strLinha & "<qVol>" & str_volumes_quantidade & "</qVol>" & vbCrLf
'      End If
'
'      If (rsSaida.Fields("obs_Especie").Value <> "") Then
'        strLinha = strLinha & "<esp>" & rsSaida.Fields("obs_Especie").Value & "</esp>" & vbCrLf
'      End If
'
'      If (rsSaida.Fields("obs_Marca").Value <> "") Then
'        strLinha = strLinha & "<marca>" & rsSaida.Fields("obs_Marca").Value & "</marca>" & vbCrLf
'      End If
'
'      'strLinha = strLinha & "          <nVol>" & "" & "</nVol>" & vbCrLf 'Campo Numero de Volumes não encontrado
'
'      Call IsDataType(dtDouble, rsSaida.Fields("obs_PesoLiquido").Value, dblPeso)
'      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
'      strLinha = strLinha & "<pesoL>" & FormataValorTexto(dblPeso, 3) & "</pesoL>" & vbCrLf
'      End If
'      Call IsDataType(dtDouble, rsSaida.Fields("obs_PesoBruto").Value, dblPeso)
'      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
'      strLinha = strLinha & "<pesoB>" & FormataValorTexto(dblPeso, 3) & "</pesoB>" & vbCrLf
'      End If
'      strLinha = strLinha & "</vol>"
'      AddLinha clcArquivo, strLinha
'
'    Else
'      MsgBox "Transportadora não encontrada para a Filial " & gnCodFilial & ", sequência " & Sequencia, vbExclamation, "Atenção"
'      AddLinha clcArquivo, strLinha
'    End If
'    strSQL = ""
'    rsTransportadora.Close
'
'    Set rsTransportadora = Nothing
'
'    strLinha = strLinha & "</transp>"
'  End If
'  AddLinha clcArquivo, strLinha
  '--------------------------------------------------------------------------------------------------
  'YA-Informações de pagamento 1
  '--------------------------------------------------------------------------------------------------
  
  AddLinha clcArquivo, strLinha
  
  strLinha = "<pag>"
  
  If rsSaida.Fields("Recebe - Dinheiro").Value > 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>01</tPag>"
    
    If rsSaida.Fields("Recebe - Dinheiro") > rsSaida.Fields("Total") Then
      rsSaida.Fields("Recebe - Dinheiro") = rsSaida.Fields("Total")
    End If
    
    strLinha = strLinha & "<vPag>" & FormataValorTexto((rsSaida.Fields("Recebe - Dinheiro").Value), 2) & "</vPag>"
    strLinha = strLinha & "</detPag>"
  End If
  
  
  Dim rsCheques As Recordset
  '''Dim totalCheque As Integer
  Dim totalCheque As Double
  
  totalCheque = 0
  Set rsCheques = db.OpenRecordset("Select * from [Movimento - Cheques] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
  Do Until rsCheques.EOF
    totalCheque = totalCheque + rsCheques.Fields("Valor")
    rsCheques.MoveNext
  Loop
  If totalCheque > 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>02</tPag>"
    If totalCheque > rsSaida.Fields("Total") Then
      totalCheque = rsSaida.Fields("Total")
    End If
    strLinha = strLinha & "<vPag>" & FormataValorTexto(totalCheque, 2) & "</vPag>"
    strLinha = strLinha & "</detPag>"
  End If
  
  If rsSaida.Fields("TotalCartaoCredito").Value <> 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>1</indPag>" '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>03</tPag>"
    
    Dim tmpTotalCartaoCredito As Double
    tmpTotalCartaoCredito = rsSaida.Fields("TotalCartaoCredito").Value
    If tmpTotalCartaoCredito > rsSaida.Fields("Total") Then
       tmpTotalCartaoCredito = rsSaida.Fields("Total").Value
    End If
    strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoCredito, 2) & "</vPag>"
    strLinha = strLinha & "<card><tpIntegra>2</tpIntegra></card>"
    strLinha = strLinha & "</detPag>"
  End If
  
  If rsSaida.Fields("TotalCartaoDebito").Value <> 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>0</indPag>" '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>04</tPag>"
    
    Dim tmpTotalCartaoDebito As Double
    tmpTotalCartaoDebito = rsSaida.Fields("TotalCartaoDebito").Value
    If tmpTotalCartaoDebito > rsSaida.Fields("Total") Then
        tmpTotalCartaoDebito = rsSaida.Fields("Total")
    End If
    strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoDebito, 2) & "</vPag>"
    strLinha = strLinha & "<card><tpIntegra>2</tpIntegra></card>"
    strLinha = strLinha & "</detPag>"
  End If
  
  Dim TotalOutros As Double
  If Not IsNull(rsSaida.Fields("Total Prazo").Value) Then
    TotalOutros = rsSaida.Fields("Total Prazo").Value
  End If
  
  If Not IsNull(rsSaida.Fields("Recebe - Vale").Value) Then
    TotalOutros = TotalOutros + rsSaida.Fields("Recebe - Vale").Value
  End If

  If TotalOutros > 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<tPag>99</tPag>"
    If TotalOutros > rsSaida.Fields("Total") Then TotalOutros = rsSaida.Fields("Total")
    strLinha = strLinha & "<vPag>" & FormataValorTexto(TotalOutros, 2) & "</vPag>"
    strLinha = strLinha & "</detPag>"
  End If
  
  strLinha = strLinha & "</pag>"
  AddLinha clcArquivo, strLinha
  
  '--------------------------------------------------------------------------------------------------
  'Z-Informações adicionais
  '--------------------------------------------------------------------------------------------------
  Dim str_observacoes_NF As String
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & "Val Aprox dos Tributos R$"
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaida.Fields("TotalNCM").Value & "("
  Dim aliquotaNCM As Double
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then aliquotaNCM = (Format(rsSaida("TotalNCM"), "##,###,##0.00") / Format(rsSaida("Total"), "##,###,##0.00")) * 100
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then aliquotaNCM = Format(aliquotaNCM, "##,###,##0.00")
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & aliquotaNCM & "%) Fonte IBPT."
  
  If Not IsNull(rsSaida.Fields("obs_infCpl1").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaida.Fields("obs_infCpl1").Value
  If Not IsNull(rsSaida.Fields("obs_infCpl2").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaida.Fields("obs_infCpl2").Value
  
  str_observacoes_NF = Trim(str_observacoes_NF)
  '---------------------------------------------------------
  '07/04/2011 - Andrea
  'Incluída mensagem de tributacao, no primeiro campo depois da tag 7000
  'Carrega as mensagens para Nota Fiscal
  Dim str_observacoes_NF_fisco As String
  Set m_clcMensagens = New Collection
  Call GetMensagemNotaFiscal(gnCodFilial, Sequencia, m_clcMensagens)
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
  strLinha = "<infAdic>"
  If (str_observacoes_NF_fisco <> "") Then
    strLinha = strLinha & "<infAdFisco>" & str_observacoes_NF_fisco & "</infAdFisco>"
  End If
  strLinha = strLinha & "<infCpl>" & str_observacoes_NF & "</infCpl>"
  strLinha = strLinha & "</infAdic>"
  strLinha = strLinha & "</infNFe>"
  strLinha = strLinha & "</NFe>"
  'strLinha = strLinha & "</enviNFe>"
  AddLinha clcArquivo, strLinha
  
  If Not DirExists(".\NFCe\") Then
    MkDir (".\NFCe\")
  End If
  
  GravarArquivo clcArquivo, ".\NFCe\" & "NFCeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_" & lngNumNFCe & ".xml"
  Dim xml As Variant
  xml = ""
  'continuar
  Dim i As Integer
  Dim cont As Integer
  If Left(rsFilial.Fields("EnderecoComplemento").Value, 60) <> "" Then
    Endereco_Emitente = Left(rsFilial.Fields("Endereço").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoNumero").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoComplemento").Value, 60) & ", " & Left(rsFilial.Fields("Bairro").Value, 60) & ", " & Left(rsFilial.Fields("Cidade").Value, 60)
  Else
    Endereco_Emitente = Left(rsFilial.Fields("Endereço").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoNumero").Value, 60) & ", " & Left(rsFilial.Fields("Bairro").Value, 60) & ", " & Left(rsFilial.Fields("Cidade").Value, 60)
  End If
  cont = clcArquivo.Count
  i = 1
  
  
  '************************************************************************************
  'TRATAMENTO PARA CUPOM WS - CHAMADA BENEFIX
On Error GoTo SegueFluxoWS
  Dim iRet_EnviarXML_SEFAZ As Integer
  '''Dim iIndice01 As Integer
  '''Dim iIndice02 As Integer
  Dim sUrlQrCode As String
  Dim sEmContingencia As String
  Dim iTipo As Integer

  sRetFazenda = ""
  iRet_EnviarXML_SEFAZ = 0
  'Chamada WS (Web Service)
  iRet_EnviarXML_SEFAZ = EnviarXML_SEFAZ(Sequencia, gnCodFilial, sCNPJ_EmitenteXML)
SegueFluxoWS:

  GravarArquivo clcArquivo, strPastaEnvio & "NFCeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_SeqVenda" & Sequencia & ".xml"

'''  MsgBox iRet_EnviarXML_SEFAZ
  
  If iRet_EnviarXML_SEFAZ = 1 Then
      Dim rsTesteConexao As Recordset
    On Error GoTo SegueFluxo
      Set rsTesteConexao = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = ''")
SegueFluxo:

      If Err.Number <> 0 Then
          Set BancoPDV = OpenDatabase(rsFilial("BancoPDV").Value & "\QuickStore.mdb", False, False, ";PWD=" & gsGetPValue())
      End If

      Dim rsNFCe_Job As Recordset
      Set rsNFCe_Job = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = 'XXXXXXX' And cnpj = '999999'")

      'Pegando Serie do retorno da fazenda/benefix
      iIndice01 = InStr(1, sRetFazenda, "<serie>")
      If iIndice01 > 0 Then
          iIndice02 = InStr(1, sRetFazenda, "</serie>")
          intSerieNFC = Mid$(sRetFazenda, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
      End If

      'Pegando NumeroNota do retorno da fazenda/benefix
      iIndice01 = InStr(1, sRetFazenda, "<numero>")
      If iIndice01 > 0 Then
          iIndice02 = InStr(1, sRetFazenda, "</numero>")
          lngNumNFCe = Mid$(sRetFazenda, iIndice01 + 8, iIndice02 - (iIndice01 + 8))
      End If

      'Pegando urlQrCode do retorno da fazenda/benefix
      iIndice01 = InStr(1, sRetFazenda, "<urlQrCode>")
      If iIndice01 > 0 Then
          iIndice02 = InStr(1, sRetFazenda, "</urlQrCode>")
          sUrlQrCode = Mid$(sRetFazenda, iIndice01 + 11, iIndice02 - (iIndice01 + 11))
      End If

      'Pegando Situação da Emissão (normal ou contingencia) do retorno da fazenda/benefix
      '<emContingencia>false</emContingencia> ou <emContingencia>true</emContingencia>
      iIndice01 = InStr(1, sRetFazenda, "<emContingencia>")
      If iIndice01 > 0 Then
          iIndice02 = InStr(1, sRetFazenda, "</emContingencia>")
          sEmContingencia = Mid$(sRetFazenda, iIndice01 + 16, iIndice02 - (iIndice01 + 16))
      End If

      If sEmContingencia = "false" Then
          iTipo = 1     ' Não esta em contingencia
      Else
          iTipo = 2     ' Esta em contingencia
      End If


      If rsNFCe_Job.EOF Then
        rsNFCe_Job.AddNew
        rsNFCe_Job!CNPJ = FormataCnpjCpf(rsFilial.Fields("CGC").Value)
        Do Until i > cont
          xml = xml + Replace(CStr(clcArquivo.Item(i)), vbCrLf, "")
          i = i + 1
        Loop
        rsNFCe_Job!xml = xml
        '''rsNFCe_Job!Tipo = 1
        rsNFCe_Job!Tipo = iTipo
        rsNFCe_Job!Serie = intSerieNFC
        rsNFCe_Job!N_NF = lngNumNFCe
        rsNFCe_Job!Chave = strChaveNFCe
        rsNFCe_Job!CPF = CPF_CnpjCliente
        rsNFCe_Job!Nome_Consumidor = strNomeCliente
        rsNFCe_Job!Data_Emissao = rsSaida.Fields("DataEmissaoNota").Value
        rsNFCe_Job!Total_Tributos = FormataValorTexto(rsSaida.Fields("TotalNCM").Value, 2)
        rsNFCe_Job!Nome_Emitente = Left(rsFilial.Fields("Razão Social").Value, 60)
        rsNFCe_Job!Endereco_Emitente = Endereco_Emitente
        rsNFCe_Job!IE_Emitente = FormataIE(rsFilial.Fields("Inscrição").Value, False)
        rsNFCe_Job!retFazenda = sRetFazenda
        '''rsNFCe_Job!retFazenda = sUrlQrCode
        rsNFCe_Job.Update
      End If
      
      rsNFCe_Job.Close
      Set rsNFCe_Job = Nothing

  End If
      ' Fim tratamento CUPOM WS
      '************************************************************************************

  'DisplayMsg Xml
  rsSaida.Edit
  rsSaida.Fields("ChaveNFCe") = strChaveNFCe
  rsSaida.Fields("CPF_CPNJ_Cliente") = CPF_CnpjCliente
  rsSaida.Fields("Emitiu_Dados_Cliente_NFCe") = mostrouCliente
  rsSaida.Fields("Emitiu_Somente_CPF_Cliente_NFCe") = mostrouSomenteCpf_CNPJ
  rsSaida.Fields("Nome_Cliente_NFCe") = strNomeCliente
  
  If iRet_EnviarXML_SEFAZ = 0 Then 'Algo inconsistente/erro/rejeição fazenda
      rsSaida.Fields("RetNFCe") = gStringRetorno
  End If
  rsSaida.Update
  
  If iRet_EnviarXML_SEFAZ = 1 Then
      'MsgBox ("Arquivo gerado com sucesso")
      sRetornoEnvioNFCe = "Cupom Fiscal enviado com sucesso. Agora vai para a impressora."
  Else
      sRetornoEnvioNFCe = ""
  End If
  
  Exit Function

TrataErro:
  If Err.Number = 3050 Then  'Couldn't lock file
    MsgBox "Erro de comunicação com a sua máquina local...Verificar se o CUPOM FISCAL foi enviado para a SEFAZ.. Erro: " & Err.Description, vbCritical, "Erro"
  Else
    MsgBox "Erro genérico...Verificar se o CUPOM FISCAL foi enviado para a SEFAZ. Erro: " & Err.Description, vbCritical, "Erro"
  End If
  
End Function

Private Function AddProdutoSevicosSaida(ByVal Filial As String, ByVal Sequencia As Long, ByRef clcArquivo As Collection, ByRef rsSaida As Recordset, ByRef rsFilial As Recordset, PercentualPIS As Double, PercentualCOFINS As Double, ByVal dblAliquotaSimples As Double, ByVal dblPercentualReducaoBC_SN As Double, ByVal CSO_NaturezaOperacao As String, ByVal Vdesc As Double)
  Dim strSQL As String
  Dim strLinha As String
  Dim strTributacao As String
  Dim rsSaidaProdutos As Recordset
  Dim rsSaidaProdutos_CFOP As Recordset
  Dim dbl_percentual_frete_rateado As Double
  Dim dbl_percentual_seguro_rateado As Double
  Dim dbl_valor_frete_rateado As Double
  Dim dbl_valor_seguro_rateado As Double
  Dim rsSaidasAUXILIAR_ICMS As Recordset
  
  Dim ValorBaseCalculoICMS As Double
  Dim AliquotaICMS_Saida As Double
  Dim AliquotaICMSST_Saida As Double
  Dim ValorBCICMSST_Saida As Double
  Dim PercentualIPI_Saida As Double
  
  
  Dim s_vlSOMA_DO_DESCONTO As String
  Dim vlSOMA_DO_DESCONTO As Double
  vlSOMA_DO_DESCONTO = 0
  s_vlSOMA_DO_DESCONTO = ""
  
  Dim iContaRS As Integer
  Dim iContaLoop As Integer
  Dim s_vlSOMA_DO_ICMS As String
  'vlSOMA_DO_ICMS = 0
  's_vlSOMA_DO_ICMS = ""
  iContaRS = 0
  iContaLoop = 0
  
  Dim iDescontoAuxZ_2 As Integer
  Dim sDescontoAuxZ_2 As String
  
  'Valores Totais da NF-e
  strSQL = "SELECT [Base ICM], [Valor ICM], [Base ICM Subs], [Valor ICM Subs], Produtos, [Serviços], Frete, Seguro, [Recebe - Dinheiro],  "
  strSQL = strSQL & "Desconto, IPI, [Total PIS], [Total COFINS], Total, Frete, Seguro, obs_Qtde, obs_Especie, obs_Marca, obs_Placa, obs_UF,"
  strSQL = strSQL & "obs_PesoLiquido, obs_PesoBruto, obs_infCpl1, obs_infCpl2, TotalNCM, TotalDesoneracaoICMS "
  strSQL = strSQL & " ,FreteSomaOuNaoEstimativa, TotalCartaoDebito, TotalCartaoCredito, [Total Prazo], [Recebe - Vale] "
  strSQL = strSQL & " FROM [Saídas] WHERE Filial = " & Filial
  strSQL = strSQL & "AND Sequência = " & Sequencia
  Set rsSaidasAUXILIAR_ICMS = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  
  'Add Produtos
  strSQL = "SELECT Produtos.Código, Produtos.Nome, Produtos.CodigoNBM, Produtos.AliqNCM, [Saídas - Produtos].CFOP,Produtos.[Nome Nota], "
  strSQL = strSQL & "Produtos.[Tipo ICM], Produtos.[Base Cálculo], Produtos.[Redução ICM],  "
  strSQL = strSQL & "Produtos.[Unidade Venda], Produtos.[Situação Tributária] As [Situação Tributária Produto], "
  strSQL = strSQL & "[Saídas - Produtos].Qtde , [Saídas - Produtos].Preço, "
  strSQL = strSQL & "[Saídas - Produtos].[Preço Final],[Saídas - Produtos].[Desconto Valor], "
  strSQL = strSQL & "Saídas.Produtos, Saídas.Frete, Saídas.Seguro, Saídas.Desconto, [Saídas - Produtos].[Descricao Adicional], "
  strSQL = strSQL & "[Saídas - Produtos].[Situação Tributária], [Saídas - Produtos].ICM, [Saídas - Produtos].[Código sem Grade],  "
  strSQL = strSQL & "[Saídas - Produtos].IPI, Produtos.[IPI_Reduzido], [Saídas - Produtos].[Linha], [Saídas - Produtos].[BaseCalculoICMS], "
  strSQL = strSQL & "[Saídas - Produtos].Valor_Aprox_Impostos, [Saídas - Produtos].MotivoDesoneracaoICMS,"
  strSQL = strSQL & "[Saídas - Produtos].ValorICMSDesonerado, [Saídas - Produtos].Percentual_Diferimento, "
  strSQL = strSQL & "[Saídas - Produtos].Desconto as DescontoNoProduto, Saídas.DescontoSubTotal, Saídas.Operação, "
  strSQL = strSQL & "Produtos.TipoSituacaoTributariaPIS, Produtos.CodigoBeneficio, "
  strSQL = strSQL & "Produtos.[Percentual Icm Saida], Produtos.CodigoBeneficio, Produtos.BaseCalculoICMSST_Saida, "
  strSQL = strSQL & "Produtos.Percentual_ICMSST_Saida, Produtos.[Percentual IPI], "
  strSQL = strSQL & "Produtos.Tipo, [Saídas - Produtos].Código as CodigoComGrade "
  
''  strSQL = strSQL & " FROM [Saídas - Produtos], Produtos, Saídas WHERE "
''  strSQL = strSQL & "[Saídas - Produtos].Filial = " & Filial & " AND [Saídas - Produtos].Sequência = " & Sequencia
''  strSQL = strSQL & " AND Produtos.Código = [Saídas - Produtos].[Código sem Grade] And Saídas.Sequência = " & Sequencia
''  strSQL = strSQL & " AND Saídas.Filial = " & Filial
  strSQL = strSQL & " FROM [Saídas - Produtos], Produtos, Saídas WHERE "
  strSQL = strSQL & " [Saídas - Produtos].Filial = " & Filial & " AND [Saídas - Produtos].Sequência = " & Sequencia
  strSQL = strSQL & " AND [Saídas - Produtos].[Código sem Grade] = Produtos.Código And Saídas.Filial = " & Filial & " And Saídas.Sequência = " & Sequencia
  strSQL = strSQL & " ORDER BY [Saídas - Produtos].[Linha] "
  
  '20/08/2010 - Andrea
  'Alterado o relacionamento da tabela Saídas Produto com a tabela Produtos
  'porque os produtos com grade não estava sendo encontrados
  'strSQL = strSQL & " And Produtos.Código = [Saídas - Produtos].Código And Saídas.Sequência = " & Sequencia
 
  Set rsSaidaProdutos = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  
  Dim BaseCalculoICMS As Double
  Dim AliquotaICMS As Double
  Dim ValorICMS As Double
  Dim BaseCalculoICMS_ST As Double
  Dim AliquotaICMS_ST As Double
  Dim ValorICMS_ST As Double
  Dim MargemValorAdicionado_ICMS As Double
  Dim PercentualReducaoBaseICMS As Double
  Dim nomeProduto As String
  Dim Linha As String
  Dim SituacaoTributaria As String
  
  '21/01/2010 - mpdea
  Dim dbl_total_produtos As Double
  
  '30/03/2011 - Andrea
  Dim str_CSO_Produto As String
  
  '01/04/2011 - Andrea
  Dim str_Desconto_Produto As String
  Dim dbl_Desconto_Produto As Double
  
  '14/02/2016 - Andrea
  Dim dbl_resto As Double
  Dim int_itens As Integer
  Dim dbl_desconto_acumulado As Double
  Dim int_item_atual As Integer
  Dim str_resto As String
  Dim str_desconto_acumulado As String
  Dim str_total_produtos As String
  
    ' *****************************************************
  'Inicio FCP (cabecalho calculo) Pilatti

'  Dim AliquotaAdesaoUFDestFCP As Double
'  Dim AliquotaFCP_EMIT_DEST As Double
'  Dim AliquotaFCP_Emitente As Double
'  Dim AliquotaFCP_Destinatario As Double
'  Dim CalculoFCPEstadosDiferentes As Double
'  Dim CalculoFCPEstadosIguais As Double
'  Dim CalculoFCP_diferenca As Double
'  Dim CalculoFCP_participacaoDest As Double
'  Dim CalculoFCP_participacaoEmitente As Double
'
'  Dim iAnoFCP As Integer
'  Dim vBC_fcp As Double
'  Dim vFCP_registroTotal As Double
'  Dim vFCPST_registroTotal As Double
'  Dim vFCPSTRet_registroTotal As Double
'  Dim vResultado As Double


  Dim ds_auxiliarSQL As Recordset
  
  vFCP_registroTotal = 0
  vFCPST_registroTotal = 0
  vFCPSTRet_registroTotal = 0
  'iAnoFCP = 2018
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
  
  
  dbl_desconto_acumulado = 0
  
  With rsSaidaProdutos
  
    'Total de produtos para rateio do frete e do seguro
    dbl_total_produtos = 0
    int_itens = 0
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        '31/03/2010 - Andrea
        'Trocado o campo preco final pelo calculo preco X qtde porque o preco final vem com ipi
        'dbl_total_produtos = dbl_total_produtos + CDbl(rsSaidaProdutos.Fields("Preço Final").Value)
        dbl_total_produtos = dbl_total_produtos + CDbl((.Fields("Qtde").Value * .Fields("Preço").Value))
        iContaRS = iContaRS + 1
        dbl_total_produtos = Round(dbl_total_produtos, 2)
        
        int_itens = int_itens + 1
        
        .MoveNext
      Loop
      .MoveFirst
    End If
    
    Dim str_CFOP As String
    
    Do Until .EOF
      If Not IsNull(.Fields("TipoSituacaoTributariaPIS").Value) Then
          localTipoSituacaoTributariaPIS = .Fields("TipoSituacaoTributariaPIS").Value
      Else
          'Caso não tenha um valor por produto....então recebe o valor configurado para a FILIAL
          localTipoSituacaoTributariaPIS = gTipoSituacaoTributariaPIS
      End If
      
      'Seleciona a Situacao Tributaria do Produto
      SituacaoTributaria = .Fields("Situação Tributária").Value
      If IsNull(SituacaoTributaria) Or Len(Trim(SituacaoTributaria)) <= 0 Then
        If IsNull(.Fields("Situação Tributária Produto").Value) Then
          SituacaoTributaria = "000"
        ElseIf Trim(.Fields("Situação Tributária Produto").Value) = "" Then
          SituacaoTributaria = "000"
        Else
          SituacaoTributaria = .Fields("Situação Tributária Produto").Value
        End If
      End If
      
      Dim strValorItem As String
      
      strValorItem = Round(.Fields("Preço").Value, 4)
      
      'Saidas Produtos CFOP
'      strSQL = "SELECT ProdutoCFOP.CFOP, ProdutoCFOP.CSO  "
'      strSQL = strSQL & " FROM ProdutoCFOP, [Saídas - Produtos],  Saídas WHERE "
'      strSQL = strSQL & "[Saídas - Produtos].Filial = " & Filial & " AND [Saídas - Produtos].Sequência = " & Sequencia
'      strSQL = strSQL & " And ProdutoCFOP.CodProduto = '" & rsSaidaProdutos.Fields("Código").Value & "' And Saídas.Sequência = " & Sequencia
'      strSQL = strSQL & " And ProdutoCFOP.CodOperacao =  Saídas.Operação "

      strSQL = "SELECT CFOP, CSO  "
      strSQL = strSQL & " FROM ProdutoCFOP "
      strSQL = strSQL & " WHERE CodProduto = '" & rsSaidaProdutos.Fields("Código").Value & "'"
      strSQL = strSQL & " And CodOperacao = " & .Fields("Operação").Value

      Set rsSaidaProdutos_CFOP = db.OpenRecordset(strSQL, dbOpenDynaset)
      
      '21/01/2010 - mpdea
      'Corrigido percentual de rateio do item
      '
      'Rateio do Frete por item
      If rsSaida.Fields("Frete").Value > 0 And dbl_total_produtos > 0 Then
        dbl_percentual_frete_rateado = ((rsSaidaProdutos.Fields("Preço Final").Value * 100) / dbl_total_produtos)
        If dbl_percentual_frete_rateado > 0 Then
          dbl_valor_frete_rateado = ((rsSaida.Fields("Frete").Value * dbl_percentual_frete_rateado) / 100)
        End If
      End If

      '21/01/2010 - mpdea
      'Corrigido percentual de rateio do item
      '
      'Rateio do valor do Seguro por item
      If rsSaida.Fields("Seguro").Value > 0 And dbl_total_produtos > 0 Then
        dbl_percentual_seguro_rateado = ((rsSaidaProdutos.Fields("Preço Final").Value * 100) / dbl_total_produtos)
        If dbl_percentual_seguro_rateado > 0 Then
          dbl_valor_frete_rateado = ((rsSaida.Fields("Seguro").Value * dbl_percentual_seguro_rateado) / 100)
        End If
      End If
      
      
      'descontro proporcional
      If (Vdesc > 0) Then
        dbl_Desconto_Produto = ((.Fields("Qtde").Value * .Fields("Preço").Value) * Vdesc) / dbl_total_produtos
        dbl_Desconto_Produto = Round(dbl_Desconto_Produto, 2)
        dbl_desconto_acumulado = dbl_desconto_acumulado + dbl_Desconto_Produto
        dbl_desconto_acumulado = Round(dbl_desconto_acumulado, 2)
        
        If bDescontoNaVendaRapida_por_valorOUperc = False Then
            If (isDescontoItem) Then
              dbl_Desconto_Produto = Round((.Fields("Qtde").Value * .Fields("Preço").Value), 2) - Round(.Fields("Preço Final").Value, 2)
            Else
              If .Fields("DescontoNoProduto").Value > 0 Then
                dbl_Desconto_Produto = Round((.Fields("Qtde").Value * .Fields("Preço").Value), 2) - Round(.Fields("Preço Final").Value, 2)
              Else
                dbl_Desconto_Produto = dbl_Desconto_Produto + Round((.Fields("Qtde").Value * .Fields("Preço").Value), 2) - Round(.Fields("Preço Final").Value, 2)
              End If
            End If
            dbl_Desconto_Produto = Round(dbl_Desconto_Produto, 2)
        Else
            dbl_Desconto_Produto = Round(dbl_Desconto_Produto, 2)
        End If
      End If
      int_item_atual = int_item_atual + 1
      
      If int_item_atual = int_itens Then
        dbl_resto = Round(Vdesc, 2) - Round(dbl_desconto_acumulado, 2)
        dbl_resto = Round(dbl_resto, 2)
        If dbl_resto > 0 Then
          dbl_Desconto_Produto = dbl_Desconto_Produto + dbl_resto
          dbl_Desconto_Produto = Round(dbl_Desconto_Produto, 2)
        End If
      End If
      
      If Len(rsSaidaProdutos.Fields("Nome Nota").Value) > 0 Then
        nomeProduto = .Fields("Nome Nota").Value
      Else
        nomeProduto = .Fields("Nome").Value
      End If
      nomeProduto = Trim(nomeProduto)
      
      Linha = .Fields("Linha").Value
      
      iDetNumItemXML = iDetNumItemXML + 1
      'strLinha = "<det nItem=""" & .Fields("Linha").Value & """>"
      strLinha = "<det nItem=""" & iDetNumItemXML & """>"
      strLinha = strLinha & "<prod>"
      strLinha = strLinha & "<cProd>" & .Fields("Código").Value & "</cProd>"
      strLinha = strLinha & "<cEAN>SEM GTIN</cEAN>"
      
      nomeProduto = trataCaracteresEspeciaisNFE(nomeProduto)
      
      If .Fields("Tipo").Value = "G" Then
          strLinha = strLinha & "<xProd>" & nomeProduto & " " & Right(.Fields("CodigoComGrade").Value, 6) & "</xProd>"
      Else
          strLinha = strLinha & "<xProd>" & nomeProduto & "</xProd>"
      End If
     
      '''strLinha = strLinha & "<xProd>" & nomeProduto & "</xProd>"
      strLinha = strLinha & "<NCM>" & .Fields("CodigoNBM").Value & "</NCM>"
      
      'Incluir código CEST
      Dim strCest As String
      strCest = ""
      Dim strSQLCest As String
      Dim rsCest As Recordset
      strSQLCest = "select cest from Ref_CEST_NCM where ncm = '" & .Fields("CodigoNBM").Value & "';"
      Set rsCest = db.OpenRecordset(strSQLCest, dbOpenDynaset)
      If Not (rsCest.BOF And rsCest.EOF) Then
        Do Until rsCest.EOF
          strCest = CStr(rsCest.Fields("cest").Value)
          strLinha = strLinha & "<CEST>" & CStr(rsCest.Fields("cest").Value) & "</CEST>"
          rsCest.MoveNext
        Loop
        rsCest.MoveFirst
      End If
      rsCest.Close
      Set rsCest = Nothing
      
      
      ' Inicio tratamento cBenef (para empresa LUCRO REAL)
      If blnSimplesNacionalGeral = False Then
          If Right(.Fields("Situação Tributária").Value, 2) = "40" Then
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
      
      'Verifica se o CFOP do produto X operacao está zerado
      'se estiver, pegao o CFOP que vem da natureza de operacao e está gravado nos itens da venda
      If Not rsSaidaProdutos_CFOP.EOF Then
        If rsSaidaProdutos_CFOP.Fields("CFOP").Value <> 0 Then
          str_CFOP = Format(rsSaidaProdutos_CFOP.Fields("CFOP").Value, String(4, "0")) 'CFOP do cadastro de produto X CFOP
        Else
          str_CFOP = Format(.Fields("CFOP").Value, String(4, "0"))  'CFOP dos itens da venda
        End If
        
        If rsSaidaProdutos_CFOP.Fields("CSO").Value <> 0 Then
          str_CSO_Produto = Format(rsSaidaProdutos_CFOP.Fields("CSO").Value, String(3, "0")) 'CSO do produto X Natureza de operacao
        ElseIf Len(SituacaoTributaria) > 3 Then
          str_CSO_Produto = Mid(SituacaoTributaria, 1, 3)
        Else
          str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
        End If
      Else
          str_CFOP = Format(.Fields("CFOP").Value, String(4, "0"))  'CFOP do Produto
          If Len(SituacaoTributaria) > 3 Then
            str_CSO_Produto = Mid(SituacaoTributaria, 1, 3)
          End If
          str_CSO_Produto = Format(CSO_NaturezaOperacao, String(3, "0")) 'CSO da Natureza de Operacao
      End If
      
      If IsEmpty(Trim(str_CFOP)) Or IsNull(Trim(str_CFOP)) Or Trim(str_CFOP) = "" Then
          str_CFOP = CFOP_Generico
      End If
      
      Dim Unidade_Venda As String
      If IsNull(Trim(.Fields("Unidade Venda").Value)) Then
          Unidade_Venda = "UN"
      Else
        If IsEmpty(Trim(.Fields("Unidade Venda").Value)) Or Trim(.Fields("Unidade Venda").Value) = "" Then
          Unidade_Venda = "UN"
        Else
          Unidade_Venda = Trim(.Fields("Unidade Venda").Value)
        End If
      End If

      strLinha = strLinha & "<CFOP>" & str_CFOP & "</CFOP>"
      strLinha = strLinha & "<uCom>" & Unidade_Venda & "</uCom>"
      strLinha = strLinha & "<qCom>" & FormataValorTexto(.Fields("Qtde").Value, 4) & "</qCom>"
      strLinha = strLinha & "<vUnCom>" & FormataValorTexto(strValorItem, 10) & "</vUnCom>"
      
      Dim vProd As Double
      'Dim vDescItem As Double
      
      
      vProd = (.Fields("Qtde").Value * .Fields("Preço").Value)
      vProd = Round(vProd, 2)
      
      total_produto = total_produto + vProd
      
      'Michel 18/03/2016
      'strLinha = strLinha & "<vProd>" & FormataValorTexto((.Fields("Qtde").Value * .Fields("Preço").Value), 2) & "</vProd>" & vbCrLf
      strLinha = strLinha & "<vProd>" & FormataValorTexto(vProd, 2) & "</vProd>"
      strLinha = strLinha & "<cEANTrib>SEM GTIN</cEANTrib>"
      strLinha = strLinha & "<uTrib>" & Unidade_Venda & "</uTrib>"
      strLinha = strLinha & "<qTrib>" & FormataValorTexto(.Fields("Qtde").Value, 4) & "</qTrib>"
      strLinha = strLinha & "<vUnTrib>" & FormataValorTexto(strValorItem, 10) & "</vUnTrib>"
      
      If dbl_valor_frete_rateado <> 0 Then
        strLinha = strLinha & "<vFrete>" & FormataValorTexto(dbl_valor_frete_rateado, 2) & "</vFrete>"
      End If
      If dbl_valor_seguro_rateado <> 0 Then
        strLinha = strLinha & "<vSeg>" & FormataValorTexto(dbl_valor_seguro_rateado, 2) & "</vSeg>"
      End If
      
      'If ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) <= 0.009 Then
        'strLinha = strLinha & ";"
      'Else
      
' PILATTI INICIO ----------------------------------------------------------------------------
' DATA: 2017/06/06
' Comentamos a condição IF abaixo

'      If dbl_Desconto_Produto <> 0 Then
'        m_bln_tem_desconto = True
'        'strLinha = strLinha & "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value), 2) & "</vDesc>" & vbCrLf
'        strLinha = strLinha & "<vDesc>" & FormataValorTexto(dbl_Desconto_Produto, 2) & "</vDesc>" & vbCrLf
'      'Else
'        'vDescItem = FormataValorTexto(rsSaidaProdutos.Fields("Desconto").Value, 2)
'
'        'strLinha = strLinha & "<vDesc>" & FormataValorTexto(vDescItem, 2) & "</vDesc>" & vbCrLf
'      End If
      
'''      If dbl_Desconto_Produto <> 0 Then
'''        ' PILATTI MAIO/2018
'''        dbl_Desconto_ProdutoACUMULADO = dbl_Desconto_ProdutoACUMULADO + dbl_Desconto_Produto
'''
'''        m_bln_tem_desconto = True
'''
'''        Dim descTotal As Double
'''        Dim descAcum As Double
'''        Dim dbl_Desconto_ProdutoAJUSTE As Double
'''
'''        If bDescontoNaVendaRapida_por_valorOUperc = True Then
'''          descTotal = .Fields("DescontoSubTotal").Value
'''        Else
'''          descTotal = .Fields("Desconto").Value
'''        End If
'''        descAcum = dbl_desconto_acumulado - 0.01
'''
'''        If dbl_Desconto_ProdutoACUMULADO > descTotal Then
'''          dbl_Desconto_ProdutoAJUSTE = dbl_Desconto_ProdutoACUMULADO - descTotal
'''          dbl_Desconto_Produto = dbl_Desconto_Produto - dbl_Desconto_ProdutoAJUSTE
'''
'''          If dbl_Desconto_Produto <= 0 Then
'''            dbl_Desconto_Produto = 0
'''          Else
'''            If FormataValorTexto(dbl_Desconto_Produto, 2) <> "0.00" Then
'''              strLinha = strLinha & "<vDesc>" & FormataValorTexto(dbl_Desconto_Produto, 2) & "</vDesc>"
'''            End If
'''          End If
'''        Else
'''          'If (.Fields("Desconto").Value = dbl_desconto_acumulado - 0.01) Then
'''          If CStr(descTotal) = CStr(descAcum) Then
'''            If FormataValorTexto(dbl_Desconto_Produto - 0.01, 2) <> "0.00" Then
'''              strLinha = strLinha & "<vDesc>" & FormataValorTexto(dbl_Desconto_Produto - 0.01, 2) & "</vDesc>"
'''            End If
'''          Else
'''            If FormataValorTexto(dbl_Desconto_Produto, 2) <> "0.00" Then
'''              strLinha = strLinha & "<vDesc>" & FormataValorTexto(dbl_Desconto_Produto, 2) & "</vDesc>"
'''            End If
'''          End If
'''        End If
'''
'''      End If
' FIM --------------------------------------------------------------------------------
      Dim vAuxiliarCalculoVendaDesconto As Double
      Dim sDescontoAuxZ As String
      If ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) <= 0.009 Then
        'strLinha = strLinha & ";"
        
        If bTemDescontoNaVenda = True Then
            vAuxiliarCalculoVendaDesconto = ((.Fields("Qtde").Value * .Fields("Preço").Value) * vTemDescontoNaVendaPercentual) / 100
            s_vlSOMA_DO_DESCONTO = FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2)
            s_vlSOMA_DO_DESCONTO = Replace(s_vlSOMA_DO_DESCONTO, ".", ",")
            vlSOMA_DO_DESCONTO = vlSOMA_DO_DESCONTO + CDbl(s_vlSOMA_DO_DESCONTO)
            
            If iContaRS = 1 Then
                sDescontoAuxZ = "<vDesc>" & FormataValorTexto(Vdesc, 2) & "</vDesc>"
                
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
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
                    'strLinha = strLinha & "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                ElseIf vlSOMA_DO_DESCONTO < v_vSOMA_TOTAL_DESCONTO Then
                    v_diferenca_DESCONTO = v_vSOMA_TOTAL_DESCONTO - vlSOMA_DO_DESCONTO
                    'strLinha = strLinha & "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto + v_diferenca_DESCONTO, 2) & "</vDesc>"
                    sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto + v_diferenca_DESCONTO, 2) & "</vDesc>"
                ElseIf vlSOMA_DO_DESCONTO > v_vSOMA_TOTAL_DESCONTO Then
                    Dim sAux1 As String
                    sAux1 = FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2)
                    sAux1 = Replace(sAux1, ".", ",")
                    vAuxiliarCalculoVendaDesconto = CDbl(sAux1)
                    
                    v_diferenca_DESCONTO = vlSOMA_DO_DESCONTO - v_vSOMA_TOTAL_DESCONTO
                    'strLinha = strLinha & "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto - v_diferenca_DESCONTO, 2) & "</vDesc>"
                    sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto - v_diferenca_DESCONTO, 2) & "</vDesc>"
                End If
                
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                  strLinha = strLinha & sDescontoAuxZ
                End If
            Else
                'strLinha = strLinha & "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                '''sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                sDescontoAuxZ = "<vDesc>" & FormataValorTexto(s_vlSOMA_DO_DESCONTO, 2) & "</vDesc>"
            
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                  strLinha = strLinha & sDescontoAuxZ
                End If
            
            End If
        End If
      Else
        m_bln_tem_desconto = True
        
        If bTemDescontoNaVenda = True Then
            vAuxiliarCalculoVendaDesconto = ((.Fields("Qtde").Value * .Fields("Preço").Value) * vTemDescontoNaVendaPercentual) / 100
            s_vlSOMA_DO_DESCONTO = FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2)
            s_vlSOMA_DO_DESCONTO = Replace(s_vlSOMA_DO_DESCONTO, ".", ",")
            vlSOMA_DO_DESCONTO = vlSOMA_DO_DESCONTO + CDbl(s_vlSOMA_DO_DESCONTO)
            
            If iContaRS = 1 Then
                sDescontoAuxZ = "<vDesc>" & FormataValorTexto(Vdesc, 2) & "</vDesc>"
                
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                  strLinha = strLinha & sDescontoAuxZ
                End If
            ElseIf iContaRS = iContaLoop + 1 Then
                'Dim s_vSOMA_TOTAL_DESCONTO As String
                'Dim v_vSOMA_TOTAL_DESCONTO As Double
                'Dim v_diferenca_DESCONTO As Double
                s_vSOMA_TOTAL_DESCONTO = FormataValorTexto(vSOMA_TOTAL_DESCONTO, 2)
                s_vSOMA_TOTAL_DESCONTO = Replace(s_vSOMA_TOTAL_DESCONTO, ".", ",")
                v_vSOMA_TOTAL_DESCONTO = CDbl(s_vSOMA_TOTAL_DESCONTO)
                    
                If vlSOMA_DO_DESCONTO = v_vSOMA_TOTAL_DESCONTO Then
                    sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                    'strLinha = strLinha & "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                ElseIf vlSOMA_DO_DESCONTO < v_vSOMA_TOTAL_DESCONTO Then
                    v_diferenca_DESCONTO = v_vSOMA_TOTAL_DESCONTO - vlSOMA_DO_DESCONTO
                    'strLinha = strLinha & "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto + v_diferenca_DESCONTO, 2) & "</vDesc>"
                    sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto + v_diferenca_DESCONTO, 2) & "</vDesc>"
                ElseIf vlSOMA_DO_DESCONTO > v_vSOMA_TOTAL_DESCONTO Then
                    v_diferenca_DESCONTO = vlSOMA_DO_DESCONTO - v_vSOMA_TOTAL_DESCONTO
                    'strLinha = strLinha & "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto - v_diferenca_DESCONTO, 2) & "</vDesc>"
                    sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto - v_diferenca_DESCONTO, 2) & "</vDesc>"
                End If
                
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                  strLinha = strLinha & sDescontoAuxZ
                End If
            Else
                'strLinha = strLinha & "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                'sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value) + vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
                sDescontoAuxZ = "<vDesc>" & FormataValorTexto(vAuxiliarCalculoVendaDesconto, 2) & "</vDesc>"
            
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
                  strLinha = strLinha & sDescontoAuxZ
                End If
            End If
            
        Else
            'strLinha = strLinha & "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value), 2) & "</vDesc>"
            sDescontoAuxZ = "<vDesc>" & FormataValorTexto(((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value), 2) & "</vDesc>"
        
            If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" Then
              strLinha = strLinha & sDescontoAuxZ
            End If
        End If
      End If
    
      
      If (.Fields("Desconto").Value > 0) Then
        m_bln_tem_desconto = True
      End If

      strLinha = strLinha & "<indTot>" & "1" & "</indTot>"
            
      strLinha = strLinha & "</prod>"
      AddLinha clcArquivo, strLinha
      
      '---------------------------------------------------------------------
      ValorBaseCalculoICMS = .Fields("Qtde").Value * .Fields("Preço").Value
      
      If Not IsNull(.Fields("Percentual Icm Saida").Value) Then
        AliquotaICMS_Saida = .Fields("Percentual Icm Saida").Value
      Else
        AliquotaICMS_Saida = 0
      End If
      
      If Not IsNull(.Fields("Percentual_ICMSST_Saida").Value) Then
        AliquotaICMSST_Saida = .Fields("Percentual_ICMSST_Saida").Value
      Else
        AliquotaICMSST_Saida = 0
      End If
      
      If Not IsNull(.Fields("BaseCalculoICMSST_Saida").Value) Then
        ValorBCICMSST_Saida = .Fields("BaseCalculoICMSST_Saida").Value
      Else
        ValorBCICMSST_Saida = 0
      End If
      
      If Not IsNull(.Fields("Percentual IPI").Value) Then
        PercentualIPI_Saida = .Fields("Percentual IPI").Value
      Else
        PercentualIPI_Saida = 0
      End If
      
      
'''      If rsSaidaProdutos.Fields("Tipo ICM").Value = "N" Then
'''        'ICMS Normal
'''        '31/03/2010 - Andrea
'''        'Trocado o campo preco final pelo calculo preco X qtde porque o preco final vem com ipi
'''        BaseCalculoICMS = (rsSaidaProdutos.Fields("Qtde").Value * rsSaidaProdutos.Fields("Preço").Value)
'''        AliquotaICMS = rsSaidaProdutos.Fields("ICM").Value
'''        ValorICMS = (.Fields("Qtde").Value * .Fields("Preço").Value) * CSng(gsHandleNull(rsSaidaProdutos.Fields("ICM").Value & "")) / 100
'''
'''      ElseIf rsSaidaProdutos.Fields("Tipo ICM").Value = "R" Then
'''        'ICMS Retido
'''        If CDbl(rsSaidaProdutos.Fields("Base Cálculo").Value) <> 0 Then
'''          'Base Fixa
'''          BaseCalculoICMS_ST = CDbl(rsSaidaProdutos.Fields("Qtde").Value) * CDbl(rsSaidaProdutos.Fields("Base Cálculo").Value)
'''          AliquotaICMS_ST = rsSaidaProdutos.Fields("ICM").Value
'''          ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsSaidaProdutos.Fields("ICM").Value & "")) / 100
'''        ElseIf CDbl(rsSaidaProdutos.Fields("Redução ICM").Value) <> 0 Then
'''          'Base Reduzida
'''          '31/03/2010 - Andrea
'''          'Trocado o campo preco final pelo calculo preco X qtde porque o preco final vem com ipi
'''          'BaseCalculoICMS_ST = CDbl(rsSaidaProdutos.Fields("Preço Final").Value) * CDbl(rsSaidaProdutos.Fields("Redução ICM").Value) / 100
'''          BaseCalculoICMS_ST = CDbl(.Fields("Qtde").Value * .Fields("Preço").Value) * CDbl(rsSaidaProdutos.Fields("Redução ICM").Value) / 100
'''          AliquotaICMS_ST = rsSaidaProdutos.Fields("ICM").Value
'''          ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsSaidaProdutos.Fields("ICM").Value & "")) / 100
'''          PercentualReducaoBaseICMS = rsSaidaProdutos.Fields("Redução ICM").Value
'''        End If
'''      ElseIf rsSaidaProdutos.Fields("Tipo ICM").Value = "Z" Then
'''        'Tratamento do
'''        'ICMS Reduzido
'''        If CDbl(rsSaidaProdutos.Fields("Base Cálculo").Value) <> 0 Then
'''          'Base Fixa
'''          BaseCalculoICMS_ST = CDbl(rsSaidaProdutos.Fields("Qtde").Value) * CDbl(rsSaidaProdutos.Fields("Base Cálculo").Value)
'''          AliquotaICMS_ST = rsSaidaProdutos.Fields("ICM").Value
'''          ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsSaidaProdutos.Fields("ICM").Value & "")) / 100
'''        ElseIf CDbl(rsSaidaProdutos.Fields("Redução ICM").Value) <> 0 Then
'''          'Base Reduzida
'''          '31/03/2010 - Andrea
'''          'Trocado o campo preco final pelo calculo preco X qtde porque o preco final vem com ipi
'''          'BaseCalculoICMS_ST = CDbl(rsSaidaProdutos.Fields("Preço Final").Value) * CDbl(rsSaidaProdutos.Fields("Redução ICM").Value) / 100
'''          BaseCalculoICMS_ST = CDbl(.Fields("Qtde").Value * .Fields("Preço").Value) * CDbl(rsSaidaProdutos.Fields("Redução ICM").Value) / 100
'''          AliquotaICMS_ST = rsSaidaProdutos.Fields("ICM").Value
'''          ValorICMS_ST = BaseCalculoICMS_ST * CSng(gsHandleNull(rsSaidaProdutos.Fields("ICM").Value & "")) / 100
'''          PercentualReducaoBaseICMS = rsSaidaProdutos.Fields("Redução ICM").Value
'''        End If
'''      End If
      
      ' AQUI CONTA LOOP
      iContaLoop = iContaLoop + 1
      
      
      strTributacao = Right(SituacaoTributaria, 2)
      
      strLinha = "<imposto>"
      AddLinha clcArquivo, strLinha
      If IsNull(.Fields("Valor_Aprox_Impostos").Value) Or .Fields("Valor_Aprox_Impostos").Value = 0 Then
        'strLinha = "<vTotTrib>0.00</vTotTrib>"
      Else
        strLinha = "<vTotTrib>" & FormataValorTexto(.Fields("Valor_Aprox_Impostos").Value, 2) & "</vTotTrib>"
        TotalAproxImpostos = TotalAproxImpostos + .Fields("Valor_Aprox_Impostos")
      End If
      strLinha = strLinha & "<ICMS>"
      AddLinha clcArquivo, strLinha
      
      Dim dblBaseCalculoSimplesItem As Double
      
      If (rsFilial.Fields("CodigoRegimeTributario").Value) <> 1 Then
      
        sDescontoAuxZ_2 = "0"
      
        Select Case Right(SituacaoTributaria, 2)
          Case "00"
              'Tributos Incidentes no Produto ou no serviço
              strLinha = strLinha & "<ICMS00>"
              strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
              strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
              strLinha = strLinha & "<modBC>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"

              If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                  iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                  sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                  sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                  strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                  '''s_vlSOMA_DO_ICMS_BaseDeCalculo = BaseCalculoICMS - CDbl(sDescontoAuxZ_2)
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2)
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                  v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
              Else
                  strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS, 2) & "</vBC>"
                  '''s_vlSOMA_DO_ICMS_BaseDeCalculo = BaseCalculoICMS
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = FormataValorTexto(ValorBaseCalculoICMS, 2)
                  s_vlSOMA_DO_ICMS_BaseDeCalculo = Replace(s_vlSOMA_DO_ICMS_BaseDeCalculo, ".", ",")
                  v_vlSOMA_DO_ICMS_BaseDeCalculo = v_vlSOMA_DO_ICMS_BaseDeCalculo + CDbl(s_vlSOMA_DO_ICMS_BaseDeCalculo)
              End If
              
              If bo_AliquotaICMS_interestadual = False Then
                  'VENDA ESTADUAL
                  strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS_Saida, 2) & "</pICMS>"
                  ValorICMS = (.Fields("Qtde").Value * .Fields("Preço").Value) * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
              Else
                  'VENDA INTERESTADUAL
                  strLinha = strLinha & "<pICMS>" & FormataValorTexto(aliquotaICMS_tab_ICMS_PERC_ESTADOS, 2) & "</pICMS>"
                  AliquotaICMS = aliquotaICMS_tab_ICMS_PERC_ESTADOS
                  ValorICMS = (.Fields("Qtde").Value * .Fields("Preço").Value) * CSng(gsHandleNull(AliquotaICMS_Saida & "")) / 100
              End If
              
              If AliquotaICMS > 0 Then
                  bTEM_ICMS_3179a = True
              End If
              
              If m_bln_tem_desconto = True Or bTemDescontoNaVenda = True Then
                  s_vlSOMA_DO_ICMS = ((ValorBaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * AliquotaICMS_Saida / 100)
                  s_vlSOMA_DO_ICMS = FormataValorTexto(s_vlSOMA_DO_ICMS, 2)
                  s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                  vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                  
                  If iContaRS = iContaLoop Then
                      Dim s_total_icms_campo As String
                      Dim v_total_icms_campo As Double
                      Dim v_diferenca As Double
                      
                      s_total_icms_campo = s_vlSOMA_DO_ICMS
                      v_total_icms_campo = CDbl(s_total_icms_campo)
                      strLinha = strLinha & "<vICMS>" & FormataValorTexto(s_vlSOMA_DO_ICMS, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
                  Else
                      'strLinha = strLinha & "<vICMS>" & FormataValorTexto((BaseCalculoICMS - ((.Fields("Qtde").Value * .Fields("Preço").Value) - .Fields("Preço Final").Value)) * AliquotaICMS / 100, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
                      strLinha = strLinha & "<vICMS>" & FormataValorTexto(s_vlSOMA_DO_ICMS, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
                  End If
              Else
                  s_vlSOMA_DO_ICMS = FormataValorTexto(ValorICMS, 2)
                  s_vlSOMA_DO_ICMS = Replace(s_vlSOMA_DO_ICMS, ".", ",")
                  vlSOMA_DO_ICMS = vlSOMA_DO_ICMS + CDbl(s_vlSOMA_DO_ICMS)
                  strLinha = strLinha & "<vICMS>" & FormataValorTexto(s_vlSOMA_DO_ICMS, 2) & "</vICMS>" 'Valor do ICMS confirmar campo
              End If
              
            
'''              ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
'''              ' IDEM [3101;pFCP;vFCP]
'''              If str_FCP_UF_destinatario <> "" And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
'''                  vBC_fcp = BaseCalculoICMS
'''
'''                  vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
'''                  strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 4) & "</pFCP>"
'''                  strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
'''                  vFCP_registroTotal = vFCP_registroTotal + vResultado
'''              End If
              
              strLinha = strLinha & "</ICMS00>"
              AddLinha clcArquivo, strLinha
            
         Case "10"
           'ICMS Normal e ST
           strLinha = strLinha & "<ICMS10>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
           strLinha = strLinha & "<modBC>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
           strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS, 2) & "</vBC>"
           strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
           strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
           
           ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
           If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
              Dim vAuxiliar As Double
              vAuxiliar = BaseCalculoICMS_ST
    
              ' COLOCAR GRUPO 3111           ou           3112
              If vAuxiliar = 0 Then
                  vBC_fcp = ValorBaseCalculoICMS
                  vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                  ' [3111;vBCFCP;pFCP;vFCP]
                  strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                  strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCP>"
                  strLinha = strLinha & "<vBC>" & FormataValorTexto(vResultado, 2) & "</vBC>"
                  vFCP_registroTotal = vFCP_registroTotal + vResultado
              End If
           End If
           
           strLinha = strLinha & "<modBCST>" & rsFilial.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
           'strLinha = strLinha & "              <pMVAST>" & FormataValorTexto(MargemValorAdicionadoICMS, 2) & "</pMVAST>"
           strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
           strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
           strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 2) & "</pICMSST>"
           strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
           
           ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
           If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
              'Dim vAuxiliar As Double
              vAuxiliar = BaseCalculoICMS_ST
    
              ' COLOCAR GRUPO 3111           ou           3112
              If vAuxiliar <> 0 Then
                  vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                  ' [3112;vBCFCPST;pFCPST;vFCPST]
                  strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                  strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPST>"
                  strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                  vFCPST_registroTotal = vFCPST_registroTotal + vResultado
              End If
           End If
           
           strLinha = strLinha & "</ICMS10>"
           AddLinha clcArquivo, strLinha

         Case "20"
           'Dados do ICMS da operação
           strLinha = strLinha & " <ICMS20>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
           strLinha = strLinha & "<modBC>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
           strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBC>"
           strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
           strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
           strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
           
           ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
           If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
              vBC_fcp = BaseCalculoICMS
              vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
    
              ' [3121;vBCFCP;pFCP;vFCP]
              strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
              strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCP>"
              strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
              vFCP_registroTotal = vFCP_registroTotal + vResultado
           End If
           
           strLinha = strLinha & "</ICMS20>"
           AddLinha clcArquivo, strLinha
                             
         Case "30"
           'Dados do icms da operação
           strLinha = strLinha & "<ICMS30>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
           strLinha = strLinha & "<modBCST>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBCST>"
           'strLinha = strLinha & "              <pMVAST>" & FormataValorTexto(MargemValorAdicionadoICMS, 2) & "</pMVAST>"
           strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
           strLinha = strLinha & " <vBCST>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBCST>"
           strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMSST>"
           strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS, 2) & "</vICMSST>"
           
           ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
           If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
              'Dim vAuxiliar As Double
              vAuxiliar = BaseCalculoICMS_ST
              vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
              ' [3131;vBCFCPST;pFCPST;vFCPST]
              strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
              strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPST>"
              strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
              vFCPST_registroTotal = vFCPST_registroTotal + vResultado
           End If
           
           strLinha = strLinha & "</ICMS30>"
           AddLinha clcArquivo, strLinha
         
         Case "40", "41", "50"
           strLinha = strLinha & "<ICMS40>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
           '''strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
           '''strLinha = strLinha & "<motDesICMS>" & "1" & "</motDesICMS>" '  1 - Taxi
           strLinha = strLinha & "</ICMS40>"
           AddLinha clcArquivo, strLinha

         Case "51"
           'Dados do icms da operação
           strLinha = strLinha & "<ICMS51>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
           strLinha = strLinha & "<modBC>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
           strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBC>"
           strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
           strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
           strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"

           If BaseCalculoICMS <= 0 Then
             ' Entra aqui se não é ST
             ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
             If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
    
                vBC_fcp = BaseCalculoICMS
                vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                ' [3146;vBCFCP;pFCP;vFCP]
                strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCP>"
                strLinha = strLinha & "<vFCP>" & FormataValorTexto(vResultado, 2) & "</vFCP>"
                vFCP_registroTotal = vFCP_registroTotal + vResultado
             End If
           End If
           
           strLinha = strLinha & "</ICMS51>"
           AddLinha clcArquivo, strLinha
         
         Case "60"
           'Dados do icms da operação
           strLinha = strLinha & "<ICMS60>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
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
           
'''           ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
'''           If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
'''              'Dim vAuxiliar As Double
'''              vAuxiliar = BaseCalculoICMS_ST
'''              vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
'''              ' [3151;vBCFCPSTRet;pFCPSTRet;vFCPSTRet]
'''              strLinha = strLinha & "<vBCFCPSTRet>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPSTRet>"
'''              strLinha = strLinha & "<pFCPSTRet>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPSTRet>"
'''              strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vResultado, 2) & "</vFCPSTRet>"
'''              vFCPSTRet_registroTotal = vFCPSTRet_registroTotal + vResultado
'''           End If
           
           strLinha = strLinha & "</ICMS60>"
           AddLinha clcArquivo, strLinha
         
         Case "70"
           strLinha = strLinha & "<ICMS70>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
           strLinha = strLinha & "<modBC>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
           strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
           strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS, 2) & "</pICMS>"
           strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
           
           ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
           If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                'Dim vAuxiliar As Double
                vAuxiliar = BaseCalculoICMS_ST
      
                ' COLOCAR GRUPO 3161           ou           3162
                If vAuxiliar = 0 Then
                    vBC_fcp = BaseCalculoICMS
                    vResultado = (vBC_fcp * AliquotaAdesaoUFDestFCP) / 100
                    ' [3161;vBCFCP;pFCP;vFCP]
                    strLinha = strLinha & "<vBCFCP>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCP>"
                    strLinha = strLinha & "<pFCP>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCP>"
                    strLinha = strLinha & "<vBC>" & FormataValorTexto(vResultado, 2) & "</vBC>"
                    vFCP_registroTotal = vFCP_registroTotal + vResultado
                End If
           End If
           
           strLinha = strLinha & "<modBCST>" & rsFilial.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
           'strLinha = strLinha & "              <pMVAST>" & FormataValorTexto(MargemValorAdicionadoICMS, 2) & "</pMVAST>"
           strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
           strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
           strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 2) & "</pICMSST>"
           strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
           
           ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
           If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                'Dim vAuxiliar As Double
                vAuxiliar = BaseCalculoICMS_ST
      
                ' COLOCAR GRUPO 3161           ou           3162
                If vAuxiliar <> 0 Then
                    vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                    ' [3162;vBCFCPST;pFCPST;vFCPST]
                    strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                    strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPST>"
                    strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                    vFCPST_registroTotal = vFCPST_registroTotal + vResultado
                End If
           End If
           
           strLinha = strLinha & "</ICMS70>"
           AddLinha clcArquivo, strLinha
        
         Case "90"
           'Dados do icms da operação
           strLinha = strLinha & "<ICMS90>"
           strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
           strLinha = strLinha & "<CST>" & Right(SituacaoTributaria, 2) & "</CST>"
           strLinha = strLinha & "</ICMS90>"
           AddLinha clcArquivo, strLinha

        End Select
      Else
        If dblPercentualReducaoBC_SN <> 0 Then
           dblBaseCalculoSimplesItem = ((.Fields("Qtde").Value * .Fields("Preço").Value) * dblPercentualReducaoBC_SN)
        Else
           dblBaseCalculoSimplesItem = (.Fields("Qtde").Value * .Fields("Preço").Value)
        End If
        
        Select Case str_CSO_Produto
        
            '20/09/2013 alterado a primeira linha de todos os cses para a variavel strLinha não sobrescrever a abertura das chaves de impostos e ICMS
            
          Case "101" 'Tributada pelo Simples Nacional com permissão de crédito
            strLinha = strLinha & "<ICMSSN101>"
            strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
            strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
            strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 2) & "</pCredSN>"
            strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
            strLinha = strLinha & "</ICMSSN101>"
            AddLinha clcArquivo, strLinha

          Case "102", "103", "300", "400"
            strLinha = strLinha & "<ICMSSN102>"
            strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
            strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
            strLinha = strLinha & "</ICMSSN102>"
            AddLinha clcArquivo, strLinha
            
          Case "201" 'Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por Substituição
            strLinha = strLinha & "<ICMSSN201>"
            strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
            strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
            strLinha = strLinha & "<modBCST>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBCST>"
            strLinha = strLinha & "<pMVAST>" & FormataValorTexto(MargemValorAdicionado_ICMS, 2) & "</pMVAST>"
            strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
            strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
            strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 2) & "</pICMSST>"
            strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
            
            ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
            If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                'Dim vAuxiliar As Double
                vAuxiliar = BaseCalculoICMS_ST
                vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                ' [3177a;vBCFCPST;pFCPST;vFCPST]
                strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPST>"
                strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                vFCPST_registroTotal = vFCPST_registroTotal + vResultado
            End If
            
            strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 2) & "</pCredSN>"
            strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
            strLinha = strLinha & "</ICMSSN201>"
            AddLinha clcArquivo, strLinha
            
          Case "202", "203"
            '202- Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por Substituição Tributária
            '203- Isenção do ICMS nos Simples Nacional para faixa de receita bruta e com cobrança do ICMS por Substituição Tributária
            strLinha = strLinha & "<ICMSSN202>"
            strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
            strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
            strLinha = strLinha & "<modBCST>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBCST>"
            strLinha = strLinha & "<pMVAST>" & FormataValorTexto(MargemValorAdicionado_ICMS, 2) & "</pMVAST>"
            strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
            strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
            strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 2) & "</pICMSST>"
            strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
            
            ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
            If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                'Dim vAuxiliar As Double
                vAuxiliar = BaseCalculoICMS_ST

                vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                ' [3178a;vBCFCPST;pFCPST;vFCPST]
                strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPST>"
                strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                vFCPST_registroTotal = vFCPST_registroTotal + vResultado
            End If
            
            strLinha = strLinha & "</ICMSSN202>"
            AddLinha clcArquivo, strLinha
            
          Case "500" 'ICMS cobrado anteriormente por substituição tributária (substituído) ou por antecipação
            strLinha = strLinha & "<ICMSSN500>"
            strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
            strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
            strLinha = strLinha & "<vBCSTRet>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCSTRet>"
        
            'Dim vTesteAux As Double
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
            If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then

                'Dim vAuxiliar As Double
                vAuxiliar = BaseCalculoICMS_ST
                vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                ' [3179a0;vBCFCPSTRet;pFCPSTRet;vFCPSTRet]
                strLinha = strLinha & "<vBCFCPSTRet>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPSTRet>"
                strLinha = strLinha & "<pFCPSTRet>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPSTRet>"
                strLinha = strLinha & "<vFCPSTRet>" & FormataValorTexto(vResultado, 2) & "</vFCPSTRet>"
                vFCPSTRet_registroTotal = vFCPSTRet_registroTotal + vResultado
            End If

            strLinha = strLinha & "</ICMSSN500>"
            AddLinha clcArquivo, strLinha

          Case "900" 'Isento
            strLinha = strLinha & "<ICMSSN900>"
            strLinha = strLinha & "<orig>" & Left(SituacaoTributaria, 1) & "</orig>"
            strLinha = strLinha & "<CSOSN>" & str_CSO_Produto & "</CSOSN>"
            strLinha = strLinha & "<modBC>" & rsFilial.Fields("ModDetBaseCalculoIcms").Value & "</modBC>"
            
            '''strLinha = strLinha & "<vBC>" & FormataValorTexto(ValorBaseCalculoICMS, 2) & "</vBC>"
            strLinha = strLinha & "<vBC>0.00</vBC>"
            
            
            If PercentualReducaoBaseICMS > 0 Then
              strLinha = strLinha & "<pRedBC>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBC>"
            End If
            
'''            If ValorBaseCalculoICMS <> 0 Then
'''                strLinha = strLinha & "<pICMS>" & FormataValorTexto(AliquotaICMS_Saida, 2) & "</pICMS>"
'''                ValorICMS = ((.Fields("Qtde").Value * .Fields("Preço").Value)) * liquotaICMS_Saida) / 100
'''                strLinha = strLinha & "<vICMS>" & FormataValorTexto(ValorICMS, 2) & "</vICMS>"
'''            Else
                strLinha = strLinha & "<pICMS>0.00</pICMS>"
                strLinha = strLinha & "<vICMS>0.00</vICMS>"
'''            End If
            
'''            iContaLoop = iContaLoop + 1
            
            strLinha = strLinha & "<modBCST>" & rsFilial.Fields("ModDetBaseCalculoIcmsSt").Value & "</modBCST>"
            If MargemValorAdicionado_ICMS > 0 Then
              strLinha = strLinha & "<pMVAST>" & FormataValorTexto(MargemValorAdicionado_ICMS, 2) & "</pMVAST>"
            End If
            If PercentualReducaoBaseICMS > 0 Then
              strLinha = strLinha & "<pRedBCST>" & FormataValorTexto(PercentualReducaoBaseICMS, 2) & "</pRedBCST>"
            End If
            strLinha = strLinha & "<vBCST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCST>"
            strLinha = strLinha & "<pICMSST>" & FormataValorTexto(AliquotaICMS_ST, 2) & "</pICMSST>"
            strLinha = strLinha & "<vICMSST>" & FormataValorTexto(ValorICMS_ST, 2) & "</vICMSST>"
            
            ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
            If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then

                'Dim vAuxiliar As Double
                vAuxiliar = BaseCalculoICMS_ST

                vResultado = (vAuxiliar * AliquotaAdesaoUFDestFCP) / 100
                ' [3179a2;vBCFCPST;pFCPST;vFCPST]
                strLinha = strLinha & "<vBCFCPST>" & FormataValorTexto(BaseCalculoICMS_ST, 2) & "</vBCFCPST>"
                strLinha = strLinha & "<pFCPST>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPST>"
                strLinha = strLinha & "<vFCPST>" & FormataValorTexto(vResultado, 2) & "</vFCPST>"
                vFCPST_registroTotal = vFCPST_registroTotal + vResultado
            End If
            
            strLinha = strLinha & "<pCredSN>" & FormataValorTexto(dblAliquotaSimples, 2) & "</pCredSN>"
            strLinha = strLinha & "<vCredICMSSN>" & FormataValorTexto((dblBaseCalculoSimplesItem * (dblAliquotaSimples / 100)), 2) & "</vCredICMSSN>"
            strLinha = strLinha & "</ICMSSN900>"
            AddLinha clcArquivo, strLinha
        End Select
      End If
      strLinha = "</ICMS>"
      AddLinha clcArquivo, strLinha
      If (rsFilial.Fields("CodigoRegimeTributario").Value) <> 1 Then
      
            ' ******************
            ' Empresa LUCRO REAL
            If Right(SituacaoTributaria, 2) = "00" Then
                
                ' ******************************************************
                ' CST acima é "00" - SIM...Destacar ICMS no PIS E COFINS
                ' ******************************************************
                
                If localTipoSituacaoTributariaPIS = 99 Then
                    'Pis
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISOutr>"
                    strLinha = strLinha & "<CST>99</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                        sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                        sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
                    End If
    
                    strLinha = strLinha & "<pPIS>0.0000</pPIS>"
                    strLinha = strLinha & "<vPIS>0.00</vPIS>"
                    strLinha = strLinha & "</PISOutr>"
                    strLinha = strLinha & "</PIS>"
    
                    'Cofins
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSOutr>"
                    strLinha = strLinha & "<CST>99</CST>"
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                        sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                        sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        'strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        'strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualCOFINS), 2) & "</vCOFINS>"
                        
                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualCOFINS), 2)
                    End If

                    strLinha = strLinha & "</COFINSAliq>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                    
                    'inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualPIS), 2)
                    'inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualCOFINS), 2)
                
                Else
                
                    'Tratando o PIS/COFINS 50, 73...por enquanto
                    
'                    strLinha = "<PIS>"
'                    strLinha = strLinha & "<PISAliq>"
'                    strLinha = strLinha & "<CST>" & "01" & "</CST>"
'                    strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
'                    strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
'                    strLinha = strLinha & "<vPIS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualPIS), 2) & "</vPIS>"
'                    strLinha = strLinha & "</PISAliq>"
'                    strLinha = strLinha & "</PIS>"
'
'                    strLinha = strLinha & "<COFINS>"
'                    strLinha = strLinha & "<COFINSAliq>"
'                    strLinha = strLinha & "<CST>" & "01" & "</CST>"
'                    strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value, 2) & "</vBC>"
'                    strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
'                    strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((.Fields("Preço Final").Value * PercentualCOFINS), 2) & "</vCOFINS>"
'                    strLinha = strLinha & "</COFINSAliq>"
'                    strLinha = strLinha & "</COFINS>"
'                    AddLinha clcArquivo, strLinha
'
'                    inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualPIS), 2)
'                    inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((.Fields("Preço Final").Value * PercentualCOFINS), 2)
                    
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISAliq>"
                    strLinha = strLinha & "<CST>01</CST>"
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        'strLinha = strLinha & "<vBC>" & FormataValorTexto(.Fields("Preço Final").Value - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
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
                
            ElseIf Right(SituacaoTributaria, 2) = "40" Or _
                Right(SituacaoTributaria, 2) = "60" Then
                
                ' ****************************************************
                ' CST acima é "40" - NÃO...Não destacar ICMS no PIS E COFINS
                ' CST acima é "60" - NÃO...Não destacar ICMS no PIS E COFINS
                ' ****************************************************
                
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
                    strLinha = strLinha & "<vBC>0.00</vBC>"
                    strLinha = strLinha & "<pPIS>0.0000</pPIS>"
                    strLinha = strLinha & "<vPIS>0.00</vPIS>"
                    strLinha = strLinha & "</PISAliq>"
                    strLinha = strLinha & "</PIS>"

                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSAliq>"
                    strLinha = strLinha & "<CST>0" & localTipoSituacaoTributariaPIS & "</CST>"
                    strLinha = strLinha & "<vBC>0.00</vBC>"
                    strLinha = strLinha & "<pCOFINS>0.0000</pCOFINS>"
                    strLinha = strLinha & "<vCOFINS>0.00</vCOFINS>"
                    strLinha = strLinha & "</COFINSAliq>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                   
                Else
                    strLinha = "<PIS>"
                    strLinha = strLinha & "<PISNT>"
                    strLinha = strLinha & "<CST>07</CST>"
                    strLinha = strLinha & "</PISNT>"
                    strLinha = strLinha & "</PIS>"
          
                    'Cofins
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSNT>"
                    strLinha = strLinha & "<CST>07</CST>"
                    strLinha = strLinha & "</COFINSNT>"
                    strLinha = strLinha & "</COFINS>"
                    AddLinha clcArquivo, strLinha
                End If
            
            Else
            
                ' ****************************************************
                ' CST acima é "10"
                ' CST acima é "20"
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                        sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                        sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
                    End If
    
                    strLinha = strLinha & "<pPIS>0.0000</pPIS>"
                    strLinha = strLinha & "<vPIS>0.00</vPIS>"
                    strLinha = strLinha & "</PISOutr>"
                    strLinha = strLinha & "</PIS>"
    
                    'Cofins
                    strLinha = strLinha & "<COFINS>"
                    strLinha = strLinha & "<COFINSOutr>"
                    strLinha = strLinha & "<CST>99</CST>"
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                        sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                        sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                    Else
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pPIS>" & FormataValorTexto((PercentualPIS * 100), 2) & "</pPIS>"
                        strLinha = strLinha & "<vPIS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2) & "</vPIS>"
                        
                        inst_dblPercentualPIS = inst_dblPercentualPIS + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualPIS, 2)
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
                    
                    If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                        strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                        strLinha = strLinha & "<pCOFINS>" & FormataValorTexto((PercentualCOFINS * 100), 2) & "</pCOFINS>"
                        strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2) & "</vCOFINS>"

                        inst_dblPercentualCofins = inst_dblPercentualCofins + FormataValorTextoComVirgula((BaseCalculoICMS - CDbl(sDescontoAuxZ_2)) * PercentualCOFINS, 2)
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
                
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                    iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                    sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                    sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                    strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                Else
                    strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
                End If

                strLinha = strLinha & "<pPIS>0.0000</pPIS>"
                strLinha = strLinha & "<vPIS>0.00</vPIS>"
                strLinha = strLinha & "</PISOutr>"
                strLinha = strLinha & "</PIS>"

                'Cofins
                strLinha = strLinha & "<COFINS>"
                strLinha = strLinha & "<COFINSOutr>"
                strLinha = strLinha & "<CST>99</CST>"
                If sDescontoAuxZ <> "<vDesc>0.00</vDesc>" And sDescontoAuxZ <> "" And BaseCalculoICMS <> 0 Then
                    iDescontoAuxZ_2 = InStr(1, sDescontoAuxZ, "</vDesc>")
                    sDescontoAuxZ_2 = Mid(sDescontoAuxZ, 8, iDescontoAuxZ_2 - 8)
                    sDescontoAuxZ_2 = Replace(sDescontoAuxZ_2, ".", ",")
                    strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS - CDbl(sDescontoAuxZ_2), 2) & "</vBC>"
                Else
                    strLinha = strLinha & "<vBC>" & FormataValorTexto(BaseCalculoICMS, 2) & "</vBC>"
                End If
                
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
        
        
        ' Grupo criado tratar FCP NFe 4.0 PILATTI/2018
        ' Codigo de regime tributário = 1 (Simples Nacional)
        ' Ressaltamos que as empresas optantes pelo Simples Nacional, deverão
        ' preencher/enviar as tags do grupo ICMSUFDest normalmente, independente dos valores estarem zerados, devido
        ' a liminar que isentou as empresas do Simples do pagamento do imposto referente à partilha do ICMS.
        If blnSimplesNacionalGeral = False Then
            ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
            If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
                ' Novo grupo ICMSUFDest
                ' [3179b;vBCUFDest;vBCFCPUFDest;pFCPUFDest;pICMSUFDest;pICMSInter;pICMSInterPart;vFCPUFDest;vICMSUFDest;vICMSUFRemet]
                strLinha = "<ICMSUFDest>"
                                
                vBC_fcp = .Fields("Qtde").Value * .Fields("Preço").Value
                'vBCUFDest ------------------------
                strLinha = strLinha & "<vBCUFDest>" & FormataValorTexto(vBC_fcp, 2) & "</vBCUFDest>"
                'vBCFCPUFDest ------------------------
                strLinha = strLinha & "<vBCFCPUFDest>" & FormataValorTexto(vBC_fcp, 2) & "</vBCFCPUFDest>"
                'pFCPUFDest ------------------------
                strLinha = strLinha & "<pFCPUFDest>" & FormataValorTexto(AliquotaAdesaoUFDestFCP, 2) & "</pFCPUFDest>"
                'pICMSUFDest ------------------------
                strLinha = strLinha & "<pICMSUFDest>" & FormataValorTexto(AliquotaFCP_Destinatario, 2) & "</pICMSUFDest>"
                'pICMSInter ------------------------
                strLinha = strLinha & "<pICMSInter>" & FormataValorTexto(AliquotaFCP_EMIT_DEST, 2) & "</pICMSInter>"
                
                'pICMSInterPart ------------------------
                If iAnoFCP = 2018 Then
                    strLinha = strLinha & "<pICMSInterPart>80.00</pICMSInterPart>"   ' 80%
                Else
                    strLinha = strLinha & "<pICMSInterPart>100.00</pICMSInterPart>"  ' 100% para 2019 em diante
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

        strLinha = strLinha & "</imposto>"
        strLinha = strLinha & "</det>"
        AddLinha clcArquivo, strLinha
        strSQL = ""
        .MoveNext
      Loop
    .Close
  End With
  Set rsSaidaProdutos = Nothing
  
  rsSaidasAUXILIAR_ICMS.Close
  Set rsSaidasAUXILIAR_ICMS = Nothing
  
End Function

Public Function CancelaNFCe(ByVal Sequencia As Long) As Boolean
On Error GoTo Erro:
  Dim frmX As Form
  Dim rsSaida As Recordset
  Dim iIndice01 As Integer
  Dim iIndice02 As Integer
  Dim strChave As String

  Set frmX = New frmMotivoCancelamento
  frmX.Show vbModal
  Set frmX = Nothing
  If gsRetornoDoc <> "OK" Then
    DisplayMsg "Nota Fiscal ao Consumidor eletrônica não cancelada."
    gsRetornoDoc = "N_OK"
    Exit Function
  End If

  Set rsSaida = db.OpenRecordset("Select * from [Saídas] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")

  If InStr(rsSaida("retNFCe").Value, "<statusAutorizacao>OK</statusAutorizacao>") _
    Or InStr(rsSaida("retNFCe").Value, "<statusAutorizacao>Pendente</statusAutorizacao>") _
    Or InStr(rsSaida("retNFCe").Value, "<detalheAutorizacao>100") Then

      iIndice01 = InStr(1, rsSaida("retNFCe").Value, "<chave>")
      If iIndice01 > 0 Then
          iIndice02 = InStr(1, rsSaida("retNFCe").Value, "</chave>")
          strChave = Mid$(rsSaida("retNFCe").Value, iIndice01 + 7, iIndice02 - (iIndice01 + 7))
      Else
          MsgBox "Não foi possível realizar a solicitação de cancelamento, pois a Chave da NFCe não foi encontrada na base de dados", vbInformation, "CANCELAMENTO DE CUPOM FISCAL"
          gsRetornoDoc = "N_OK"
          Exit Function
      End If

      CancelarXML_SEFAZ Sequencia, gCNPJ_CPFControleDeLicencaWebApi, strChave, strMotivoCancelamento
  
      If gRetOID = 0 Then
          gsRetornoDoc = "OK"
      Else
          gsRetornoDoc = "N_OK"
      End If
  Else
      DisplayMsg "NFC-e só pode ser Cancelada depois de autorizada. Esta venda não tem CUPOM FISCAL autorizado."
      gsRetornoDoc = "N_OK"
      Exit Function
  End If

  Exit Function
Erro:
  gsRetornoDoc = "N_OK"
  MsgBox "Erro na função CancelarNFCe " & Err.Number & " " & Err.Description, vbCritical, "CUPOM FISCAL CANCELAMENTO"
End Function

Public Function ReimprimeNFCe(ByVal Sequencia As Long, ByVal Vdesc As Double) As Boolean
    Dim clcArquivo As New Collection
  Dim strSQL As String
  Dim dblPeso As Double
  Dim rsSaida As Recordset
  Dim rsFilial As Recordset
  Dim rsCliFor As Recordset
  Dim rsOpSaida As Recordset
  Dim rsTransportadora As Recordset
  Dim strLinha As String
  Dim strChaveNFCe As String
  Dim strCodigoAleatorioChaveNFCe As String
  Dim intDV As Integer
  Dim lngNumNFCe As Long
  Dim blnShowObs As Boolean
  Dim frmX As Form
  Dim intX As Integer
  Dim intSerieNFC As Integer
  Dim dblPercentualPIS As Double
  Dim dblPercentualCofins As Double
  Dim dblAliquotaSimples As Double
  Dim dblPercentualReducaoSimplesNacional As Double
  Dim TpEmis As Integer
  Dim mostraCliente As Boolean
  Dim Resposta As Integer
  mostraCliente = False
  Dim somenteCPF_CNPJ As Boolean
  somenteCPF_CNPJ = False
  Dim CPF_CnpjCliente As String
  Dim IE_Cliente As String
  Dim PessoaJuridica As String
  PessoaJuridica = False
  Dim IndicadorIE As String
  Dim IE As String
  Dim strNomeCliente As String
  
  str_FCP_UF_destinatario = ""
    
  bTEM_ICMS_3179a = False
  
  strNomeCliente = ""
  
  TpEmis = 1
  
  Set rsSaida = db.OpenRecordset("Select * from [Saídas] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
  Call IsDataType(dtLong, rsSaida.Fields("NFCe").Value, lngNumNFCe)
  If lngNumNFCe = 0 Then
    MsgBox "Venda sem NFCE, não pode Reimprimir", vbExclamation, "Atenção"
  End If
  
  Set rsFilial = db.OpenRecordset("Select * from [Parâmetros Filial] Where Filial = " & gnCodFilial & "")
  Set rsCliFor = db.OpenRecordset("Select * from [Cli_For] Where Código = " & rsSaida("Cliente") & "")
  Set rsOpSaida = db.OpenRecordset("Select * from [Operações Saída] where Código = " & rsSaida("Operação"))
  
  If GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") = 0 Then
    MsgBox "Código IBGE inválido para a UF da Filial: " & rsFilial.Fields("Filial").Value, vbExclamation, "Atenção"
    rsFilial.Close
    Set rsFilial = Nothing
    Exit Function
  End If
  
  ' Pilatti Outubro/17
  sEstadoFilial = rsFilial.Fields("Estado").Value
  
  ' *******************************************
  ' PILATTI Outubro/17
  ' Tratar horario de verão
  Dim sFuso As String
  sFuso = ""
  'bHorarioDeVerao = True    ' Sim, estamos em HORARIO DE VERÃO
  func_horarioFuso bHorarioDeVerao, sEstadoFilial, sFuso
  ' *******************************************
  
  'Valida CNPJ
  If (Trim(rsFilial.Fields("CGC").Value) & "") = "" Or Not Valid_CGC(rsFilial.Fields("CGC").Value & "") Then
    MsgBox "CNPJ inválido para a Filial: " & rsFilial.Fields("Filial").Value, vbExclamation, "Atenção"
    rsFilial.Close
    Set rsFilial = Nothing
    Exit Function
  End If
  
  If rsFilial.Fields("CodigoRegimeTributario").Value = 1 Then
    blnSimplesNacionalGeral = True
  Else
    blnSimplesNacionalGeral = False
  End If
  
  If GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") = 0 Then
    MsgBox "Código IBGE inválido para a UF do Cliente: " & rsSaida.Fields("Cliente").Value, vbExclamation, "Atenção"
    rsFilial.Close
    Set rsCliFor = Nothing
    Exit Function
  End If
  
    '05/03/2010 - mpdea
    'Incluído validação de CPF
    '
    'Valida CNPJ
'    If PessoaJuridica = True Then
'      If (Trim(CPF_CnpjCliente) & "") = "" Or Not Valid_CGC(CPF_CnpjCliente & "") Then
'        MsgBox "CNPJ inválido para o Cliente: " & rsSaida.Fields("Cliente").Value, vbExclamation, "Atenção"
'        rsCliFor.Close
'        Set rsCliFor = Nothing
'        Exit Function
'      End If
'    Else
'      If (Trim(CPF_CnpjCliente) & "") = "" Or Not Valid_CPF(CPF_CnpjCliente & "") Then
'        MsgBox "CPF inválido para o Cliente: " & rsSaida.Fields("Cliente").Value, vbExclamation, "Atenção"
'        rsCliFor.Close
'        Set rsCliFor = Nothing
'        Exit Function
'      End If
'    End If
'  End If
  
  
'  Call IsDataType(dtLong, rsSaida.Fields("NFCe").Value, lngNumNFCe)
'  If lngNumNFCe = 0 Then
'    lngNumNFCe = g_longNextNFCe(rsFilial("Filial"))
'    rsSaida.Edit
'    rsSaida("NFCe").Value = lngNumNFCe
'    rsSaida.Update
'  End If
  
  Call IsDataType(dtInteger, rsSaida.Fields("SerieNF").Value, intSerieNFC)
'    If intSerieNFC = 0 Then
'      intSerieNFC = 1
'      rsSaida.Edit
'      rsSaida.Fields("SerieNF").Value = intSerieNFC
'      rsSaida.Update
'    End If
    
  Dim intTipoEmissao As Integer
  intTipoEmissao = GetTipoEmissao(rsFilial.Fields("Filial").Value, Sequencia, 1)
    
  strChaveNFCe = GetChaveAcesso(rsFilial.Fields("Filial").Value, Sequencia, 1)
  
  
  '--------------------------------------------------------------------------------------------------
  ' A- Dados da Nota Fiscal Consumidor eletrônica
  '--------------------------------------------------------------------------------------------------
  
  Dim timezone As String
  
  'timezone = "-03:00"
  timezone = sFuso
  
  
  'Dim utc As String = System.TimeZoneInfo.Local.ToString
  'Dim timezone As String = Mid(utc, 5, 6)
  'If System.TimeZoneInfo.Local.IsDaylightSavingTime(Date.Now()) Then
  '    timezone = "-02:00"
  'End If
  
  strLinha = "<NFe xmlns=""http://www.portalfiscal.inf.br/nfe"">"
  strLinha = strLinha & "<infNFe Id=""NFe" & strChaveNFCe & """ versao=""4.00"">"
  strLinha = strLinha & "<ide>"
  strLinha = strLinha & "<cUF>" & GetTerritorioUfCodigoIbge(rsFilial.Fields("Estado").Value & "") & "</cUF>"
  strLinha = strLinha & "<cNF>" & strCodigoAleatorioChaveNFCe & "</cNF>"
  strLinha = strLinha & "<natOp>" & Left(rsOpSaida.Fields("Nome").Value, 60) & "</natOp>"
    
'  'Itentificação da forma de pagamento
'  If rsSaida.Fields("Total Vista").Value = rsSaida.Fields("Total").Value Then
'    strLinha = strLinha & "<indPag>0</indPag>"
'  ElseIf rsSaida.Fields("Total Prazo").Value = rsSaida.Fields("Total").Value Then
'    strLinha = strLinha & "<indPag>1</indPag>"
'  Else
'    strLinha = strLinha & "<indPag>2</indPag>"
'  End If
  
  strLinha = strLinha & "<mod>65</mod>"
  strLinha = strLinha & "<serie>" & intSerieNFC & "</serie>"
  strLinha = strLinha & "<nNF>" & lngNumNFCe & "</nNF>"
  'strLinha = strLinha & "<dhEmi>" & Format(rsSaida.Fields("DataEmissaoNota").Value, "yyyy-MM-dd") & "T" & Format(rsSaida.Fields("DataEmissaoNota").Value, "hh:mm:ss") & "-02:00</dhEmi>"
  
  'strLinha = strLinha & "<dhEmi>" & Format(rsSaida.Fields("DataEmissaoNota").Value, "yyyy-MM-dd") & "T" & Format(rsSaida.Fields("DataEmissaoNota").Value, "hh:mm:ss") & "-03:00</dhEmi>"
  strLinha = strLinha & "<dhEmi>" & Format(rsSaida.Fields("DataEmissaoNota").Value, "yyyy-MM-dd") & "T" & Format(rsSaida.Fields("DataEmissaoNota").Value, "hh:mm:ss") & sFuso & "</dhEmi>"
  
  strLinha = strLinha & "<tpNF>1</tpNF>"
  If Len(rsCliFor.Fields("Estado").Value) > 0 Then
    If rsFilial.Fields("Estado").Value = rsCliFor.Fields("Estado").Value Then
      strLinha = strLinha & "<idDest>1</idDest>"
      ElseIf rsFilial.Fields("Estado").Value = "EX" Then
      strLinha = strLinha & "<idDest>3</idDest>"
      Else
      strLinha = strLinha & "<idDest>2</idDest>"
    End If
  Else
    strLinha = strLinha & "<idDest>1</idDest>"
  End If
    
  str_FCP_UF_emitente = rsFilial.Fields("Estado").Value
  
  strLinha = strLinha & "<cMunFG>" & GetTerritorioMunicipioCodigoIbge(rsFilial.Fields("Estado").Value & "", rsFilial.Fields("Cidade").Value & "") & "</cMunFG>"
  strLinha = strLinha & "<tpImp>4</tpImp>"
  strLinha = strLinha & "<tpEmis>1</tpEmis>"
  strLinha = strLinha & " <cDV>" & intDV & "</cDV>"
  strLinha = strLinha & "<tpAmb>" & rsFilial.Fields("AmbienteNFe").Value & "</tpAmb>"
  If rsSaida.Fields("FinalidadeNFe").Value > 0 Then
    strLinha = strLinha & "<finNFe>" & rsSaida.Fields("FinalidadeNFe").Value & "</finNFe>"  '& Finalidade de emissão da NF-e, 1- NF-e normal/2-NF-e complementar/3 – NF-e de ajuste NOT_FOUND
  Else
    strLinha = strLinha & "<finNFe>1</finNFe>"
  End If
  If IsNull(rsSaida.Fields("Consumidor_Final").Value) Then
    strLinha = strLinha & "<indFinal>0</indFinal>"
  Else
    strLinha = strLinha & "<indFinal>" & rsSaida.Fields("Consumidor_Final").Value & "</indFinal>"
  End If
  If IsNull(rsSaida.Fields("Presenca_Comprador").Value) Then
    strLinha = strLinha & "<indPres>1</indPres>"
  Else
    strLinha = strLinha & "<indPres>" & rsSaida.Fields("Presenca_Comprador").Value & "</indPres>"
  End If
  strLinha = strLinha & "<procEmi>0</procEmi>"
  strLinha = strLinha & "<verProc>" & rsFilial.Fields("VersaoLayoutEnvio").Value & "</verProc>"
  strLinha = strLinha & "</ide>"
  AddLinha clcArquivo, strLinha
  
  '--------------------------------------------------------------------------------------------------
  'C - Identificação do Emitente da Nota Fiscal Eletrônica
  '--------------------------------------------------------------------------------------------------
  With rsFilial
    dblPercentualPIS = (.Fields("PIS").Value / 100)
    dblPercentualCofins = (.Fields("COFINS").Value / 100)
    dblAliquotaSimples = .Fields("PercentualSimplesNacional").Value
    dblPercentualReducaoSimplesNacional = (1 - (.Fields("PercentualReducaoBCSimplesNacional").Value / 100))
    
    strLinha = "<emit>"
    strLinha = strLinha & "<CNPJ>" & FormataCnpjCpf(.Fields("CGC").Value & "") & "</CNPJ>"
    strLinha = strLinha & "<xNome>" & Left(.Fields("Razão Social").Value, 60) & "</xNome>"
    strLinha = strLinha & "<xFant>" & Left(.Fields("Nome").Value, 60) & "</xFant>"
    strLinha = strLinha & "<enderEmit>"
    strLinha = strLinha & "<xLgr>" & Left(.Fields("Endereço").Value, 60) & "</xLgr>"
    strLinha = strLinha & "<nro>" & Left(.Fields("EnderecoNumero").Value, 60) & "</nro>"
   
    If Left(.Fields("EnderecoComplemento").Value, 60) <> "" Then
      strLinha = strLinha & "<xCpl>" & Left(.Fields("EnderecoComplemento").Value, 60) & "</xCpl>"
    End If
   
   
    strLinha = strLinha & "<xBairro>" & Left(.Fields("Bairro").Value, 60) & "</xBairro>"
   
    If UCase(.Fields("Estado").Value) = "EX" Then
      strLinha = strLinha & "<cMun>" & "9999999" & "</cMun>"
      strLinha = strLinha & "<xMun>" & "EXTERIOR" & "</xMun>"
      strLinha = strLinha & "<UF>" & "EX" & "</UF>"
    Else
      strLinha = strLinha & "<cMun>" & GetTerritorioMunicipioCodigoIbge(.Fields("Estado").Value, .Fields("Cidade").Value & "") & "</cMun>"
      strLinha = strLinha & "<xMun>" & Left(.Fields("Cidade").Value, 60) & "</xMun>"
      strLinha = strLinha & "<UF>" & .Fields("Estado").Value & "</UF>"
    End If
   
    strLinha = strLinha & "<CEP>" & FormataNumero(.Fields("CEP").Value & "", 8) & "</CEP>"
    If Len(.Fields("Pais").Value) > 0 Then
      strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(.Fields("Pais").Value & ""), 4) & "</cPais>"
      strLinha = strLinha & "<xPais>" & Left(.Fields("Pais").Value, 60) & "</xPais>"
    Else
      strLinha = strLinha & "<cPais>1058</cPais>"
      strLinha = strLinha & "<xPais>Brasil</xPais>"
    End If
   
    If Left(RetornaSomenteNumero(.Fields("Fone").Value), 10) <> "" Then
      strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(.Fields("Fone").Value), 10) & "</fone>"
    End If
   
    strLinha = strLinha & "</enderEmit>"
  
    strLinha = strLinha & "<IE>" & Left(FormataIE(.Fields("Inscrição").Value, False), 14) & "</IE>"
   
    If rsSaida.Fields("Produtos").Value > 0 And rsSaida.Fields("Serviços").Value > 0 Then
      strLinha = strLinha & "<IM>" & Left(.Fields("InscricaoMunicipal").Value, 15) & "</IM>"
      strLinha = strLinha & "<CNAE>" & Left(.Fields("CNAE").Value, 7) & "</CNAE>"
    End If
     
    strLinha = strLinha & "<CRT>" & rsFilial.Fields("CodigoRegimeTributario").Value & "</CRT>"
    strLinha = strLinha & "</emit>"
    AddLinha clcArquivo, strLinha
  End With
  
  Dim sTrataDado As String
  
  '--------------------------------------------------------------------------------------------------
  'C - Identificação do Destinatario da Nota Fiscal Eletrônica
  '--------------------------------------------------------------------------------------------------
  If rsSaida.Fields("Emitiu_Somente_CPF_Cliente_NFCe") = True Then
    strLinha = strLinha & "<dest>"
    If Len(rsSaida.Fields("CPF_CPNJ_Cliente")) = 14 Then
      PessoaJuridica = True
    Else
      PessoaJuridica = False
    End If
    If PessoaJuridica = True Then
      strLinha = strLinha & "<CNPJ>" & rsSaida.Fields("CPF_CPNJ_Cliente") & "</CNPJ>"
    Else
      strLinha = strLinha & "<CPF>" & rsSaida.Fields("CPF_CPNJ_Cliente") & "</CPF>"
    End If
    If Len(rsSaida.Fields("Nome_Cliente")) > 0 Then
      sTrataDado = trataCaracteresEspeciaisNFE(rsSaida.Fields("Nome_Cliente"))
      strLinha = strLinha & "<xNome>" & sTrataDado & "</xNome>"
    End If
    IE = Left(IE_Cliente, 14)
    
    str_FCP_RG_INS_EST_destinatario = IE
    
    If IsNumeric(IE) Then
      strLinha = strLinha & "<indIEDest>1</indIEDest>"
      IndicadorIE = "1"
    ElseIf IE = "ISENTO" Then
      strLinha = strLinha & "<indIEDest>2</indIEDest>"
      IndicadorIE = "2"
    Else
      strLinha = strLinha & "<indIEDest>9</indIEDest>"
      IndicadorIE = "9"
    End If
    If IndicadorIE = "1" Then
      strLinha = strLinha & "<IE>" & Left(IE_Cliente, 14) & "</IE>"
    End If
    strLinha = strLinha & "</dest>"
        
    'Adicionar linha
    AddLinha clcArquivo, strLinha
  End If
  
  
  If mostraCliente = True Then
    'Validar se é para mostrar dados do cliente
    With rsCliFor
      strLinha = strLinha & "<dest>"
      If (.Fields("Física_Jurídica").Value) = "F" Then
        strLinha = strLinha & "<CPF>" & CPF_CnpjCliente & "</CPF>"
      Else
        strLinha = strLinha & "<CNPJ>" & CPF_CnpjCliente & "</CNPJ>"
      End If
      strNomeCliente = Left(.Fields("Nome").Value, 60)
      strNomeCliente = trataCaracteresEspeciaisNFE(strNomeCliente)
      strLinha = strLinha & "<xNome>" & strNomeCliente & "</xNome>"
      
      strLinha = strLinha & "<enderDest>"
      sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Endereço").Value, 60))
      strLinha = strLinha & "<xLgr>" & sTrataDado & "</xLgr>"
      strLinha = strLinha & "<nro>" & Left(.Fields("Endereço Número").Value, 60) & "</nro>"
      
      If Len(.Fields("Complemento").Value) <> 0 Then
        sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Complemento").Value, 60))
        strLinha = strLinha & "<xCpl>" & sTrataDado & "</xCpl>"
      End If
      
      If Left(.Fields("Bairro").Value, 60) = "" Then
        strLinha = strLinha & "<xBairro>" & "CENTRO" & "</xBairro>"
      Else
        sTrataDado = trataCaracteresEspeciaisNFE(Left(.Fields("Bairro").Value, 60))
        strLinha = strLinha & "<xBairro>" & sTrataDado & "</xBairro>"
      End If
      
      If UCase(.Fields("Estado").Value) = "EX" Then
        strLinha = strLinha & "<cMun>" & "9999999" & "</cMun>"
        strLinha = strLinha & "<xMun>" & "EXTERIOR" & "</xMun>"
        strLinha = strLinha & "<UF>" & "EX" & "</UF>"
      Else
        strLinha = strLinha & "<cMun>" & GetTerritorioMunicipioCodigoIbge(.Fields("Estado").Value, .Fields("Cidade").Value & "") & "</cMun>"
        strLinha = strLinha & "<xMun>" & Left(.Fields("Cidade").Value, 60) & "</xMun>"
        strLinha = strLinha & "<UF>" & .Fields("Estado").Value & "</UF>"
      End If
      
      str_FCP_UF_destinatario = .Fields("Estado").Value
      
      strLinha = strLinha & "<CEP>" & FormataNumero(.Fields("CEP").Value & "", 8) & "</CEP>"
      If Len(.Fields("Pais").Value) > 0 Then
        strLinha = strLinha & "<cPais>" & FormataNumero(GetTerritorioPaisCodigoBacen(.Fields("Pais").Value & ""), 4) & "</cPais>"
        strLinha = strLinha & "<xPais>" & Left(.Fields("Pais").Value, 60) & "</xPais>"
      Else
        strLinha = strLinha & "<cPais>1058</cPais>"
        strLinha = strLinha & "<xPais>Brasil</xPais>"
      End If
      
      If Left(RetornaSomenteNumero(.Fields("Fone 1").Value), 10) <> "" Then
        strLinha = strLinha & "<fone>" & Left(RetornaSomenteNumero(.Fields("Fone 1").Value), 10) & "</fone>"
      End If
      
      strLinha = strLinha & "</enderDest>"
      
      'Dim IndicadorIE As String
      'Dim IE As String
      IE = Left(IE_Cliente, 14)
      If IsNumeric(IE) Then
        strLinha = strLinha & "<indIEDest>1</indIEDest>"
        IndicadorIE = "1"
      ElseIf IE = "ISENTO" Then
        strLinha = strLinha & "<indIEDest>2</indIEDest>"
        IndicadorIE = "2"
      Else
        strLinha = strLinha & "<indIEDest>9</indIEDest>"
        IndicadorIE = "9"
      End If
    '      If IsNull(.Fields("IndicadorIE").Value) Then
    '        strLinha = strLinha & "9;"
    '      Else
    '        strLinha = strLinha & .Fields("IndicadorIE").Value & ";"
    '      End If
      If IndicadorIE = "1" Then
        strLinha = strLinha & "<IE>" & Left(IE_Cliente, 14) & "</IE>"
      End If
     
      'strLinha = strLinha & "        <IE>" & Left(FormataIE(.Fields("Inscrição").Value, UCase(.Fields("Física_Jurídica").Value) = "F"), 14) & "</IE>"
      
      If Len(.Fields("email").Value) <> 0 Then
        sTrataDado = trataCaracteresEspeciaisNFE(.Fields("email").Value)
        strLinha = strLinha & "<email>" & sTrataDado & "</email>"
      End If
    
      strLinha = strLinha & "</dest>"
        
      'Adicionar linha
      AddLinha clcArquivo, strLinha
    End With
  End If
  '--------------------------------------------------------------------------------------------------
  'I - Produtos e Serviços da NF-e
  '--------------------------------------------------------------------------------------------------
  iDetNumItemXML = 0
  AddProdutoSevicosSaida rsFilial.Fields("Filial").Value, Sequencia, clcArquivo, rsSaida, rsFilial, dblPercentualPIS, dblPercentualCofins, dblAliquotaSimples, dblPercentualReducaoSimplesNacional, rsOpSaida.Fields("CSO").Value, Vdesc
  '--------------------------------------------------------------------------------------------------
  
  '--------------------------------------------------------------------------------------------------
  'W - Valores Totais da NF-e
  '--------------------------------------------------------------------------------------------------
  Dim str_volumes_quantidade As String
  strLinha = "<total>"
  strLinha = strLinha & "<ICMSTot>"
  
  If (rsFilial.Fields("CodigoRegimeTributario").Value) <> 1 Then
    'strLinha = strLinha & "<vBC>" & FormataValorTexto(rsSaida.Fields("Base ICM").Value, 2) & "</vBC>"
    'strLinha = strLinha & "<vICMS>" & FormataValorTexto(rsSaida.Fields("Valor ICM").Value, 2) & "</vICMS>"
    
    If rsSaida.Fields("Base ICM").Value > 0 And vlSOMA_DO_ICMS > 0 Then
          strLinha = strLinha & "<vBC>" & FormataValorTexto(v_vlSOMA_DO_ICMS_BaseDeCalculo, 2) & "</vBC>"
    Else
          strLinha = strLinha & "<vBC>0.00</vBC>"
    End If
    
    If vlSOMA_DO_ICMS > 0 Then
        If vlSOMA_DO_ICMS <> rsSaida.Fields("Valor ICM").Value Then
          strLinha = strLinha & " <vICMS>" & FormataValorTexto(vlSOMA_DO_ICMS, 2) & "</vICMS>"
        Else
          strLinha = strLinha & "<vICMS>" & FormataValorTexto(rsSaida.Fields("Valor ICM").Value, 2) & "</vICMS>"
        End If
    Else
        strLinha = strLinha & "<vICMS>" & FormataValorTexto(rsSaida.Fields("Valor ICM").Value, 2) & "</vICMS>"
    End If
    
    strLinha = strLinha & "<vICMSDeson>" & FormataValorTexto(rsSaida.Fields("TotalDesoneracaoICMS").Value, 2) & "</vICMSDeson>"
    
' FCP tratamento NFe 4.0    Pilatti/2018
    ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
    If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
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
    
    strLinha = strLinha & "<vBCST>" & FormataValorTexto(rsSaida.Fields("Base ICM Subs").Value, 2) & "</vBCST>"
    strLinha = strLinha & "<vST>" & FormataValorTexto(rsSaida.Fields("Valor ICM Subs").Value, 2) & "</vST>"
  Else
    strLinha = strLinha & "<vBC>" & "0.00" & "</vBC>"
    strLinha = strLinha & "<vICMS>" & "0.00" & "</vICMS>"
    strLinha = strLinha & "<vICMSDeson>" & "0.00" & "</vICMSDeson>"
    strLinha = strLinha & "<vFCP>0.00</vFCP>"
    strLinha = strLinha & "<vBCST>" & "0.00" & "</vBCST>"
    strLinha = strLinha & "<vST>" & "0.00" & "</vST>"
  End If
  
  ' FCP tratamento NFe 4.0    Pilatti/2018
  ' Vale para Vendas Interestadual sem IE destinatario (Inscricao estadual), basicamente temos que incluir o grupo FCP
  If blnSimplesNacionalGeral = False Then
    If Len(str_FCP_UF_destinatario) > 0 And str_FCP_UF_emitente <> str_FCP_UF_destinatario And (str_FCP_RG_INS_EST_destinatario = "") Then ' Or isIEIsento = True) Then
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
  
  'strLinha = strLinha & "<vProd>" & FormataValorTexto(rsSaida.Fields("Produtos").Value + rsSaida.Fields("Serviços").Value, 2) & "</vProd>"
  strLinha = strLinha & "<vProd>" & FormataValorTexto(total_produto, 2) & "</vProd>"
  strLinha = strLinha & "<vFrete>" & FormataValorTexto(rsSaida.Fields("Frete").Value, 2) & "</vFrete>"
  strLinha = strLinha & "<vSeg>" & FormataValorTexto(rsSaida.Fields("Seguro").Value, 2) & "</vSeg>"
  
  If m_bln_tem_desconto = True Then
    strLinha = strLinha & "<vDesc>" & FormataValorTexto(rsSaida.Fields("Desconto").Value, 2) & "</vDesc>"
  Else
    strLinha = strLinha & "<vDesc>" & "0.00" & "</vDesc>"
  End If
  m_bln_tem_desconto = False
  
  strLinha = strLinha & "<vII>" & "0.00" & "</vII>"
  strLinha = strLinha & "<vIPI>" & FormataValorTexto(rsSaida.Fields("IPI").Value, 2) & "</vIPI>"
  
  ' NFe 4.0 Pilatti 2018
  'vIPIDevol
  strLinha = strLinha & "<vIPIDevol>0.00</vIPIDevol>"
  
  strLinha = strLinha & "<vPIS>" & FormataValorTexto((rsSaida.Fields("Total").Value * dblPercentualPIS), 2) & "</vPIS>"
  strLinha = strLinha & "<vCOFINS>" & FormataValorTexto((rsSaida.Fields("Total").Value * dblPercentualPIS), 2) & "</vCOFINS>"
  strLinha = strLinha & "<vOutro>" & "0.00" & "</vOutro>"
  'michel 18/03/2016
  'strLinha = strLinha & "<vNF>" & FormataValorTexto(rsSaida.Fields("Total").Value, 2) & "</vNF>"
  strLinha = strLinha & "<vNF>" & FormataValorTexto(total_produto, 2) & "</vNF>"
  If IsNull(rsSaida.Fields("TotalNCM").Value) Or rsSaida.Fields("TotalNCM").Value = 0 Then
  Else
    strLinha = strLinha & "<vTotTrib>" & FormataValorTexto(TotalAproxImpostos, 2) & "</vTotTrib>"
  End If
  strLinha = strLinha & "</ICMSTot>"
  strLinha = strLinha & "</total>"
  AddLinha clcArquivo, strLinha
  strSQL = ""
  
  '--------------------------------------------------------------------------------------------------
  'X-Informações do Transporte da NF-e
  '--------------------------------------------------------------------------------------------------
  strSQL = "SELECT [obs_FretePago] FROM [Saídas] WHERE Filial = " & gnCodFilial & " AND Sequência = " & Sequencia
  Set rsTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  If rsTransportadora.EOF Then
    strLinha = "<transp>"
    strLinha = strLinha & "<modFrete>" & "9" & "</modFrete>"
    strLinha = "</transp>"
  Else
    strLinha = "<transp>"
    
    
    '16ªlinha [5000;0]
    ' Modalidade do frete, 0 – por conta do emitente; 1 – por conta do destinatário
    If rsTransportadora.Fields("obs_FretePago").Value = 1 Then
      strLinha = strLinha & "<modFrete>" & "0" & "</modFrete>"
    ElseIf rsTransportadora.Fields("obs_FretePago").Value = 2 Then
      strLinha = strLinha & "<modFrete>" & "1" & "</modFrete>"
    Else
      strLinha = strLinha & "<modFrete>" & "9" & "</modFrete>"
    End If
    
    AddLinha clcArquivo, strLinha
    
    strSQL = ""
    rsTransportadora.Close
    
    Set rsTransportadora = Nothing
     
    'Transporta
    strSQL = "SELECT [Saídas].obs_Transportadora, Transportadoras.CGC, Transportadoras.Nome, "
    strSQL = strSQL & "Transportadoras.[inscrição], Transportadoras.[Endereço], Transportadoras.Cidade, "
    strSQL = strSQL & "Transportadoras.Estado FROM [Saídas], Transportadoras WHERE [Saídas].Filial = " & gnCodFilial & "  AND "
    strSQL = strSQL & "[Saídas].Sequência = " & Sequencia & " AND Transportadoras.Nome = [Saídas].obs_Transportadora"
     
    Set rsTransportadora = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
      
    If Not (rsTransportadora.BOF And rsTransportadora.EOF) Then
      
      strLinha = strLinha & "<transporta>"
      If (rsTransportadora.Fields("CGC").Value <> "") Then
      strLinha = strLinha & "<CNPJ>" & RetornaSomenteNumero(rsTransportadora.Fields("CGC").Value & "") & "</CNPJ>"
      End If
      If (rsTransportadora.Fields("Nome").Value <> "") Then
      strLinha = strLinha & "<xNome>" & rsTransportadora.Fields("Nome").Value & "</xNome>"
      End If
      strLinha = strLinha & "<IE>" & RetornaSomenteNumero(rsTransportadora.Fields("inscrição").Value & "") & "</IE>"
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
      AddLinha clcArquivo, strLinha
    
      If Not IsNull(rsSaida.Fields("obs_Placa").Value) Then
        If rsSaida.Fields("obs_Placa").Value <> "" Then
          strLinha = "<veicTransp>"
          strLinha = strLinha & "<placa>" & rsSaida.Fields("obs_Placa").Value & "</placa>"
          strLinha = strLinha & "<UF>" & rsSaida.Fields("obs_UF").Value & "</UF>"
          strLinha = strLinha & "</veicTransp>"
          AddLinha clcArquivo, strLinha
        End If
      End If
      
      strLinha = "<vol>"
      
      If IsNull(rsSaida.Fields("obs_Qtde").Value) Then
        str_volumes_quantidade = "0"
      Else
        str_volumes_quantidade = rsSaida.Fields("obs_Qtde").Value
      End If
       
      str_volumes_quantidade = Format(str_volumes_quantidade, String(15, "0"))
       
      If (str_volumes_quantidade <> "") Then
      strLinha = strLinha & "<qVol>" & str_volumes_quantidade & "</qVol>"
      End If
      
      If (rsSaida.Fields("obs_Especie").Value <> "") Then
        strLinha = strLinha & "<esp>" & rsSaida.Fields("obs_Especie").Value & "</esp>"
      End If
       
      If (rsSaida.Fields("obs_Marca").Value <> "") Then
        strLinha = strLinha & "<marca>" & rsSaida.Fields("obs_Marca").Value & "</marca>"
      End If
      
      'strLinha = strLinha & "          <nVol>" & "" & "</nVol>" 'Campo Numero de Volumes não encontrado
       
      Call IsDataType(dtDouble, rsSaida.Fields("obs_PesoLiquido").Value, dblPeso)
      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
      strLinha = strLinha & "<pesoL>" & FormataValorTexto(dblPeso, 3) & "</pesoL>"
      End If
      Call IsDataType(dtDouble, rsSaida.Fields("obs_PesoBruto").Value, dblPeso)
      If (FormataValorTexto(dblPeso, 3) <> 0# And FormataValorTexto(dblPeso, 3) <> "") Then
      strLinha = strLinha & "<pesoB>" & FormataValorTexto(dblPeso, 3) & "</pesoB>"
      End If
      strLinha = strLinha & "</vol>"
      AddLinha clcArquivo, strLinha
       
    Else
      MsgBox "Transportadora não encontrada para a Filial " & gnCodFilial & ", sequência " & Sequencia, vbExclamation, "Atenção"
      AddLinha clcArquivo, strLinha
    End If
    strSQL = ""
    rsTransportadora.Close
    
    Set rsTransportadora = Nothing
    
    strLinha = strLinha & "</transp>"
  End If
  AddLinha clcArquivo, strLinha
  '--------------------------------------------------------------------------------------------------
  'YA-Informações de pagamento
  '--------------------------------------------------------------------------------------------------
  
    strLinha = "<pag>"
  
  If rsSaida.Fields("Recebe - Dinheiro").Value > 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>0</indPag>"  '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>01</tPag>"
    
    If rsSaida.Fields("Recebe - Dinheiro") > rsSaida.Fields("Total") Then
      rsSaida.Fields("Recebe - Dinheiro") = rsSaida.Fields("Total")
    End If
    
    strLinha = strLinha & "<vPag>" & FormataValorTexto((rsSaida.Fields("Recebe - Dinheiro").Value), 2) & "</vPag>"
    strLinha = strLinha & "</detPag>"
  End If
  
  
  Dim rsCheques As Recordset
  Dim totalCheque As Integer
  totalCheque = 0
  Set rsCheques = db.OpenRecordset("Select * from [Movimento - Cheques] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
  Do Until rsCheques.EOF
    totalCheque = totalCheque + rsCheques.Fields("Valor")
    rsCheques.MoveNext
  Loop
  If totalCheque > 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>0</indPag>"  '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>02</tPag>"
    If totalCheque > rsSaida.Fields("Total") Then
      totalCheque = rsSaida.Fields("Total")
    End If
    strLinha = strLinha & "<vPag>" & FormataValorTexto(totalCheque, 2) & "</vPag>"
    strLinha = strLinha & "</detPag>"
  End If
  
  If rsSaida.Fields("TotalCartaoCredito").Value <> 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>1</indPag>"  '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>03</tPag>"
    
    Dim tmpTotalCartaoCredito As Double
    tmpTotalCartaoCredito = rsSaida.Fields("TotalCartaoCredito").Value
    If tmpTotalCartaoCredito > rsSaida.Fields("Total") Then
       tmpTotalCartaoCredito = rsSaida.Fields("Total").Value
    End If
    strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoCredito, 2) & "</vPag>"
    strLinha = strLinha & "<card><tpIntegra>2</tpIntegra></card>"
    strLinha = strLinha & "</detPag>"
  End If
  
  If rsSaida.Fields("TotalCartaoDebito").Value <> 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<indPag>0</indPag>"  '0-Pagamento à Vista;1-Pagamento à Prazo;
    strLinha = strLinha & "<tPag>04</tPag>"
    
    Dim tmpTotalCartaoDebito As Double
    tmpTotalCartaoDebito = rsSaida.Fields("TotalCartaoDebito").Value
    If tmpTotalCartaoDebito > rsSaida.Fields("Total") Then
        tmpTotalCartaoDebito = rsSaida.Fields("Total")
    End If
    strLinha = strLinha & "<vPag>" & FormataValorTexto(tmpTotalCartaoDebito, 2) & "</vPag>"
    strLinha = strLinha & "<card><tpIntegra>2</tpIntegra></card>"
    strLinha = strLinha & "</detPag>"
  End If
  
  Dim TotalOutros As Double
  If Not IsNull(rsSaida.Fields("Total Prazo").Value) Then
    TotalOutros = rsSaida.Fields("Total Prazo").Value
  End If
  
  If Not IsNull(rsSaida.Fields("Recebe - Vale").Value) Then
    TotalOutros = TotalOutros + rsSaida.Fields("Recebe - Vale").Value
  End If

  If TotalOutros > 0 Then
    strLinha = strLinha & "<detPag>"
    strLinha = strLinha & "<tPag>99</tPag>"
    If TotalOutros > rsSaida.Fields("Total") Then TotalOutros = rsSaida.Fields("Total")
    strLinha = strLinha & "<vPag>" & FormataValorTexto(TotalOutros, 2) & "</vPag>"
    strLinha = strLinha & "</detPag>"
  End If
  
  strLinha = strLinha & "</pag>"
  AddLinha clcArquivo, strLinha
  
  '--------------------------------------------------------------------------------------------------
  'Z-Informações adicionais
  '--------------------------------------------------------------------------------------------------
  Dim str_observacoes_NF As String
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & "Val Aprox dos Tributos R$"
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaida.Fields("TotalNCM").Value & "("
  Dim aliquotaNCM As Double
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then aliquotaNCM = (Format(rsSaida("TotalNCM"), "##,###,##0.00") / Format(rsSaida("Total"), "##,###,##0.00")) * 100
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then aliquotaNCM = Format(aliquotaNCM, "##,###,##0.00")
  If Not IsNull(rsSaida.Fields("TotalNCM").Value) Then str_observacoes_NF = str_observacoes_NF & aliquotaNCM & "%) Fonte IBPT."
  
  If Not IsNull(rsSaida.Fields("obs_infCpl1").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaida.Fields("obs_infCpl1").Value
  If Not IsNull(rsSaida.Fields("obs_infCpl2").Value) Then str_observacoes_NF = str_observacoes_NF & rsSaida.Fields("obs_infCpl2").Value
  
  str_observacoes_NF = Trim(str_observacoes_NF)
  '---------------------------------------------------------
  '07/04/2011 - Andrea
  'Incluída mensagem de tributacao, no primeiro campo depois da tag 7000
  'Carrega as mensagens para Nota Fiscal
  Dim str_observacoes_NF_fisco As String
  Set m_clcMensagens = New Collection
  Call GetMensagemNotaFiscal(gnCodFilial, Sequencia, m_clcMensagens)
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
  strLinha = "<infAdic>"
  If (str_observacoes_NF_fisco <> "") Then
    strLinha = strLinha & "<infAdFisco>" & str_observacoes_NF_fisco & "</infAdFisco>"
  End If
  strLinha = strLinha & "<infCpl>" & str_observacoes_NF & "</infCpl>"
  strLinha = strLinha & "</infAdic>"
  strLinha = strLinha & "</infNFe>"
  strLinha = strLinha & "</NFe>"
  'strLinha = strLinha & "</enviNFe>"
  AddLinha clcArquivo, strLinha
  
  If Not DirExists(".\NFCe\") Then
    MkDir (".\NFCe\")
  End If
  
  GravarArquivo clcArquivo, ".\NFCe\" & "NFCeEnvio_" & gCNPJ_CPFControleDeLicencaWebApi & "_" & lngNumNFCe & ".xml"
  Dim xml As Variant
  xml = ""
  'continuar
  Dim i As Integer
  Dim cont As Integer
  Dim Endereco_Emitente As String
  If Left(rsFilial.Fields("EnderecoComplemento").Value, 60) <> "" Then
    Endereco_Emitente = Left(rsFilial.Fields("Endereço").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoNumero").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoComplemento").Value, 60) & ", " & Left(rsFilial.Fields("Bairro").Value, 60) & ", " & Left(rsFilial.Fields("Cidade").Value, 60)
  Else
    Endereco_Emitente = Left(rsFilial.Fields("Endereço").Value, 60) & ", " & Left(rsFilial.Fields("EnderecoNumero").Value, 60) & ", " & Left(rsFilial.Fields("Bairro").Value, 60) & ", " & Left(rsFilial.Fields("Cidade").Value, 60)
  End If
  cont = clcArquivo.Count
  i = 1
  Dim BancoPDV As Database
  Set BancoPDV = OpenDatabase(rsFilial("BancoPDV").Value & "\QuickStore.mdb", False, False, ";PWD=" & gsGetPValue())
  Dim rsNFCe_Job As Recordset
  Set rsNFCe_Job = BancoPDV.OpenRecordset("Select * from [NFCE_job] where Chave = '" & strChaveNFCe & "' And cnpj = '" & FormataCnpjCpf(rsFilial.Fields("CGC").Value) & "'")
  If rsNFCe_Job.EOF Then
    rsNFCe_Job.AddNew
    rsNFCe_Job!CNPJ = FormataCnpjCpf(rsFilial.Fields("CGC").Value)
    Do Until i > cont
      xml = xml + Replace(CStr(clcArquivo.Item(i)), vbCrLf, "")
      'Xml = Xml + RTrim(LTrim(Replace(CStr(clcArquivo.Item(i)), vbCrLf, "")))
      i = i + 1
    Loop
    rsNFCe_Job!xml = xml
    rsNFCe_Job!Tipo = 4
    rsNFCe_Job!Serie = intSerieNFC
    rsNFCe_Job!N_NF = lngNumNFCe
    rsNFCe_Job!Chave = strChaveNFCe
    rsNFCe_Job!CPF = CPF_CnpjCliente
    rsNFCe_Job!Nome_Consumidor = strNomeCliente
    rsNFCe_Job!Data_Emissao = rsSaida.Fields("DataEmissaoNota").Value
    rsNFCe_Job!Total_Tributos = FormataValorTexto(rsSaida.Fields("TotalNCM").Value, 2)
    rsNFCe_Job!Nome_Emitente = Left(rsFilial.Fields("Razão Social").Value, 60)
    rsNFCe_Job!Endereco_Emitente = Endereco_Emitente
    rsNFCe_Job!IE_Emitente = FormataIE(rsFilial.Fields("Inscrição").Value, False)
    rsNFCe_Job.Update
  Else
    rsNFCe_Job.Edit
    rsNFCe_Job!CNPJ = FormataCnpjCpf(rsFilial.Fields("CGC").Value)
    Do Until i > cont
      xml = xml + Replace(CStr(clcArquivo.Item(i)), vbCrLf, "")
      'Xml = Xml + RTrim(LTrim(Replace(CStr(clcArquivo.Item(i)), vbCrLf, "")))
      i = i + 1
    Loop
    rsNFCe_Job!xml = xml
    rsNFCe_Job!Tipo = 4
    rsNFCe_Job!Serie = intSerieNFC
    rsNFCe_Job!N_NF = lngNumNFCe
    rsNFCe_Job!Chave = strChaveNFCe
    rsNFCe_Job!CPF = CPF_CnpjCliente
    rsNFCe_Job!Nome_Consumidor = strNomeCliente
    rsNFCe_Job!Data_Emissao = rsSaida.Fields("DataEmissaoNota").Value
    rsNFCe_Job!Total_Tributos = FormataValorTexto(rsSaida.Fields("TotalNCM").Value, 2)
    rsNFCe_Job!Nome_Emitente = Left(rsFilial.Fields("Razão Social").Value, 60)
    rsNFCe_Job!Endereco_Emitente = Endereco_Emitente
    rsNFCe_Job!IE_Emitente = FormataIE(rsFilial.Fields("Inscrição").Value, False)
    rsNFCe_Job.Update
  End If
End Function

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
    'intDV = 11 - (intResultado Mod 11)
    intDV = 11 - intDV
  End If
  
  CalculaDV = intDV
    
End Function
Private Function GetChaveAcesso(ByVal Filial As Integer, ByVal Sequencia As Long, ByVal TipoMovimento As Integer) As String
  Dim rstNFe As Recordset
  Dim strSQL As String
  Dim strRet As String
  
  strSQL = "SELECT ChaveNFCe "
  strSQL = strSQL & "FROM Saídas "
  strSQL = strSQL & "WHERE Filial = " & Filial & " AND Sequência = " & Sequencia
  Set rstNFe = db.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
  With rstNFe
    If Not (.BOF And .EOF) Then
      strRet = .Fields("ChaveNFCe").Value & ""
    End If
    .Close
  End With
  Set rstNFe = Nothing
  
  'Retorno
  GetChaveAcesso = strRet

End Function

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
Private Function AddLinha(ByRef clcArquivo As Collection, ByRef strLinha As String)
  'Insere linha
  clcArquivo.Add strLinha
  
  sXML_40 = sXML_40 + strLinha
  
  'Limpa linha
  strLinha = ""
End Function

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

Public Function VerificaRetorno(ByVal Sequencia As String) As Boolean
  Dim rsSaida As Recordset
  Set rsSaida = db.OpenRecordset("Select ChaveNFCe from [Saídas] where Filial = " & gnCodFilial & " And Sequência = " & Sequencia & "")
  
  If rsSaida.EOF Then
    MsgBox ("Nenhuma NFCe encontrada para esta venda")
    Exit Function
  End If
  
  If IsNull(rsSaida("ChaveNFCe").Value) Then
    MsgBox ("Nenhuma NFCe encontrada para esta venda")
    Exit Function
  End If
  
  Dim rsFilial As Recordset
  Set rsFilial = db.OpenRecordset("Select BancoPDV from [Parâmetros Filial] Where Filial = " & gnCodFilial & "")

  Dim BancoPDV As Database
  Set BancoPDV = OpenDatabase(rsFilial("BancoPDV").Value & "\QuickStore.mdb", False, False, ";PWD=" & gsGetPValue())
  Dim rsRetorno As Recordset
  Set rsRetorno = BancoPDV.OpenRecordset("Select * from NFCE_ENVI where Chave = '" & rsSaida("ChaveNFCe") & "'")
  
  If rsRetorno.EOF Then
    MsgBox ("NFCe não teve retorno ainda")
    Exit Function
  End If
  
  rsRetorno.MoveLast
  
  Dim ctrl As Control
  Dim Coluna As Field
  For Each ctrl In frmRetornoNFCe.Controls
    For Each Coluna In rsRetorno.Fields
      If ctrl.Name = Coluna.Name And Not IsNull(Coluna.Value) Then
        ctrl.Tag = Coluna.Value
      End If
    Next
  Next
  frmRetornoNFCe.CarregaValores
  frmRetornoNFCe.Show vbModal
End Function
