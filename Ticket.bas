Attribute VB_Name = "modPrintTicket"
Option Explicit

Private rsSaidas As Recordset
Private rsCliFor As Recordset
Private rsFuncionarios As Recordset
Private rsProdutos As Recordset
Private rstClassFiscal As Recordset
Private rsParametros As Recordset
Private rsOp_Saída As Recordset
Private rsCaixas As Recordset
Private rsSaidas_Prod As Recordset
Private rsSaidas_Serv As Recordset
Private rsPesquisa1 As Recordset
Private rsPesquisa2 As Recordset
Private rsPesquisa3 As Recordset
Private rsSaída_Parcelas As Recordset
Private rsSaída_Cheques As Recordset
Private rsCor As Recordset
Private rsTamanho As Recordset
'17/05/2004 - Daniel
'Finalidade de popular o campo Diferimento.ObsDiferimento
Private rstDiferimento As Recordset

'09/01/2004 - Daniel
'Para Finalidade de soma da Qtde da tabela [Saídas - Produtos]
Private rsSomaQtde As Recordset

'30/01/2009 - mpdea
'Implementado "impressão" para email
Public Function Imprime_Ticket(ByVal Nome_Ticket As String, ByVal Filial As Integer, ByVal Sequência As Long, _
  Optional ByVal blnEmail As Boolean = False, Optional ByRef strMessageEmail As String) As Integer
  
  Dim Final As Integer
  Dim Texto As String
  Dim Final_Linha As Integer
  Dim Especial2 As Integer
  Dim Linhas As Integer
  Dim Linha As Integer
  Dim Str_Impre As String
  Dim i As Integer
  Dim Nome_Pesq1 As String
  Dim Nome_Pesq2 As String
  Dim Nome_Pesq3 As String
  Dim Conta_Fat As Integer
  Dim nFileNum As Integer
  Dim nI As Integer
  Dim nComprPag As Integer
  Dim sParte As String
  Dim nCtLin As Integer
  Dim nCtItens As Integer
  Dim Extenso_Tot As String
  Dim sLocalizacao As String
  Dim nCor As Long
  Dim nTamanho As Long
  Dim sNomeCor As String
  Dim sNomeTamanho As String
  Dim sAuxGrade As String
  Dim intVolumagem As Integer
  
  Set rsSaidas = db.OpenRecordset("Saídas")
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcionários", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsOp_Saída = db.OpenRecordset("Operações Saída", , dbReadOnly)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsSaidas_Prod = db.OpenRecordset("Saídas - Produtos", , dbReadOnly)
  Set rsSaidas_Serv = db.OpenRecordset("Saídas - Serviços", , dbReadOnly)
  Set rsPesquisa1 = db.OpenRecordset("Pesquisa 1", , dbReadOnly)
  Set rsPesquisa2 = db.OpenRecordset("Pesquisa 2", , dbReadOnly)
  Set rsPesquisa3 = db.OpenRecordset("Pesquisa 3", , dbReadOnly)
  Set rsSaída_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsSaída_Cheques = db.OpenRecordset("Movimento - Cheques", , dbReadOnly)
  Set rsCor = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsTamanho = db.OpenRecordset("Tamanhos", , dbReadOnly)
  
  
  On Error GoTo ErrHandler
  
  
  '30/01/2009 - mpdea
  If Not blnEmail Then
    SetPrinterName ("TICKET")
    gsInitPrinter = ""
    Call ResetPrinter
  End If
  
  nFileNum = FreeFile
  Open Nome_Ticket For Input As #nFileNum
  
  Input #nFileNum, Texto
  If Left(Texto, 24) <> "*** Configurações Ticket" Then
    gsTitle = LoadResString(201)
    gsMsg = "Layout do cabeçalho do arquivo de configuração """ & Nome_Ticket & """ diferente do esperado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Ticket = 3
    Close #nFileNum
    Exit Function
  End If
  
  nComprPag = 0
  sParte = UCase(Mid(Texto, 75, 3))
  If Len(sParte) > 0 Then
    If sParte <> "NÃO" And sParte <> "LIN" Then
      If Not IsNumeric(sParte) Then
        DisplayMsg "Valor para parâmetro de comprimento da página pode ser: NÃO, LIN ou <99> (inteiro dois digitos)."
        Imprime_Ticket = 3
        Close #nFileNum
        Exit Function
      End If
      If Val(sParte) <= 0 Or Val(sParte) > 20 Then
        DisplayMsg "Comprimento da página em polegadas inválido."
        Imprime_Ticket = 3
        Close #nFileNum
        Exit Function
      End If
      nComprPag = Val(sParte)
    Else
      If sParte = "LIN" Then 'Conte o numero de linhas úteis do doc
        nCtLin = 0
        Do While Not EOF(nFileNum)
          Input #nFileNum, Texto
          If Mid(Texto, 1, 3) <> "***" Then
            nCtLin = nCtLin + 1
          End If
        Loop
        Close #nFileNum
        nFileNum = FreeFile
        Open Nome_Ticket For Input As #nFileNum
        Input #nFileNum, Texto
      End If
    End If
  End If

  '30/01/2009 - mpdea
  If Not blnEmail Then
    If Mid(Texto, 40, 3) = "SIM" Then
      If SetCompressPrinter(Filial) <> 0 Then
        gsTitle = LoadResString(201)
        gsMsg = "Não foi possível usar compressão na impressora solicitada pelo arquivo de configuração: """ & Nome_Ticket & """."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        SetPrinterName ("REL")
        Imprime_Ticket = 3
        Close #nFileNum
        Exit Function
      End If
    End If
  
    If Mid(Texto, 55, 3) = "SIM" Then
      If SetOitavoPrinter(Filial) <> 0 Then
        gsTitle = LoadResString(201)
        gsMsg = "Não foi possível ajustar a impressora para 1/8 solicitada pelo arquivo de configuração: """ & Nome_Ticket & """."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Imprime_Ticket = 3
        SetPrinterName ("REL")
        Close #nFileNum
        Exit Function
      End If
    End If
    
    If sParte = "LIN" Then
      If SetComprimPagLinPrinter(Filial, nCtLin) <> 0 Then
        gsTitle = LoadResString(201)
        gsMsg = "Não foi possível alterar o comprimento de página na impressora."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Imprime_Ticket = 4
        SetPrinterName ("REL")
        Close #nFileNum
        Exit Function
      End If
    Else
      If nComprPag > 0 Then
        If SetComprimPagPrinter(Filial, nComprPag) <> 0 Then
          gsTitle = LoadResString(201)
          gsMsg = "Não foi possível alterar o comprimento de página na impressora."
          gnStyle = vbOKOnly + vbExclamation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          Imprime_Ticket = 4
          SetPrinterName ("REL")
          Close #nFileNum
          Exit Function
        End If
      End If
    End If
    
    Call SetPrinterCommand(gsInitPrinter)
  End If
  
  Rem Acha Saída
  rsSaidas.Index = "Sequência"
  rsSaidas.Seek "=", Filial, Sequência
  If rsSaidas.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Não é possível imprimir o ticket, registro de saída não encontrado para Filial= " & Filial & ", " & "Seqüência= " & Sequência
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 1
    Close #nFileNum
    Exit Function
  End If
  
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Filial
  If rsParametros.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Não é possível imprimir o ticket, parâmetros não encontrados para Filial=" & Filial
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 2
    Close #nFileNum
    Exit Function
  End If
    
    
  '------------------------------------------------------------------------------
  '30/01/2009 - mpdea
  If Not blnEmail Then
    '24/07/2006 - Andrea
    'Se, no parâmetros da filial campo ExigeSenhaGerReimpTicket = True,
    'Verifica se o ticket já foi impresso uma vez e se foi, pede a senha do gerente
    'para liberar a reimpressão.
    If rsParametros.Fields("ExigeSenhaGerReimpTicket").Value Then
      If rsSaidas.Fields("Ticket Impresso").Value Then
        'Senha do gerente
        If Not frmGerente.gbSenhaGerente Then
          Exit Function
        End If
      End If
    End If
  End If
  '------------------------------------------------------------------------------
  
  rsOp_Saída.Index = "Código"
  rsOp_Saída.Seek "=", rsSaidas("Operação")
  If rsOp_Saída.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Operação de Saída referida pelo registro de Saídas não foi localizada: Operação=" & rsSaidas("Operação")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 3
    Close #nFileNum
    Exit Function
  End If

  
  rsCliFor.Index = "Código"
  rsCliFor.Seek "=", rsSaidas("Cliente")
  If rsCliFor.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Cliente referido pelo registro de Saídas não foi localizado: Cliente=" & rsSaidas("Cliente")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 4
    Close #nFileNum
    Exit Function
  End If
  
  
  rsCaixas.Index = "Caixa"
  rsCaixas.Seek "=", rsSaidas("Caixa")
  If rsCaixas.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Caixa referido pelo registro de Saídas não foi localizado: Caixa=" & rsSaidas("Caixa")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 5
    Close #nFileNum
    Exit Function
  End If
  
  
  '26/08/2003 - mpdea
  'Limpa as variáveis públicas
  Call Limpa_Variáveis_Nota
  
  
  GLOB_RESUMO_PAGTO = ""
  If rsSaidas("Recebe - Conta") = True Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor da compra enviado para conta do cliente." + Chr(13)
  End If
  If rsSaidas("Recebe - Dinheiro") <> 0 Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor recebido em dinheiro : " + Format(rsSaidas("Recebe - Dinheiro"), "###,###,##0.00") + Chr(13)
  End If
  If rsSaidas("Recebe - Cartão") <> 0 Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor recebido em cartão : " + Format(rsSaidas("Recebe - Cartão"), "###,###,##0.00") + Chr(13)
  End If
  If rsSaidas("Recebe - Vale") <> 0 Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor recebido em vale : " + Format(rsSaidas("Recebe - Vale"), "###,###,##0.00") + Chr(13)
  End If
  
  Glob_gnCodFilial = Filial
  Glob_Nome_Filial = rsParametros("Nome") & ""
  '----------------------------------------------
  '17/05/2004 - Daniel
  'Adição do campo ObsDiferimento da tabela Diferimento
  Set rstDiferimento = db.OpenRecordset("SELECT ObsDiferimento, EstadoCorrente FROM Diferimento WHERE Filial = " & Filial, dbOpenDynaset)
  
  With rstDiferimento
    If Not (.BOF And .EOF) Then
      '21/06/2004 - Imprimir somente quando o Diferimento.EstadoCorrente é PR
      'e Verificar se o cliente é PR e "J"
      If .Fields("EstadoCorrente") = "PR" Then
        Dim rstCliente As Recordset
        Dim strQuery   As String
        
        strQuery = "SELECT Código, Física_Jurídica, Estado "
        strQuery = strQuery & " FROM Cli_For "
        strQuery = strQuery & " WHERE Código = " & rsSaidas.Fields("Cliente").Value
        
        Set rstCliente = db.OpenRecordset(strQuery, dbOpenDynaset)
        
        With rstCliente
          If Not (.BOF And .EOF) Then
            If Mid((.Fields("Estado").Value), 1, 2) = "PR" And .Fields("Física_Jurídica") = "J" Then
              g_strObsDiferimento = rstDiferimento.Fields("ObsDiferimento").Value & ""
            Else
              g_strObsDiferimento = ""
            End If
          End If
          .Close
        End With
        
        Set rstCliente = Nothing
      
      Else
        g_strObsDiferimento = ""
      End If
      '--------------------------[21/06/2004]---------------------------------
    
    Else
      g_strObsDiferimento = ""
    End If
    .Close
  End With
  
  Set rstDiferimento = Nothing
  '----------------------------------------------
  Glob_Data = Format(rsSaidas("Data"), "dd/mm/yyyy")
  Glob_Data_Saída = Format(Date, "dd/mm/yyyy")
  Glob_Hora_Saída = Format(Time, "hh:mm:ss")
  Glob_Cod_Operação = rsSaidas("Operação")
  Glob_Nome_Operação = rsOp_Saída("Nome") & ""
  Glob_Código_Fiscal = rsOp_Saída("Código Fiscal") & "" '02/10/2006 - Anderson - Correção para impressão de ticket
  Glob_Sequência = rsSaidas("Sequência")
  '----------------------------------------------
  '13/04/2004 - Daniel
  'Populando as vars g_lngNumAutorizacao e g_intMesX
  'Case: STC de Caxias do Sul
  If IsNumeric(rsSaidas.Fields("Num Autorizacao").Value) Then
    g_lngNumAutorizacao = rsSaidas.Fields("Num Autorizacao").Value
  End If
  
  If IsNumeric(rsSaidas.Fields("MesX").Value) Then
    g_intMesX = rsSaidas.Fields("MesX").Value
  End If
  '----------------------------------------------
  Glob_Cod_Vendedor = rsSaidas("Digitador")
  
  '20/05/2005 - Daniel
  '
  'Solicitante: Pedágio - Esta otimização está disponível
  '             para todos usuários do Quick Store
  '
  'O sistema Julga se a nota fiscal foi criada
  'automaticamente ou manualmente a partir da operação escolhida
  'Nota: Caso seja manualmente (notas de bloquinho), o sistema não
  'incrementou o contador pois o sistema estava fora do ar exibirá
  'a nota fiscal que o usuário digitou
  If Not (gbNotaManual(rsSaidas.Fields("Operação").Value, "SAIDA")) Then
    Glob_Nota_Impressa = rsSaidas("Nota Impressa")
  Else
    Glob_Nota_Impressa = CLng("0" & rsSaidas("Nota Fiscal"))
  End If
  
  
  '30/01/2004 - Daniel
  'Populando Percentuais de Impostos em Serviços
  '02/04/2004 - Busca a partir da tabela de Saídas e não mais
  'de Parâmetros
  If Not IsNull(rsSaidas.Fields("Percentual CSLL").Value) Then
    g_dblPercentualCSLL = rsSaidas.Fields("Percentual CSLL").Value
  Else
    g_dblPercentualCSLL = 0
  End If
  
  If Not IsNull(rsSaidas.Fields("Percentual COFINS").Value) Then
    g_dblPercentualCOFINS = rsSaidas.Fields("Percentual COFINS").Value
  Else
    g_dblPercentualCOFINS = 0
  End If
  
  If Not IsNull(rsSaidas.Fields("Percentual PIS").Value) Then
    g_dblPercentualPIS = rsSaidas.Fields("Percentual PIS").Value
  Else
    g_dblPercentualPIS = 0
  End If
  
  If Not IsNull(rsSaidas.Fields("Percentual IRRF").Value) Then
    g_dblPercentualIRRF = rsSaidas.Fields("Percentual IRRF").Value
  Else
    g_dblPercentualIRRF = 0
  End If
  '------------------------------------------------------------

  '27/04/2005 - Daniel
  If IsNumeric(rsSaidas.Fields("Seguro").Value) Then
    g_dblSeguro = Format(rsSaidas.Fields("Seguro").Value, FORMAT_VALUE)
  Else
    g_dblSeguro = 0
  End If

  '15/08/2002 - mpdea
  'Incluído o campo de informações sobre o orçamento (número do orçamento e terminal)
  gstrInfoNrOrcamento = rsSaidas.Fields("InfoNrOrcamento").Value & ""
  
  '----------------------------------------------
  '08/01/2004 - Daniel
  'Incluído vars para o campo Valor Recebido e
  'Troco da tabela de Saídas
  'Populo as variáveis
  'g_dblValorRecebido = rsSaidas.Fields("Valor Recebido").Value
  'g_dblTroco = rsSaidas.Fields("Troco").Value
  'Usando o IsDataType para evitar erros
  Call IsDataType(dtDouble, rsSaidas.Fields("Valor Recebido").Value, g_dblValorRecebido)
  Call IsDataType(dtDouble, rsSaidas.Fields("Troco").Value, g_dblTroco)
  '----------------------------------------------
  
  '----------------------------------------------
  '09/01/2004 - Daniel
  Set rsSomaQtde = db.OpenRecordset("SELECT SUM(Qtde) AS Soma FROM [Saídas - Produtos] WHERE Filial =" & Glob_gnCodFilial & " AND Sequência =" & Glob_Sequência, dbOpenDynaset)
  'g_sngQtdeItens = rsSomaQtde.Fields("Soma")
  Call IsDataType(dtSingle, rsSomaQtde.Fields("Soma"), g_sngQtdeItens)
  '----------------------------------------------
  
  
  rsFuncionarios.Index = "Código"
  rsFuncionarios.Seek "=", Glob_Cod_Vendedor
  If rsFuncionarios.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Digitador referido pelo registro de Saídas não foi localizado: Digitador=" & rsSaidas("Digitador")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 6
    Close #nFileNum
    Exit Function
  End If
  
  '20/11/2006 - Anderson
  'Incluído o campo apelido do funcionário.
  'Solicitante: Technomax
  Glob_Nome_Vendedor = rsFuncionarios("Nome") & ""
  Glob_Apelido = rsFuncionarios("Apelido") & ""
  
  
  '-------------------------------------------------------------------
  '07/08/2003 - mpdea
  'Incluído Código e Nome do Técnico
  Call IsDataType(dtInteger, rsSaidas.Fields("Técnico").Value, g_intCodigoTecnico)
  
  
  '14/08/2003 - Adicionada a cláusula abaixo que verifica se o código do técnico é maior do que ZERO
  If g_intCodigoTecnico > 0 Then
    rsFuncionarios.Seek "=", g_intCodigoTecnico
    If rsFuncionarios.NoMatch Then
      gsTitle = LoadResString(201)
      gsMsg = "Técnico referido pelo registro de Saídas não foi localizado: Técnico = " & g_intCodigoTecnico
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      SetPrinterName ("REL")
      Imprime_Ticket = 7
      Close #nFileNum
      Exit Function
    End If
    g_strNomeTecnico = rsFuncionarios.Fields("Apelido").Value & ""
  End If
  
  '31/05/2007 - Anderson
  'Inclusão de campos no lay-out
  Glob_Prometido = "" & rsSaidas("Prometido Para")
  Glob_Aprovado = "" & rsSaidas("Orçamento Aprovado")
  
  '-------------------------------------------------------------------
  Glob_Código_Cli = rsSaidas("Cliente")
  Glob_Nome = rsCliFor("Nome") & ""
  Glob_Fantasia = rsCliFor("Fantasia") & ""

  Glob_Endereço = rsCliFor("Endereço") & ""
  Glob_NumeroEndereco = rsCliFor.Fields("Endereço Número").Value & "" '23/10/2009 - mpdea
  
  Glob_Complemento = rsCliFor("Complemento") & ""
  Glob_Bairro = rsCliFor("Bairro") & ""
  Glob_CEP = rsCliFor("CEP") & ""
  Glob_Cidade = rsCliFor("Cidade") & ""
  Glob_Estado = rsCliFor("Estado") & ""
  Glob_Fone1 = rsCliFor("Fone 1") & ""
  Glob_Fone2 = rsCliFor("Fone 2") & ""
  Glob_CGC = rsCliFor("CGC") & ""
  Glob_Inscrição = rsCliFor("Inscrição") & ""
  '------------------------------------------------------------
  '13/04/2004 - Daniel
  'Populando as vars g_strEnderecoCob, g_strComplementoCob
  'g_strBairroCob, g_strCidadeCob, g_strEstadoCob e g_strCepCob
  g_strEnderecoCob = rsCliFor("Endereço Cob").Value & ""
  g_strComplementoCob = rsCliFor("Complemento Cob").Value & ""
  g_strBairroCob = rsCliFor("Bairro Cob").Value & ""
  g_strCidadeCob = rsCliFor("Cidade Cob").Value & ""
  g_strEstadoCob = rsCliFor("Estado Cob").Value & ""
  g_strCepCob = rsCliFor("CEP Cob").Value & ""
  '------------------------------------------------------------
  '06/05/2004 - Daniel
  'Adição do campo ObsIsentoIPI da tabela Cli_For
  g_strObsIsentoIPI = rsCliFor("ObsIsentoIPI").Value & ""
  '------------------------------------------------------------
  Glob_Cod_Caixa = rsSaidas("Caixa")
  Glob_Nome_Caixa = rsCaixas("Descrição") & ""
  Glob_Tab_Preço = rsSaidas("Tabela") & ""
  Glob_RefrmInterna = rsSaidas("Referência") & ""
  Glob_Obs_Mov = rsSaidas("Observações") & ""
  
'  For nI = 0 To 7
'    gsObsDoc(nI) = frmObsNota.Obs(nI).Text & ""
'  Next nI
'
'  gsTransportadora = frmObsNota.cboTransp.Text & ""
'  gsPlaca = frmObsNota.Placa.Text & ""
'  gsUfrmPlaca = frmObsNota.UfrmPlaca.Text & ""
'  gsQtdeTrans = frmObsNota.Qtde.Text & ""
'  gsMarcaTrans = frmObsNota.Marca.Text & ""
'  gsEspecieTrans = frmObsNota.Espécie.Text & ""
'  gsPesoBruto = frmObsNota.Bruto.Text & ""
'  gsPesoLiquido = frmObsNota.Líquido.Text & ""
'  If frmObsNota.O_Destinatário.Value = True Then gsFretePago = "2"
'  If frmObsNota.O_Destinatário.Value = False Then gsFretePago = "1"

  Glob_Base_ICM = rsSaidas("Base ICM")
  Glob_Valor_ICM = rsSaidas("Valor ICM")
  Glob_Base_ICM_Sub = rsSaidas("Base ICM Subs")
  Glob_Valor_ICM_Sub = rsSaidas("Valor ICM Subs")
  Glob_Total_Produto = rsSaidas("Produtos")
  Glob_Frete = rsSaidas("Frete")
  Glob_IPI = rsSaidas("IPI")
  Glob_Total_Pagar = rsSaidas("Total")
  
  Glob_Total_Serviço = rsSaidas("Serviços")
  Glob_Total_ISS = rsSaidas("Valor ISS")
  
  '''Glob_Total_Desconto = rsSaidas("Desconto")
  Glob_Total_Desconto = Glob_Total_Produto + Glob_Frete - Glob_Total_Pagar
  
  '30/01/2004 - Daniel
  'Populando Totais de impostos requeridos para Serviços
  '20/04/2004 - Manutenido a lógica
  If g_dblPercentualCSLL <> 0 Then
    g_dblTotalCSLL = Format(rsSaidas("Total CSLL").Value, FORMAT_VALUE)
  Else
    g_dblTotalCSLL = 0
  End If
  
  If g_dblPercentualCOFINS <> 0 Then
    g_dblTotalCOFINS = Format(rsSaidas("Total COFINS").Value, FORMAT_VALUE)
  Else
    g_dblTotalCOFINS = 0
  End If
  
  If g_dblPercentualPIS <> 0 Then
    g_dblTotalPIS = Format(rsSaidas("Total PIS").Value, FORMAT_VALUE)
  Else
    g_dblTotalPIS = 0
  End If
  
  If g_dblPercentualIRRF <> 0 Then
    g_dblTotalIRRF = Format(rsSaidas("Total IRRF").Value, FORMAT_VALUE)
  Else
    g_dblTotalIRRF = 0
  End If
  '----------------------------------------------------------------------------
  
  '06/09/2002 - mpdea
  'Incluído o campo para exibição do Desconto no SubTotal
'''  Call IsDataType(dtDouble, rsSaidas.Fields("DescontoSubTotal").Value, g_dblDescontoSubTotal)
  Call IsDataType(dtDouble, Glob_Total_Desconto, g_dblDescontoSubTotal)
  
  
  '19/08/2003 - mpdea
  'Incluído campo para totalizador de produtos com desconto no subtotal
  g_dblTotalProdutosDST = Glob_Total_Produto - g_dblDescontoSubTotal
  
  
  '26/08/2003 - mpdea
  'Incluído campo para totalizador de produtos menos total de descontos
  g_dblTotalProdutosMenosDescontos = Glob_Total_Produto - Glob_Total_Desconto
  
  
  Extenso_Tot = Extenso(rsSaidas("Total"))
  Extenso_Tot = Extenso_Tot + "                                                                               "
  Extenso_Tot = Extenso_Tot + "                                                                               "
  
  Extenso1_60 = Mid(Extenso_Tot, 1, 60)
  Extenso61_120 = Mid(Extenso_Tot, 61, 60)
  Extenso121_180 = Mid(Extenso_Tot, 121, 60)
  
  Extenso1_45 = Mid(Extenso_Tot, 1, 45)
  Extenso46_90 = Mid(Extenso_Tot, 46, 45)
  Extenso91_135 = Mid(Extenso_Tot, 91, 45)
  Extenso136_180 = Mid(Extenso_Tot, 136, 45)
    
  Extenso1_30 = Mid(Extenso_Tot, 1, 30)
  Extenso31_60 = Mid(Extenso_Tot, 31, 30)
  Extenso61_90 = Mid(Extenso_Tot, 61, 30)
  Extenso91_120 = Mid(Extenso_Tot, 91, 30)
  Extenso121_150 = Mid(Extenso_Tot, 121, 30)
  Extenso151_180 = Mid(Extenso_Tot, 151, 30)
  
  Rem Monta tabela dos produtos
  Erase Tab_Prod
  
  gnCtItemProd = 0
  Linha = 0
  Glob_Conta_Prod = 0
  
  rsPesquisa1.Index = "Código"
  rsPesquisa2.Index = "Código"
  rsPesquisa3.Index = "Código"
  rsProdutos.Index = "Código"
  rsSaidas_Prod.Index = "Sequência"

Lp_Prod:
  rsSaidas_Prod.Seek ">", rsSaidas("Filial"), rsSaidas("Sequência"), Linha
  If rsSaidas_Prod.NoMatch Then GoTo Ve_Serviços
  If rsSaidas("Filial") <> rsSaidas_Prod("Filial") Then GoTo Ve_Serviços
  If rsSaidas("Sequência") <> rsSaidas_Prod("Sequência") Then GoTo Ve_Serviços
  Linha = rsSaidas_Prod("Linha")
  
  Nome_Pesq1 = ""
  Nome_Pesq2 = ""
  Nome_Pesq3 = ""
  
  rsProdutos.Seek "=", rsSaidas_Prod("Código Sem Grade")
  If rsProdutos.NoMatch Then GoTo Lp_Prod
  
  If rsProdutos("Pesquisa 1") <> 0 Then
    rsPesquisa1.Seek "=", rsProdutos("Pesquisa 1")
    If Not rsPesquisa1.NoMatch Then Nome_Pesq1 = rsPesquisa1("Nome")
  End If
  If rsProdutos("Pesquisa 2") <> 0 Then
    rsPesquisa2.Seek "=", rsProdutos("Pesquisa 2")
    If Not rsPesquisa2.NoMatch Then Nome_Pesq2 = rsPesquisa2("Nome")
  End If
  If rsProdutos("Pesquisa 3") <> 0 Then
    rsPesquisa3.Seek "=", rsProdutos("Pesquisa 3")
    If Not rsPesquisa3.NoMatch Then Nome_Pesq3 = rsPesquisa3("Nome")
  End If
  
  
  Tab_Prod(gnCtItemProd).Código = rsProdutos("Código")
  Tab_Prod(gnCtItemProd).Código_Prod_Forn = rsProdutos("Código do Fornecedor") & ""
  
  '25/08/2004 - Daniel
  'Tratamento para impressão da Descrição Adicional
  'no lugar do Nome do Produto
  If rsProdutos("UsaDescrAdic").Value Then
    Tab_Prod(gnCtItemProd).Nome = rsSaidas_Prod("Descricao Adicional") & ""
  Else
    Tab_Prod(gnCtItemProd).Nome = rsProdutos("Nome Nota") & ""
  End If
  
  If rsProdutos("Nome Nota") = "" Then
    '25/08/2004 - Daniel
    'Tratamento para impressão da Descrição Adicional
    If rsProdutos("UsaDescrAdic").Value Then
      Tab_Prod(gnCtItemProd).Nome = rsSaidas_Prod("Descricao Adicional") & ""
    Else
      Tab_Prod(gnCtItemProd).Nome = rsProdutos("Nome") & ""
    End If
  End If
  
  '04/09/2002 - mpdea
  'Incluído os campos para impressão específica do nome do produto como
  'está no campo Nome do cadastro ou o campo Nome para nota (Fixo)
  Tab_Prod(gnCtItemProd).NomeCadastro = rsProdutos("Nome") & ""
  
  '25/08/2004 - Daniel
  'Tratamento para impressão da Descrição Adicional
  If rsProdutos("UsaDescrAdic").Value Then
    Tab_Prod(gnCtItemProd).NomeNota = rsSaidas_Prod("Descricao Adicional") & ""
  Else
    Tab_Prod(gnCtItemProd).NomeNota = rsProdutos("Nome Nota") & ""
  End If
  
  Tab_Prod(gnCtItemProd).C_Fiscal = rsProdutos("Classificação Fiscal") & ""
  
  '11/11/2004 - Daniel
  'Tratamento da impressão da Descrição da Classificação Fiscal
  If IsNumeric(rsProdutos("Classificação Fiscal")) And rsProdutos("Classificação Fiscal") <> 0 Then
    Set rstClassFiscal = db.OpenRecordset("SELECT * FROM [Classificação Fiscal] WHERE Código = " & rsProdutos("Classificação Fiscal"), dbOpenDynaset)
    
    With rstClassFiscal
      If Not (.BOF And .EOF) Then
        .MoveFirst
        '22/09/2005 - mpdea
        'Corrigido descrição da classificação fiscal
        'Estava armazenando somente a do último produto
        Tab_Prod(gnCtItemProd).DescricaoClassificaoFiscal = .Fields("Nome").Value & ""
        'g_strDescrClassFiscal = .Fields("Nome").Value & ""
      End If
      .Close
    End With
    
    Set rstClassFiscal = Nothing
  End If
  
  Tab_Prod(gnCtItemProd).S_Tributária = rsProdutos("Situação Tributária") & ""
  Tab_Prod(gnCtItemProd).Unid = rsProdutos("Unidade Venda") & ""
  
  '27/04/2005 - Daniel
  'Tratamento para Produtos.Fabricante
  Tab_Prod(gnCtItemProd).Fabricante = rsProdutos.Fields("Fabricante").Value & ""
  
  '29/11/2004 - Daniel
  'Incluído os campos Lote e Data de Validade
  If Len(rsProdutos("Lote").Value) > 0 Then Tab_Prod(gnCtItemProd).Lote = rsProdutos.Fields("Lote").Value
  If IsDate(rsProdutos.Fields("DataValidade").Value) Then Tab_Prod(gnCtItemProd).DataValidade = CStr(rsProdutos.Fields("DataValidade").Value)
  
  Tab_Prod(gnCtItemProd).Qtde = rsSaidas_Prod("Qtde")
  '04/05/2004 - Daniel
  'Personalização Embalavi
  If g_bln5CasasDecimais Then
    Tab_Prod(gnCtItemProd).Valor_Unit = Format((rsSaidas_Prod("Preço")), "##,###,##0.00000")
    Tab_Prod(gnCtItemProd).Valor_Total = (Format((rsSaidas_Prod("Preço")), "##,###,##0.00000")) * rsSaidas_Prod("Qtde")
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Tab_Prod(gnCtItemProd).Valor_Unit = Format((rsSaidas_Prod("Preço")), "##,###,##0.000")
    Tab_Prod(gnCtItemProd).Valor_Total = (Format((rsSaidas_Prod("Preço")), "##,###,##0.000")) * rsSaidas_Prod("Qtde")
  Else
    Tab_Prod(gnCtItemProd).Valor_Unit = rsSaidas_Prod("Preço")
    Tab_Prod(gnCtItemProd).Valor_Total = rsSaidas_Prod("Preço") * rsSaidas_Prod("Qtde")
  End If
    
  Tab_Prod(gnCtItemProd).Desconto_Perc = rsSaidas_Prod("Desconto")
  Tab_Prod(gnCtItemProd).Aliq_ICM = rsSaidas_Prod("ICM")
  Tab_Prod(gnCtItemProd).Valor_ICM = 0
  Tab_Prod(gnCtItemProd).Aliq_IPI = rsSaidas_Prod("IPI")
  Tab_Prod(gnCtItemProd).Valor_IPI = 0
  Tab_Prod(gnCtItemProd).Valor_Final = rsSaidas_Prod("Preço Final")
  Tab_Prod(gnCtItemProd).Pesq1 = rsProdutos("Pesquisa 1")
  Tab_Prod(gnCtItemProd).Pesq2 = rsProdutos("Pesquisa 2")
  Tab_Prod(gnCtItemProd).Pesq3 = rsProdutos("Pesquisa 3")
  Tab_Prod(gnCtItemProd).Nome_Pesq1 = Nome_Pesq1
  Tab_Prod(gnCtItemProd).Nome_Pesq2 = Nome_Pesq2
  Tab_Prod(gnCtItemProd).Nome_Pesq3 = Nome_Pesq3
  Tab_Prod(gnCtItemProd).Local = rsProdutos("Localização") & ""
  Tab_Prod(gnCtItemProd).Descr_Adicional = rsSaidas_Prod("Descricao Adicional") & ""
  
  sAuxGrade = ""
  nCor = 0
  nTamanho = 0
  
  sNomeCor = ""
  sNomeTamanho = ""
  If UCase(rsSaidas_Prod("Código")) <> UCase(rsSaidas_Prod("Código sem Grade")) Then
     
     sAuxGrade = Right(rsSaidas_Prod("Código"), 6)
     nCor = Right(sAuxGrade, 3)
     nTamanho = Left(sAuxGrade, 3)
     
     rsCor.Index = "Código"
     rsCor.Seek "=", nCor
     If rsCor.NoMatch Then sNomeCor = ""
     sNomeCor = rsCor("Nome") & ""
     
     rsTamanho.Index = "Código"
     rsTamanho.Seek "=", nTamanho
     If rsTamanho.NoMatch Then sNomeTamanho = ""
     sNomeTamanho = rsTamanho("Nome") & ""
        
     Tab_Prod(gnCtItemProd).Cor = nCor & ""
     Tab_Prod(gnCtItemProd).Nome_Cor = sNomeCor & ""
     Tab_Prod(gnCtItemProd).Tamanho = nTamanho & ""
     Tab_Prod(gnCtItemProd).Nome_Tamanho = sNomeTamanho & ""
  End If
  
 
  '27/09/2004 - mpdea
  'Incluído o campo de Volumagem por Quantidade
  With Tab_Prod(gnCtItemProd)
    Call IsDataType(dtInteger, rsProdutos.Fields("Volumagem").Value, intVolumagem)
    If intVolumagem > 0 Then
      .VolumagemQtde = "(" & Format(.Qtde \ intVolumagem, "000") & "/" & Format(.Qtde Mod intVolumagem, "000") & ")"
    End If
  End With
  
   
  gnCtItemProd = gnCtItemProd + 1
  
  GoTo Lp_Prod
  
Ve_Serviços:
  Erase Tab_Serv
  Linha = 0
  Glob_Conta_Serv = 0
  gnCtItemServ = 0
  
  rsSaidas_Serv.Index = "Sequência"
Lp_Serv:
  rsSaidas_Serv.Seek ">", rsSaidas("Filial"), rsSaidas("Sequência"), Linha
  If rsSaidas_Serv.NoMatch Then GoTo Ve_Cheque
  If rsSaidas("Filial") <> rsSaidas_Serv("Filial") Then GoTo Ve_Cheque
  If rsSaidas("Sequência") <> rsSaidas_Serv("Sequência") Then GoTo Ve_Cheque
  Linha = rsSaidas_Serv("Linha")

  Tab_Serv(gnCtItemServ).Código = rsSaidas_Serv("Código")
  If rsSaidas_Serv("Completo") = True Then Tab_Serv(gnCtItemServ).Concluído = "Sim"
  If rsSaidas_Serv("Completo") = False Then Tab_Serv(gnCtItemServ).Concluído = "Não"
  Tab_Serv(gnCtItemServ).Descrição = rsSaidas_Serv("Descrição") & ""
  Tab_Serv(gnCtItemServ).Preço_Unit = rsSaidas_Serv("Preço")
  Tab_Serv(gnCtItemServ).Qtde = rsSaidas_Serv("Tempo") & ""
 ' Tab_Serv(gnCtItemServ).Preço_Total = Format(CStr(Tab_Serv(gnCtItemServ).Preço_Unit * Tab_Serv(gnCtItemServ).Qtde), "##############0,00")
  Tab_Serv(gnCtItemServ).Preço_Total = Format(CStr(Tab_Serv(gnCtItemServ).Preço_Unit * Tab_Serv(gnCtItemServ).Qtde))
  '27/07/2005 - Daniel
  'CST (Código de Situação Tributária)
  'Finalidade: Atender a realidade da empresa W.V. Hidroanálise Ltda (J.R. Hidroquímica)
  Tab_Serv(gnCtItemServ).CST = rsSaidas_Serv("CST").Value & ""
  
  gnCtItemServ = gnCtItemServ + 1
  
  GoTo Lp_Serv


Ve_Cheque:

  Conta_Fat = 0
  Linhas = 0
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + " " + Chr(13)
  
  rsSaída_Cheques.Index = "Ordem"
Lp_Cheque:
  rsSaída_Cheques.Seek ">", gnCodFilial, rsSaidas("Sequência"), Linhas
  If rsSaída_Cheques.NoMatch Then GoTo Ve_Fatura
  If rsSaída_Cheques("Filial") <> gnCodFilial Then GoTo Ve_Fatura
  If rsSaída_Cheques("Sequência") <> rsSaidas("Sequência") Then GoTo Ve_Fatura
  Linhas = rsSaída_Cheques("Ordem")
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Recebimento em cheque : " + Format(rsSaída_Cheques("Valor"), "###,###,##0.00")
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "  para " + Format(rsSaída_Cheques("Bom"), "dd/mm/yyyy") + Chr(13)

  Conta_Fat = Conta_Fat + 1
  GoTo Lp_Cheque


Ve_Fatura:
  Erase Tab_Fat
  
  gnCtParcFat = 0
  Linhas = 0
  
  Glob_Conta_Fat = 0
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + " " + Chr(13)
  
  
  rsSaída_Parcelas.Index = "Ordem"
Lp_Fat:
  rsSaída_Parcelas.Seek ">", gnCodFilial, rsSaidas("Sequência"), Linhas
  If rsSaída_Parcelas.NoMatch Then GoTo Final
  If rsSaída_Parcelas("Filial") <> gnCodFilial Then GoTo Final
  If rsSaída_Parcelas("Sequência") <> rsSaidas("Sequência") Then GoTo Final
  Linhas = rsSaída_Parcelas("Ordem")
  '20/05/2005 - Daniel
  '
  'Solicitante: Pedágio - Esta otimização está disponível
  '             para todos usuários do Quick Store
  '             Tratamento para impressão de nota gerada manualmente
  If Not (gbNotaManual(rsSaidas.Fields("Operação").Value, "SAIDA")) Then
    Tab_Fat(gnCtParcFat).Número = LTrim(str(rsSaidas("Nota Impressa"))) + "/" + LTrim(str((gnCtParcFat + 1)))
  Else
    Tab_Fat(gnCtParcFat).Número = LTrim(str(rsSaidas("Nota Fiscal"))) + "/" + LTrim(str((gnCtParcFat + 1)))
  End If
  '
    Tab_Fat(gnCtParcFat).Valor = rsSaída_Parcelas("Valor")
    Tab_Fat(gnCtParcFat).Vencimento = rsSaída_Parcelas("Bom")
  gnCtParcFat = gnCtParcFat + 1
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Parcelamento : " + Format(rsSaída_Parcelas("Valor"), "###,###,##0.00")
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "  para " + Format(rsSaída_Parcelas("Bom"), "dd/mm/yyyy") + Chr(13)
  
  GoTo Lp_Fat


Final:
  Final = False
  Dim TextoAnt As String
  Do
    DoEvents
    Input #nFileNum, Texto
    If Texto = "*** Fim de arquivo ***" Then Final = True
    If Final = False Then
      Texto = Apaga_Aspas(Texto)
      Final_Linha = False
      If Len(Texto) < 3 Then
        MsgBox "Arquivo de configuração inválido."
        Imprime_Ticket = 1
        Exit Function
      End If
      Especial2 = False
      If Left(Texto, 13) = "[LINHA_BRANCO" Then
        Especial2 = True
        Linhas = Val(Mid(Texto, 15))
        Do
          '30/01/2009 - mpdea
          If blnEmail Then
            strMessageEmail = strMessageEmail & vbCrLf
          Else
            Printer.Print
          End If
          Linhas = Linhas - 1
        Loop Until Linhas = 0
      End If
      If Especial2 = False Then
        Str_Impre = Retorna_Texto(Texto)
        'Não imprime linha só com -
        If Trim(Str_Impre) <> "-" Then
          
          '30/01/2009 - mpdea
          If Not blnEmail Then
            '16/08/2002 - mpdea
            'Incluído início da formatação em negrito
            If InStr(Texto, "LINHA_EM_NEGRITO") > 0 Then
              Printer.FontBold = True
            End If
          End If
          
          '30/01/2009 - mpdea
          If blnEmail Then
            strMessageEmail = strMessageEmail & Str_Impre & vbCrLf
          Else
            Printer.Print Str_Impre
          End If
          
          '30/01/2009 - mpdea
          If Not blnEmail Then
            '16/08/2002 - mpdea
            'Término da formatação em negrito
            If InStr(Texto, "LINHA_EM_NEGRITO") > 0 Then
              Printer.FontBold = False
            End If
          End If
          
        End If
      End If
    End If
    TextoAnt = Texto
  Loop Until Final = True
       
  If Glob_Conta_Prod > 0 Then
    nCtItens = 0
    For nI = 0 To 500
      If Len(Trim(Tab_Prod(nI).Código)) > 0 Then
        nCtItens = nCtItens + 1
      End If
    Next nI
    If gnCtItemProd > nCtItens Then
      DisplayMsg "AVISO: Número de itens de produtos existentes excedeu a quantidade de itens definida no lay-out do documento..."
    End If
  End If
  
  If Glob_Conta_Serv > 0 Then
    nCtItens = 0
    For nI = 0 To 50
      If Len(Trim(Tab_Serv(nI).Código)) > 0 Then
        nCtItens = nCtItens + 1
      End If
    Next nI
    If gnCtItemServ > nCtItens Then
      DisplayMsg "AVISO: Número de itens de serviços existentes excedeu a quantidade de itens definida no lay-out do documento..."
    End If
  End If
  
  If Glob_Conta_Fat > 0 Then
    nCtItens = 0
    For nI = 0 To 50
      If Len(Trim(Tab_Fat(nI).Número)) > 0 Then
        nCtItens = nCtItens + 1
      End If
    Next nI
    If gnCtParcFat > nCtItens Then
      DisplayMsg "AVISO: Número de parcelas de fatura existentes excedeu a quantidade de itens definida no lay-out do documento..."
    End If
  End If
  
  '------------------------------------------------------------------------------------
  '30/01/2009 - mpdea
  If Not blnEmail Then
    '24/07/2006 - Andrea
    'Gravação do campo Ticket Impresso na tabela de saídas.
    rsSaidas.Edit
    rsSaidas.Fields("Ticket Impresso") = "-1"
    rsSaidas.Update
  End If
  '-------------------------------------------------------------------------------------
  
  Imprime_Ticket = 0
  
  Close #nFileNum
  
  Printer.FontName = "Lucida Console"
  Printer.FontSize = 7
  
  '30/01/2009 - mpdea
  If Not blnEmail Then
    'Printer.FontName = "Lucida Console"
    'Printer.FontSize = 2
    Printer.Print
    Printer.EndDoc
  End If
  
  SetPrinterName ("REL")
  
  rsSaidas.Close
  Set rsSaidas = Nothing
  rsCliFor.Close
  Set rsCliFor = Nothing
  rsFuncionarios.Close
  Set rsFuncionarios = Nothing
  rsProdutos.Close
  Set rsProdutos = Nothing
  rsParametros.Close
  Set rsParametros = Nothing
  rsOp_Saída.Close
  Set rsOp_Saída = Nothing
  rsCaixas.Close
  Set rsCaixas = Nothing
  rsSaidas_Prod.Close
  Set rsSaidas_Prod = Nothing
  rsSaidas_Serv.Close
  Set rsSaidas_Serv = Nothing
  rsPesquisa1.Close
  Set rsPesquisa1 = Nothing
  rsPesquisa2.Close
  Set rsPesquisa2 = Nothing
  rsPesquisa3.Close
  Set rsPesquisa3 = Nothing
  rsSaída_Parcelas.Close
  Set rsSaída_Parcelas = Nothing
  rsSaída_Cheques.Close
  Set rsSaída_Cheques = Nothing
  rsCor.Close
  Set rsCor = Nothing
  rsTamanho.Close
  Set rsTamanho = Nothing
  '09/01/2004 - Daniel
  rsSomaQtde.Close
  Set rsSomaQtde = Nothing
  
  Exit Function
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao imprimir Nota usando o arquivo de configuração: """ & Nome_Ticket & """."
  If Err.Number = 53 Then
    gsMsg = gsMsg & vbCrLf & "Arquivo de configuração não encontrado."
  Else
    gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  End If
  
  '09/01/2004 - mpdea
  'Adicionado a exibição da mensagem de erro
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  SetPrinterName ("REL")
  On Error Resume Next
  Close #nFileNum
  Exit Function
  
End Function
