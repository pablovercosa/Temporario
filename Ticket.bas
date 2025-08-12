Attribute VB_Name = "modPrintTicket"
Option Explicit

Private rsSaidas As Recordset
Private rsCliFor As Recordset
Private rsFuncionarios As Recordset
Private rsProdutos As Recordset
Private rstClassFiscal As Recordset
Private rsParametros As Recordset
Private rsOp_Sa�da As Recordset
Private rsCaixas As Recordset
Private rsSaidas_Prod As Recordset
Private rsSaidas_Serv As Recordset
Private rsPesquisa1 As Recordset
Private rsPesquisa2 As Recordset
Private rsPesquisa3 As Recordset
Private rsSa�da_Parcelas As Recordset
Private rsSa�da_Cheques As Recordset
Private rsCor As Recordset
Private rsTamanho As Recordset
'17/05/2004 - Daniel
'Finalidade de popular o campo Diferimento.ObsDiferimento
Private rstDiferimento As Recordset

'09/01/2004 - Daniel
'Para Finalidade de soma da Qtde da tabela [Sa�das - Produtos]
Private rsSomaQtde As Recordset

'30/01/2009 - mpdea
'Implementado "impress�o" para email
Public Function Imprime_Ticket(ByVal Nome_Ticket As String, ByVal Filial As Integer, ByVal Sequ�ncia As Long, _
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
  
  Set rsSaidas = db.OpenRecordset("Sa�das")
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  Set rsOp_Sa�da = db.OpenRecordset("Opera��es Sa�da", , dbReadOnly)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsSaidas_Prod = db.OpenRecordset("Sa�das - Produtos", , dbReadOnly)
  Set rsSaidas_Serv = db.OpenRecordset("Sa�das - Servi�os", , dbReadOnly)
  Set rsPesquisa1 = db.OpenRecordset("Pesquisa 1", , dbReadOnly)
  Set rsPesquisa2 = db.OpenRecordset("Pesquisa 2", , dbReadOnly)
  Set rsPesquisa3 = db.OpenRecordset("Pesquisa 3", , dbReadOnly)
  Set rsSa�da_Parcelas = db.OpenRecordset("Movimento - Parcelas", , dbReadOnly)
  Set rsSa�da_Cheques = db.OpenRecordset("Movimento - Cheques", , dbReadOnly)
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
  If Left(Texto, 24) <> "*** Configura��es Ticket" Then
    gsTitle = LoadResString(201)
    gsMsg = "Layout do cabe�alho do arquivo de configura��o """ & Nome_Ticket & """ diferente do esperado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Ticket = 3
    Close #nFileNum
    Exit Function
  End If
  
  nComprPag = 0
  sParte = UCase(Mid(Texto, 75, 3))
  If Len(sParte) > 0 Then
    If sParte <> "N�O" And sParte <> "LIN" Then
      If Not IsNumeric(sParte) Then
        DisplayMsg "Valor para par�metro de comprimento da p�gina pode ser: N�O, LIN ou <99> (inteiro dois digitos)."
        Imprime_Ticket = 3
        Close #nFileNum
        Exit Function
      End If
      If Val(sParte) <= 0 Or Val(sParte) > 20 Then
        DisplayMsg "Comprimento da p�gina em polegadas inv�lido."
        Imprime_Ticket = 3
        Close #nFileNum
        Exit Function
      End If
      nComprPag = Val(sParte)
    Else
      If sParte = "LIN" Then 'Conte o numero de linhas �teis do doc
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
        gsMsg = "N�o foi poss�vel usar compress�o na impressora solicitada pelo arquivo de configura��o: """ & Nome_Ticket & """."
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
        gsMsg = "N�o foi poss�vel ajustar a impressora para 1/8 solicitada pelo arquivo de configura��o: """ & Nome_Ticket & """."
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
        gsMsg = "N�o foi poss�vel alterar o comprimento de p�gina na impressora."
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
          gsMsg = "N�o foi poss�vel alterar o comprimento de p�gina na impressora."
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
  
  Rem Acha Sa�da
  rsSaidas.Index = "Sequ�ncia"
  rsSaidas.Seek "=", Filial, Sequ�ncia
  If rsSaidas.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "N�o � poss�vel imprimir o ticket, registro de sa�da n�o encontrado para Filial= " & Filial & ", " & "Seq��ncia= " & Sequ�ncia
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
    gsMsg = "N�o � poss�vel imprimir o ticket, par�metros n�o encontrados para Filial=" & Filial
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
    'Se, no par�metros da filial campo ExigeSenhaGerReimpTicket = True,
    'Verifica se o ticket j� foi impresso uma vez e se foi, pede a senha do gerente
    'para liberar a reimpress�o.
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
  
  rsOp_Sa�da.Index = "C�digo"
  rsOp_Sa�da.Seek "=", rsSaidas("Opera��o")
  If rsOp_Sa�da.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Opera��o de Sa�da referida pelo registro de Sa�das n�o foi localizada: Opera��o=" & rsSaidas("Opera��o")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 3
    Close #nFileNum
    Exit Function
  End If

  
  rsCliFor.Index = "C�digo"
  rsCliFor.Seek "=", rsSaidas("Cliente")
  If rsCliFor.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Cliente referido pelo registro de Sa�das n�o foi localizado: Cliente=" & rsSaidas("Cliente")
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
    gsMsg = "Caixa referido pelo registro de Sa�das n�o foi localizado: Caixa=" & rsSaidas("Caixa")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 5
    Close #nFileNum
    Exit Function
  End If
  
  
  '26/08/2003 - mpdea
  'Limpa as vari�veis p�blicas
  Call Limpa_Vari�veis_Nota
  
  
  GLOB_RESUMO_PAGTO = ""
  If rsSaidas("Recebe - Conta") = True Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor da compra enviado para conta do cliente." + Chr(13)
  End If
  If rsSaidas("Recebe - Dinheiro") <> 0 Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor recebido em dinheiro : " + Format(rsSaidas("Recebe - Dinheiro"), "###,###,##0.00") + Chr(13)
  End If
  If rsSaidas("Recebe - Cart�o") <> 0 Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor recebido em cart�o : " + Format(rsSaidas("Recebe - Cart�o"), "###,###,##0.00") + Chr(13)
  End If
  If rsSaidas("Recebe - Vale") <> 0 Then
    GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Valor recebido em vale : " + Format(rsSaidas("Recebe - Vale"), "###,###,##0.00") + Chr(13)
  End If
  
  Glob_gnCodFilial = Filial
  Glob_Nome_Filial = rsParametros("Nome") & ""
  '----------------------------------------------
  '17/05/2004 - Daniel
  'Adi��o do campo ObsDiferimento da tabela Diferimento
  Set rstDiferimento = db.OpenRecordset("SELECT ObsDiferimento, EstadoCorrente FROM Diferimento WHERE Filial = " & Filial, dbOpenDynaset)
  
  With rstDiferimento
    If Not (.BOF And .EOF) Then
      '21/06/2004 - Imprimir somente quando o Diferimento.EstadoCorrente � PR
      'e Verificar se o cliente � PR e "J"
      If .Fields("EstadoCorrente") = "PR" Then
        Dim rstCliente As Recordset
        Dim strQuery   As String
        
        strQuery = "SELECT C�digo, F�sica_Jur�dica, Estado "
        strQuery = strQuery & " FROM Cli_For "
        strQuery = strQuery & " WHERE C�digo = " & rsSaidas.Fields("Cliente").Value
        
        Set rstCliente = db.OpenRecordset(strQuery, dbOpenDynaset)
        
        With rstCliente
          If Not (.BOF And .EOF) Then
            If Mid((.Fields("Estado").Value), 1, 2) = "PR" And .Fields("F�sica_Jur�dica") = "J" Then
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
  Glob_Data_Sa�da = Format(Date, "dd/mm/yyyy")
  Glob_Hora_Sa�da = Format(Time, "hh:mm:ss")
  Glob_Cod_Opera��o = rsSaidas("Opera��o")
  Glob_Nome_Opera��o = rsOp_Sa�da("Nome") & ""
  Glob_C�digo_Fiscal = rsOp_Sa�da("C�digo Fiscal") & "" '02/10/2006 - Anderson - Corre��o para impress�o de ticket
  Glob_Sequ�ncia = rsSaidas("Sequ�ncia")
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
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '
  'O sistema Julga se a nota fiscal foi criada
  'automaticamente ou manualmente a partir da opera��o escolhida
  'Nota: Caso seja manualmente (notas de bloquinho), o sistema n�o
  'incrementou o contador pois o sistema estava fora do ar exibir�
  'a nota fiscal que o usu�rio digitou
  If Not (gbNotaManual(rsSaidas.Fields("Opera��o").Value, "SAIDA")) Then
    Glob_Nota_Impressa = rsSaidas("Nota Impressa")
  Else
    Glob_Nota_Impressa = CLng("0" & rsSaidas("Nota Fiscal"))
  End If
  
  
  '30/01/2004 - Daniel
  'Populando Percentuais de Impostos em Servi�os
  '02/04/2004 - Busca a partir da tabela de Sa�das e n�o mais
  'de Par�metros
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
  'Inclu�do o campo de informa��es sobre o or�amento (n�mero do or�amento e terminal)
  gstrInfoNrOrcamento = rsSaidas.Fields("InfoNrOrcamento").Value & ""
  
  '----------------------------------------------
  '08/01/2004 - Daniel
  'Inclu�do vars para o campo Valor Recebido e
  'Troco da tabela de Sa�das
  'Populo as vari�veis
  'g_dblValorRecebido = rsSaidas.Fields("Valor Recebido").Value
  'g_dblTroco = rsSaidas.Fields("Troco").Value
  'Usando o IsDataType para evitar erros
  Call IsDataType(dtDouble, rsSaidas.Fields("Valor Recebido").Value, g_dblValorRecebido)
  Call IsDataType(dtDouble, rsSaidas.Fields("Troco").Value, g_dblTroco)
  '----------------------------------------------
  
  '----------------------------------------------
  '09/01/2004 - Daniel
  Set rsSomaQtde = db.OpenRecordset("SELECT SUM(Qtde) AS Soma FROM [Sa�das - Produtos] WHERE Filial =" & Glob_gnCodFilial & " AND Sequ�ncia =" & Glob_Sequ�ncia, dbOpenDynaset)
  'g_sngQtdeItens = rsSomaQtde.Fields("Soma")
  Call IsDataType(dtSingle, rsSomaQtde.Fields("Soma"), g_sngQtdeItens)
  '----------------------------------------------
  
  
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", Glob_Cod_Vendedor
  If rsFuncionarios.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Digitador referido pelo registro de Sa�das n�o foi localizado: Digitador=" & rsSaidas("Digitador")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Ticket = 6
    Close #nFileNum
    Exit Function
  End If
  
  '20/11/2006 - Anderson
  'Inclu�do o campo apelido do funcion�rio.
  'Solicitante: Technomax
  Glob_Nome_Vendedor = rsFuncionarios("Nome") & ""
  Glob_Apelido = rsFuncionarios("Apelido") & ""
  
  
  '-------------------------------------------------------------------
  '07/08/2003 - mpdea
  'Inclu�do C�digo e Nome do T�cnico
  Call IsDataType(dtInteger, rsSaidas.Fields("T�cnico").Value, g_intCodigoTecnico)
  
  
  '14/08/2003 - Adicionada a cl�usula abaixo que verifica se o c�digo do t�cnico � maior do que ZERO
  If g_intCodigoTecnico > 0 Then
    rsFuncionarios.Seek "=", g_intCodigoTecnico
    If rsFuncionarios.NoMatch Then
      gsTitle = LoadResString(201)
      gsMsg = "T�cnico referido pelo registro de Sa�das n�o foi localizado: T�cnico = " & g_intCodigoTecnico
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
  'Inclus�o de campos no lay-out
  Glob_Prometido = "" & rsSaidas("Prometido Para")
  Glob_Aprovado = "" & rsSaidas("Or�amento Aprovado")
  
  '-------------------------------------------------------------------
  Glob_C�digo_Cli = rsSaidas("Cliente")
  Glob_Nome = rsCliFor("Nome") & ""
  Glob_Fantasia = rsCliFor("Fantasia") & ""

  Glob_Endere�o = rsCliFor("Endere�o") & ""
  Glob_NumeroEndereco = rsCliFor.Fields("Endere�o N�mero").Value & "" '23/10/2009 - mpdea
  
  Glob_Complemento = rsCliFor("Complemento") & ""
  Glob_Bairro = rsCliFor("Bairro") & ""
  Glob_CEP = rsCliFor("CEP") & ""
  Glob_Cidade = rsCliFor("Cidade") & ""
  Glob_Estado = rsCliFor("Estado") & ""
  Glob_Fone1 = rsCliFor("Fone 1") & ""
  Glob_Fone2 = rsCliFor("Fone 2") & ""
  Glob_CGC = rsCliFor("CGC") & ""
  Glob_Inscri��o = rsCliFor("Inscri��o") & ""
  '------------------------------------------------------------
  '13/04/2004 - Daniel
  'Populando as vars g_strEnderecoCob, g_strComplementoCob
  'g_strBairroCob, g_strCidadeCob, g_strEstadoCob e g_strCepCob
  g_strEnderecoCob = rsCliFor("Endere�o Cob").Value & ""
  g_strComplementoCob = rsCliFor("Complemento Cob").Value & ""
  g_strBairroCob = rsCliFor("Bairro Cob").Value & ""
  g_strCidadeCob = rsCliFor("Cidade Cob").Value & ""
  g_strEstadoCob = rsCliFor("Estado Cob").Value & ""
  g_strCepCob = rsCliFor("CEP Cob").Value & ""
  '------------------------------------------------------------
  '06/05/2004 - Daniel
  'Adi��o do campo ObsIsentoIPI da tabela Cli_For
  g_strObsIsentoIPI = rsCliFor("ObsIsentoIPI").Value & ""
  '------------------------------------------------------------
  Glob_Cod_Caixa = rsSaidas("Caixa")
  Glob_Nome_Caixa = rsCaixas("Descri��o") & ""
  Glob_Tab_Pre�o = rsSaidas("Tabela") & ""
  Glob_RefrmInterna = rsSaidas("Refer�ncia") & ""
  Glob_Obs_Mov = rsSaidas("Observa��es") & ""
  
'  For nI = 0 To 7
'    gsObsDoc(nI) = frmObsNota.Obs(nI).Text & ""
'  Next nI
'
'  gsTransportadora = frmObsNota.cboTransp.Text & ""
'  gsPlaca = frmObsNota.Placa.Text & ""
'  gsUfrmPlaca = frmObsNota.UfrmPlaca.Text & ""
'  gsQtdeTrans = frmObsNota.Qtde.Text & ""
'  gsMarcaTrans = frmObsNota.Marca.Text & ""
'  gsEspecieTrans = frmObsNota.Esp�cie.Text & ""
'  gsPesoBruto = frmObsNota.Bruto.Text & ""
'  gsPesoLiquido = frmObsNota.L�quido.Text & ""
'  If frmObsNota.O_Destinat�rio.Value = True Then gsFretePago = "2"
'  If frmObsNota.O_Destinat�rio.Value = False Then gsFretePago = "1"

  Glob_Base_ICM = rsSaidas("Base ICM")
  Glob_Valor_ICM = rsSaidas("Valor ICM")
  Glob_Base_ICM_Sub = rsSaidas("Base ICM Subs")
  Glob_Valor_ICM_Sub = rsSaidas("Valor ICM Subs")
  Glob_Total_Produto = rsSaidas("Produtos")
  Glob_Frete = rsSaidas("Frete")
  Glob_IPI = rsSaidas("IPI")
  Glob_Total_Pagar = rsSaidas("Total")
  
  Glob_Total_Servi�o = rsSaidas("Servi�os")
  Glob_Total_ISS = rsSaidas("Valor ISS")
  
  '''Glob_Total_Desconto = rsSaidas("Desconto")
  Glob_Total_Desconto = Glob_Total_Produto + Glob_Frete - Glob_Total_Pagar
  
  '30/01/2004 - Daniel
  'Populando Totais de impostos requeridos para Servi�os
  '20/04/2004 - Manutenido a l�gica
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
  'Inclu�do o campo para exibi��o do Desconto no SubTotal
'''  Call IsDataType(dtDouble, rsSaidas.Fields("DescontoSubTotal").Value, g_dblDescontoSubTotal)
  Call IsDataType(dtDouble, Glob_Total_Desconto, g_dblDescontoSubTotal)
  
  
  '19/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos com desconto no subtotal
  g_dblTotalProdutosDST = Glob_Total_Produto - g_dblDescontoSubTotal
  
  
  '26/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos menos total de descontos
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
  
  rsPesquisa1.Index = "C�digo"
  rsPesquisa2.Index = "C�digo"
  rsPesquisa3.Index = "C�digo"
  rsProdutos.Index = "C�digo"
  rsSaidas_Prod.Index = "Sequ�ncia"

Lp_Prod:
  rsSaidas_Prod.Seek ">", rsSaidas("Filial"), rsSaidas("Sequ�ncia"), Linha
  If rsSaidas_Prod.NoMatch Then GoTo Ve_Servi�os
  If rsSaidas("Filial") <> rsSaidas_Prod("Filial") Then GoTo Ve_Servi�os
  If rsSaidas("Sequ�ncia") <> rsSaidas_Prod("Sequ�ncia") Then GoTo Ve_Servi�os
  Linha = rsSaidas_Prod("Linha")
  
  Nome_Pesq1 = ""
  Nome_Pesq2 = ""
  Nome_Pesq3 = ""
  
  rsProdutos.Seek "=", rsSaidas_Prod("C�digo Sem Grade")
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
  
  
  Tab_Prod(gnCtItemProd).C�digo = rsProdutos("C�digo")
  Tab_Prod(gnCtItemProd).C�digo_Prod_Forn = rsProdutos("C�digo do Fornecedor") & ""
  
  '25/08/2004 - Daniel
  'Tratamento para impress�o da Descri��o Adicional
  'no lugar do Nome do Produto
  If rsProdutos("UsaDescrAdic").Value Then
    Tab_Prod(gnCtItemProd).Nome = rsSaidas_Prod("Descricao Adicional") & ""
  Else
    Tab_Prod(gnCtItemProd).Nome = rsProdutos("Nome Nota") & ""
  End If
  
  If rsProdutos("Nome Nota") = "" Then
    '25/08/2004 - Daniel
    'Tratamento para impress�o da Descri��o Adicional
    If rsProdutos("UsaDescrAdic").Value Then
      Tab_Prod(gnCtItemProd).Nome = rsSaidas_Prod("Descricao Adicional") & ""
    Else
      Tab_Prod(gnCtItemProd).Nome = rsProdutos("Nome") & ""
    End If
  End If
  
  '04/09/2002 - mpdea
  'Inclu�do os campos para impress�o espec�fica do nome do produto como
  'est� no campo Nome do cadastro ou o campo Nome para nota (Fixo)
  Tab_Prod(gnCtItemProd).NomeCadastro = rsProdutos("Nome") & ""
  
  '25/08/2004 - Daniel
  'Tratamento para impress�o da Descri��o Adicional
  If rsProdutos("UsaDescrAdic").Value Then
    Tab_Prod(gnCtItemProd).NomeNota = rsSaidas_Prod("Descricao Adicional") & ""
  Else
    Tab_Prod(gnCtItemProd).NomeNota = rsProdutos("Nome Nota") & ""
  End If
  
  Tab_Prod(gnCtItemProd).C_Fiscal = rsProdutos("Classifica��o Fiscal") & ""
  
  '11/11/2004 - Daniel
  'Tratamento da impress�o da Descri��o da Classifica��o Fiscal
  If IsNumeric(rsProdutos("Classifica��o Fiscal")) And rsProdutos("Classifica��o Fiscal") <> 0 Then
    Set rstClassFiscal = db.OpenRecordset("SELECT * FROM [Classifica��o Fiscal] WHERE C�digo = " & rsProdutos("Classifica��o Fiscal"), dbOpenDynaset)
    
    With rstClassFiscal
      If Not (.BOF And .EOF) Then
        .MoveFirst
        '22/09/2005 - mpdea
        'Corrigido descri��o da classifica��o fiscal
        'Estava armazenando somente a do �ltimo produto
        Tab_Prod(gnCtItemProd).DescricaoClassificaoFiscal = .Fields("Nome").Value & ""
        'g_strDescrClassFiscal = .Fields("Nome").Value & ""
      End If
      .Close
    End With
    
    Set rstClassFiscal = Nothing
  End If
  
  Tab_Prod(gnCtItemProd).S_Tribut�ria = rsProdutos("Situa��o Tribut�ria") & ""
  Tab_Prod(gnCtItemProd).Unid = rsProdutos("Unidade Venda") & ""
  
  '27/04/2005 - Daniel
  'Tratamento para Produtos.Fabricante
  Tab_Prod(gnCtItemProd).Fabricante = rsProdutos.Fields("Fabricante").Value & ""
  
  '29/11/2004 - Daniel
  'Inclu�do os campos Lote e Data de Validade
  If Len(rsProdutos("Lote").Value) > 0 Then Tab_Prod(gnCtItemProd).Lote = rsProdutos.Fields("Lote").Value
  If IsDate(rsProdutos.Fields("DataValidade").Value) Then Tab_Prod(gnCtItemProd).DataValidade = CStr(rsProdutos.Fields("DataValidade").Value)
  
  Tab_Prod(gnCtItemProd).Qtde = rsSaidas_Prod("Qtde")
  '04/05/2004 - Daniel
  'Personaliza��o Embalavi
  If g_bln5CasasDecimais Then
    Tab_Prod(gnCtItemProd).Valor_Unit = Format((rsSaidas_Prod("Pre�o")), "##,###,##0.00000")
    Tab_Prod(gnCtItemProd).Valor_Total = (Format((rsSaidas_Prod("Pre�o")), "##,###,##0.00000")) * rsSaidas_Prod("Qtde")
  '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Tab_Prod(gnCtItemProd).Valor_Unit = Format((rsSaidas_Prod("Pre�o")), "##,###,##0.000")
    Tab_Prod(gnCtItemProd).Valor_Total = (Format((rsSaidas_Prod("Pre�o")), "##,###,##0.000")) * rsSaidas_Prod("Qtde")
  Else
    Tab_Prod(gnCtItemProd).Valor_Unit = rsSaidas_Prod("Pre�o")
    Tab_Prod(gnCtItemProd).Valor_Total = rsSaidas_Prod("Pre�o") * rsSaidas_Prod("Qtde")
  End If
    
  Tab_Prod(gnCtItemProd).Desconto_Perc = rsSaidas_Prod("Desconto")
  Tab_Prod(gnCtItemProd).Aliq_ICM = rsSaidas_Prod("ICM")
  Tab_Prod(gnCtItemProd).Valor_ICM = 0
  Tab_Prod(gnCtItemProd).Aliq_IPI = rsSaidas_Prod("IPI")
  Tab_Prod(gnCtItemProd).Valor_IPI = 0
  Tab_Prod(gnCtItemProd).Valor_Final = rsSaidas_Prod("Pre�o Final")
  Tab_Prod(gnCtItemProd).Pesq1 = rsProdutos("Pesquisa 1")
  Tab_Prod(gnCtItemProd).Pesq2 = rsProdutos("Pesquisa 2")
  Tab_Prod(gnCtItemProd).Pesq3 = rsProdutos("Pesquisa 3")
  Tab_Prod(gnCtItemProd).Nome_Pesq1 = Nome_Pesq1
  Tab_Prod(gnCtItemProd).Nome_Pesq2 = Nome_Pesq2
  Tab_Prod(gnCtItemProd).Nome_Pesq3 = Nome_Pesq3
  Tab_Prod(gnCtItemProd).Local = rsProdutos("Localiza��o") & ""
  Tab_Prod(gnCtItemProd).Descr_Adicional = rsSaidas_Prod("Descricao Adicional") & ""
  
  sAuxGrade = ""
  nCor = 0
  nTamanho = 0
  
  sNomeCor = ""
  sNomeTamanho = ""
  If UCase(rsSaidas_Prod("C�digo")) <> UCase(rsSaidas_Prod("C�digo sem Grade")) Then
     
     sAuxGrade = Right(rsSaidas_Prod("C�digo"), 6)
     nCor = Right(sAuxGrade, 3)
     nTamanho = Left(sAuxGrade, 3)
     
     rsCor.Index = "C�digo"
     rsCor.Seek "=", nCor
     If rsCor.NoMatch Then sNomeCor = ""
     sNomeCor = rsCor("Nome") & ""
     
     rsTamanho.Index = "C�digo"
     rsTamanho.Seek "=", nTamanho
     If rsTamanho.NoMatch Then sNomeTamanho = ""
     sNomeTamanho = rsTamanho("Nome") & ""
        
     Tab_Prod(gnCtItemProd).Cor = nCor & ""
     Tab_Prod(gnCtItemProd).Nome_Cor = sNomeCor & ""
     Tab_Prod(gnCtItemProd).Tamanho = nTamanho & ""
     Tab_Prod(gnCtItemProd).Nome_Tamanho = sNomeTamanho & ""
  End If
  
 
  '27/09/2004 - mpdea
  'Inclu�do o campo de Volumagem por Quantidade
  With Tab_Prod(gnCtItemProd)
    Call IsDataType(dtInteger, rsProdutos.Fields("Volumagem").Value, intVolumagem)
    If intVolumagem > 0 Then
      .VolumagemQtde = "(" & Format(.Qtde \ intVolumagem, "000") & "/" & Format(.Qtde Mod intVolumagem, "000") & ")"
    End If
  End With
  
   
  gnCtItemProd = gnCtItemProd + 1
  
  GoTo Lp_Prod
  
Ve_Servi�os:
  Erase Tab_Serv
  Linha = 0
  Glob_Conta_Serv = 0
  gnCtItemServ = 0
  
  rsSaidas_Serv.Index = "Sequ�ncia"
Lp_Serv:
  rsSaidas_Serv.Seek ">", rsSaidas("Filial"), rsSaidas("Sequ�ncia"), Linha
  If rsSaidas_Serv.NoMatch Then GoTo Ve_Cheque
  If rsSaidas("Filial") <> rsSaidas_Serv("Filial") Then GoTo Ve_Cheque
  If rsSaidas("Sequ�ncia") <> rsSaidas_Serv("Sequ�ncia") Then GoTo Ve_Cheque
  Linha = rsSaidas_Serv("Linha")

  Tab_Serv(gnCtItemServ).C�digo = rsSaidas_Serv("C�digo")
  If rsSaidas_Serv("Completo") = True Then Tab_Serv(gnCtItemServ).Conclu�do = "Sim"
  If rsSaidas_Serv("Completo") = False Then Tab_Serv(gnCtItemServ).Conclu�do = "N�o"
  Tab_Serv(gnCtItemServ).Descri��o = rsSaidas_Serv("Descri��o") & ""
  Tab_Serv(gnCtItemServ).Pre�o_Unit = rsSaidas_Serv("Pre�o")
  Tab_Serv(gnCtItemServ).Qtde = rsSaidas_Serv("Tempo") & ""
 ' Tab_Serv(gnCtItemServ).Pre�o_Total = Format(CStr(Tab_Serv(gnCtItemServ).Pre�o_Unit * Tab_Serv(gnCtItemServ).Qtde), "##############0,00")
  Tab_Serv(gnCtItemServ).Pre�o_Total = Format(CStr(Tab_Serv(gnCtItemServ).Pre�o_Unit * Tab_Serv(gnCtItemServ).Qtde))
  '27/07/2005 - Daniel
  'CST (C�digo de Situa��o Tribut�ria)
  'Finalidade: Atender a realidade da empresa W.V. Hidroan�lise Ltda (J.R. Hidroqu�mica)
  Tab_Serv(gnCtItemServ).CST = rsSaidas_Serv("CST").Value & ""
  
  gnCtItemServ = gnCtItemServ + 1
  
  GoTo Lp_Serv


Ve_Cheque:

  Conta_Fat = 0
  Linhas = 0
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + " " + Chr(13)
  
  rsSa�da_Cheques.Index = "Ordem"
Lp_Cheque:
  rsSa�da_Cheques.Seek ">", gnCodFilial, rsSaidas("Sequ�ncia"), Linhas
  If rsSa�da_Cheques.NoMatch Then GoTo Ve_Fatura
  If rsSa�da_Cheques("Filial") <> gnCodFilial Then GoTo Ve_Fatura
  If rsSa�da_Cheques("Sequ�ncia") <> rsSaidas("Sequ�ncia") Then GoTo Ve_Fatura
  Linhas = rsSa�da_Cheques("Ordem")
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Recebimento em cheque : " + Format(rsSa�da_Cheques("Valor"), "###,###,##0.00")
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "  para " + Format(rsSa�da_Cheques("Bom"), "dd/mm/yyyy") + Chr(13)

  Conta_Fat = Conta_Fat + 1
  GoTo Lp_Cheque


Ve_Fatura:
  Erase Tab_Fat
  
  gnCtParcFat = 0
  Linhas = 0
  
  Glob_Conta_Fat = 0
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + " " + Chr(13)
  
  
  rsSa�da_Parcelas.Index = "Ordem"
Lp_Fat:
  rsSa�da_Parcelas.Seek ">", gnCodFilial, rsSaidas("Sequ�ncia"), Linhas
  If rsSa�da_Parcelas.NoMatch Then GoTo Final
  If rsSa�da_Parcelas("Filial") <> gnCodFilial Then GoTo Final
  If rsSa�da_Parcelas("Sequ�ncia") <> rsSaidas("Sequ�ncia") Then GoTo Final
  Linhas = rsSa�da_Parcelas("Ordem")
  '20/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '             Tratamento para impress�o de nota gerada manualmente
  If Not (gbNotaManual(rsSaidas.Fields("Opera��o").Value, "SAIDA")) Then
    Tab_Fat(gnCtParcFat).N�mero = LTrim(str(rsSaidas("Nota Impressa"))) + "/" + LTrim(str((gnCtParcFat + 1)))
  Else
    Tab_Fat(gnCtParcFat).N�mero = LTrim(str(rsSaidas("Nota Fiscal"))) + "/" + LTrim(str((gnCtParcFat + 1)))
  End If
  '
    Tab_Fat(gnCtParcFat).Valor = rsSa�da_Parcelas("Valor")
    Tab_Fat(gnCtParcFat).Vencimento = rsSa�da_Parcelas("Bom")
  gnCtParcFat = gnCtParcFat + 1
  
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "Parcelamento : " + Format(rsSa�da_Parcelas("Valor"), "###,###,##0.00")
  GLOB_RESUMO_PAGTO = GLOB_RESUMO_PAGTO + "  para " + Format(rsSa�da_Parcelas("Bom"), "dd/mm/yyyy") + Chr(13)
  
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
        MsgBox "Arquivo de configura��o inv�lido."
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
        'N�o imprime linha s� com -
        If Trim(Str_Impre) <> "-" Then
          
          '30/01/2009 - mpdea
          If Not blnEmail Then
            '16/08/2002 - mpdea
            'Inclu�do in�cio da formata��o em negrito
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
            'T�rmino da formata��o em negrito
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
      If Len(Trim(Tab_Prod(nI).C�digo)) > 0 Then
        nCtItens = nCtItens + 1
      End If
    Next nI
    If gnCtItemProd > nCtItens Then
      DisplayMsg "AVISO: N�mero de itens de produtos existentes excedeu a quantidade de itens definida no lay-out do documento..."
    End If
  End If
  
  If Glob_Conta_Serv > 0 Then
    nCtItens = 0
    For nI = 0 To 50
      If Len(Trim(Tab_Serv(nI).C�digo)) > 0 Then
        nCtItens = nCtItens + 1
      End If
    Next nI
    If gnCtItemServ > nCtItens Then
      DisplayMsg "AVISO: N�mero de itens de servi�os existentes excedeu a quantidade de itens definida no lay-out do documento..."
    End If
  End If
  
  If Glob_Conta_Fat > 0 Then
    nCtItens = 0
    For nI = 0 To 50
      If Len(Trim(Tab_Fat(nI).N�mero)) > 0 Then
        nCtItens = nCtItens + 1
      End If
    Next nI
    If gnCtParcFat > nCtItens Then
      DisplayMsg "AVISO: N�mero de parcelas de fatura existentes excedeu a quantidade de itens definida no lay-out do documento..."
    End If
  End If
  
  '------------------------------------------------------------------------------------
  '30/01/2009 - mpdea
  If Not blnEmail Then
    '24/07/2006 - Andrea
    'Grava��o do campo Ticket Impresso na tabela de sa�das.
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
  rsOp_Sa�da.Close
  Set rsOp_Sa�da = Nothing
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
  rsSa�da_Parcelas.Close
  Set rsSa�da_Parcelas = Nothing
  rsSa�da_Cheques.Close
  Set rsSa�da_Cheques = Nothing
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
  gsMsg = "Erro ao imprimir Nota usando o arquivo de configura��o: """ & Nome_Ticket & """."
  If Err.Number = 53 Then
    gsMsg = gsMsg & vbCrLf & "Arquivo de configura��o n�o encontrado."
  Else
    gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  End If
  
  '09/01/2004 - mpdea
  'Adicionado a exibi��o da mensagem de erro
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  SetPrinterName ("REL")
  On Error Resume Next
  Close #nFileNum
  Exit Function
  
End Function
