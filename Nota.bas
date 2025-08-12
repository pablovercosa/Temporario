Attribute VB_Name = "modPrintNota"
Option Explicit

Public gsCPF_Cnpj As String
Public gsIE As String
Public gsNomeCliente As String

Public Glob_Sequ�ncia As Double
Public Glob_Nota_Impressa As Long
Public Glob_Nome_Opera��o As String
Public Glob_C�digo_Fiscal As String
'06/05/2007 - Anderson
'Implementa��o de impress�o de CFOP's por servi�o
'-------------------------------------------
Public Glob_C�digo_Fiscal_Completo As String
Public Glob_C�digo_Fiscal_Item(4) As String
'-------------------------------------------

'24/04/2008 - mpdea
'Descri��o e total por CFOP relacionado a movimenta��o
Public Glob_Nome_Operacao_CFOP(4) As String
Public Glob_Total_CFOP(4) As Double

Public Glob_Cod_Opera��o As Integer
Public Glob_C�digo_Cli As Long
Public Glob_Nome As String
Public Glob_Fantasia As String
Public Glob_CGC As String
Public Glob_Inscri��o As String
Public Glob_Endere�o As String

'23/10/2009 - mpdea
Public Glob_NumeroEndereco As String

Public Glob_Complemento As String
Public Glob_Bairro As String
Public Glob_CEP As String
Public Glob_Cidade As String
Public Glob_Fone1 As String
Public Glob_Fone2 As String
Public Glob_Estado As String
Public Glob_Base_ICM As Double
Public Glob_Valor_ICM As Double
Public Glob_Base_ICM_Sub As Double
Public Glob_Valor_ICM_Sub As Double
Public Glob_Total_Produto As Double
Public Glob_Total_Desconto As Double
Public Glob_Total_Servi�o As Double
Public Glob_Total_ISS As Double
Public Glob_Frete As Double
Public Glob_IPI As Double
Public Glob_Total_Pagar As Double
Public Glob_Cod_Vendedor As String
Public Glob_Nome_Vendedor As String
Public Glob_Apelido As String            '20/11/2006 - Anderson - Informa o apelido do vendedor. Solicitante: Technomax

'15/08/2002 - mpdea
'Inclu�do o campo de informa��es sobre o or�amento (n�mero do or�amento e terminal)
Public gstrInfoNrOrcamento As String

'06/09/2002 - mpdea
'Inclu�do o campo para exibi��o do Desconto no SubTotal
Public g_dblDescontoSubTotal As Double

'07/08/2003 - mpdea
'Inclu�do C�digo e Nome do T�cnico
Public g_intCodigoTecnico As Integer
Public g_strNomeTecnico As String

'19/08/2003 - mpdea
'Inclu�do campo para totalizador de produtos com desconto no subtotal
Public g_dblTotalProdutosDST As Double

'26/08/2003 - mpdea
'Inclu�do campo para totalizador de produtos menos total de descontos
Public g_dblTotalProdutosMenosDescontos As Double


'01/02/2006 - mpdea
'Cole��o com as mensagens para Nota Fiscal
Private m_clcMensagens As Collection

  
'08/01/2004 - Daniel
'Inclu�do vars para o campo Valor Recebido e
'Troco da tabela de Sa�das
Public g_dblValorRecebido As Double
Public g_dblTroco As Double

'27/04/2005 - Daniel
'Inclu�do var para tratamento do campo Seguro
'da table Sa�das
Public g_dblSeguro As Double

'08/01/2004 - Daniel
'Inclu�do var para impress�o
'da quantidade de itens - Sum na Tabela
'[Sa�das - Produto]
Public g_sngQtdeItens As Single

'30/01/2004 - Daniel
'Inclus�o dos campos Percentual CSLL,
'Percentual COFINS, Percentual PIS,
'Percentual IRFF da tabela Par�metros Filial
'e Totais: Total CSLL, Total COFINS
'Total PIS, Total IRRF
Public g_dblPercentualCSLL As Double
Public g_dblPercentualCOFINS As Double
Public g_dblPercentualPIS As Double
Public g_dblPercentualIRRF As Double

Public g_dblTotalCSLL As Double
Public g_dblTotalCOFINS As Double
Public g_dblTotalPIS As Double
Public g_dblTotalIRRF As Double
'----------------------------------------------

'----------------------------------------------------------------------
'13/04/2004 - Daniel
'Inclus�o dos Campos:
'Sa�das.[Num Autoriza��o], Sa�das.MesX, Cli_For.[Endere�o Cob],
'Cli_For.[Complemento Cob], Cli_For.[Bairro Cob], Cli_For.[Cidade Cob],
'Cli_For.[Estado Cob] e Cli_For.[CEP Cob]
Public g_lngNumAutorizacao As Long
Public g_intMesX           As Integer
Public g_strEnderecoCob    As String
Public g_strComplementoCob As String
Public g_strBairroCob      As String
Public g_strCidadeCob      As String
Public g_strEstadoCob      As String
Public g_strCepCob         As String
'----------------------------------------------------------------------

'----------------------------------------------------------------------
'06/05/2004 - Daniel
'Inclus�o do campo ObsIsentoIPI da tabela Cli_For
'case Embalavi, dispon�vel para os demais clientes
Public g_strObsIsentoIPI   As String
'17/05/2004 - Daniel
'Inclus�o do campo Diferimento.ObsDiferimento
Public g_strObsDiferimento As String
'----------------------------------------------------------------------


Public gsTransportadora As String
Public gsCNPJTransportadora As String
Public gsIETransportadora As String
Public gsEnderTransportadora As String
Public gsMunicipioTransportadora As String
Public gsUFTransportadora As String
Public gsPlaca As String
Public gsUfrmPlaca As String

Public gsQtdeTrans As String
Public gsMarcaTrans As String
Public gsEspecieTrans As String
Public gsPesoBruto As String
Public gsPesoLiquido As String
Public gnPesoLiquido As Double
Public gnPesoBruto As Double

Public gsFretePago As String
Public Glob_Data_Dev_Emp As Date
Public Glob_Prometido As String
Public Glob_Aprovado As String
Public Glob_Cod_T�cnico As String
Public Glob_Nome_T�cnico As String
Public Glob_gnCodFilial As Integer
Public Glob_Nome_Filial As String
Public Glob_Data As Date
Public Glob_Data_Sa�da As Date
Public Glob_Hora_Sa�da As String
Public Glob_Cod_Caixa As Integer
Public Glob_Nome_Caixa As String
Public Glob_Tab_Pre�o As String
Public Glob_RefrmInterna As String
Public Glob_Obs_Mov As String

Public gsRetornoDoc As String
Public gsObsDoc(8) As String
Public gsDocFileName As String
Public strMotivoCancelamento As String


Rem As abaixo s�o usadas pelo boleto
Public Glob_Data_Emiss�o As String
Public Glob_Fatura As String
Public Glob_Descri��o As String
Public Glob_Vencimento As String
Public Glob_Valor As String
Public Glob_Desconto As String
Public Glob_Acr�scimo As String
Public Glob_Mensagem_Cli As String
'15/01/2004 - Daniel
'Var para armazenamento do Valor Recebido
'proveniente da tabela [Contas a Receber]
Public g_dblValorRecebidoCR As Double


Rem as abaixo s�o do extenso
Public Extenso1_60 As String
Public Extenso61_120 As String
Public Extenso121_180 As String

Public Extenso1_45 As String
Public Extenso46_90 As String
Public Extenso91_135 As String
Public Extenso136_180 As String

Public Extenso1_30 As String
Public Extenso31_60 As String
Public Extenso61_90 As String
Public Extenso91_120 As String
Public Extenso121_150 As String
Public Extenso151_180 As String


'Campo abaixo � usado pelo Resumo de Pagamento
Public GLOB_RESUMO_PAGTO As String


'19/12/2007 - Anderson
'Implementa��o do NSU
Public NSU As String
Public NSU_Data As String
Public NSU_Hora As String

'04/05/2004 - Daniel
'Personaliza��o Embalavi
Private m_blnEmbalavi As Boolean

Private Type Tabela1
  N�mero As String
  Valor As Double
  Vencimento As Date
End Type
Public Tab_Fat(50) As Tabela1

Private Type Tabela2
  C�digo As String
  C�digo_Prod_Forn As String
  Nome As String
  
  '04/08/2002 - mpdea
  'Inclu�do os campos para impress�o espec�fica do nome do produto como
  'est� no campo Nome do cadastro ou o campo Nome para nota (Fixo)
  NomeCadastro As String
  NomeNota As String
  
  C_Fiscal As String
  
  '22/09/2005 - mpdea
  'Inclu�do campo para descri��o da Classifica��o Fiscal
  DescricaoClassificaoFiscal As String
  
  '29/04/2008 - mpdea
  'CFOP do produto
  CFOP As String

  S_Tribut�ria As String
  
  '05/05/2011 - mpdea
  'NBM/NCM do produto
  CodigoNbmNcm As String

  Unid As String
  Qtde As Single
  Valor_Unit As Double
  Valor_Total As Double
  Desconto_Perc As Double
  'Aliq_ICM As Integer
  Aliq_ICM As Double
  Valor_ICM As Double
  'Aliq_IPI As Integer
  Aliq_IPI As Double
  Valor_IPI As Double
  Valor_Final As Double
  Pesq1 As Long
  Nome_Pesq1 As String
  Pesq2 As Long
  Nome_Pesq2 As String
  Pesq3 As Long
  Nome_Pesq3 As String
  Local  As String
  Descr_Adicional As String
  Cor As Long
  Nome_Cor As String
  Tamanho As Long
  Nome_Tamanho As String
  
  '27/09/2004 - mpdea
  'Inclu�do campo para exibi��o da Volumagem por Quantidade
  VolumagemQtde As String
  '29/11/2004 - Daniel
  'Inclu�do os campos Lote e Data de Validade
  Lote As String
  DataValidade As String
  '27/04/2005 - Daniel
  'Inclu�do o campo Fabricante
  Fabricante As String
End Type

Public Tab_Prod(500) As Tabela2

Private Type Tabela3
  C�digo      As Integer
  Descri��o   As String
  Qtde        As Single
  Pre�o_Unit  As Double
  Pre�o_Total As Double
  Conclu�do   As String
  '27/07/2005 - Daniel
  'CST (C�digo de Situa��o Tribut�ria)
  'Finalidade: Atender a realidade da empresa W.V. Hidroan�lise Ltda (J.R. Hidroqu�mica)
  CST         As String
  '29/04/2008 - mpdea
  'CFOP do servi�o
  CFOP As String
End Type
Public Tab_Serv(50) As Tabela3

Private nFileNum As Integer

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
Private rsEntradas As Recordset
Private rsEntradas_Prod As Recordset
Private rsOp_Entrada As Recordset
Private rsCor As Recordset
Private rsTamanho As Recordset
'17/05/2004 - Daniel - Inclus�o da tabela Diferimento
Private rstDiferimento As Recordset
'06/05/2007 - Anderson
'Implementa��o de CFOP's por produto
Private rstCFOP 's As Recordset

Public Funcionario As String

Public Function Imprime_Nota(ByVal Nome_Nota As String, ByVal Filial As Integer, ByVal Sequ�ncia As Long) As Integer
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
  Dim Extenso_Tot As String
  Dim Conta_Fat As Integer
  Dim Val_Imposto As Double
  Dim nFileNum As Integer
  Dim nCtLin As Integer
  Dim nComprPag As Integer
  Dim sParte As String
  Dim nCtItens As Integer
  Dim nI As Integer
  Dim sAuxGrade As String
  Dim nCor As Long
  Dim sNomeCor As String
  Dim nTamanho As Long
  Dim sNomeTamanho As String
  Dim intVolumagem As Integer
  Dim intContadorCFOP As Integer '16/10/2007 - Anderson - Implementa��o da impress�o de v�rios CFOP's
  
  Dim clcLayoutFile As Collection
  Dim strLayoutLinha As String
  
  '27/05/2003 - mpdea
  'C�digo comentado para uso futuro
'  '21/05/2003 - mpdea
'  'Flag para plataformas NT
'  Dim blnWindowsNT As Boolean
'
'  blnWindowsNT = IsWindowsNT()
  
  '----------------------------------------------
  '09/01/2004 - Daniel
  'Para Finalidade de soma da Qtde da tabela [Sa�das - Produtos]
  Dim rsSomaQtde As Recordset
  
  '24/04/2008 - mpdea
  'Total por CFOP relacionado a movimenta��o
  Dim dbl_total_cfop_produtos As Double
  Dim dbl_total_cfop_servicos As Double
  
  
  On Error GoTo ErrHandler
  
  
  Set rsSaidas = db.OpenRecordset("Sa�das", , dbReadOnly)
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
  Set rsCor = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsTamanho = db.OpenRecordset("Tamanhos", , dbReadOnly)
  
  SetPrinterName ("NOTA")
  
  gsInitPrinter = ""
  Call ResetPrinter
  
  nFileNum = FreeFile
  Open Nome_Nota For Input As #nFileNum
  
  Input #nFileNum, Texto
  If Left(Texto, 23) <> "*** Configura��es Nota:" Then
    gsTitle = LoadResString(201)
    gsMsg = "Layout do cabe�alho do arquivo de configura��o """ & Nome_Nota & """ diferente do esperado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Nota = 3
    Close #nFileNum
    Exit Function
  End If
  
  nComprPag = 0
  sParte = UCase(Trim(Mid(Texto, 74, 4)))
  If Len(sParte) > 0 Then
    If sParte <> "N�O" And sParte <> "LIN" Then
      If Not IsNumeric(sParte) Then
        DisplayMsg "Valor para par�metro de comprimento da p�gina pode ser: N�O, LIN ou <99> (inteiro dois digitos)."
        Imprime_Nota = 3
        Close #nFileNum
        Exit Function
      End If
      If Val(sParte) <= 0 Or Val(sParte) > 20 Then
        DisplayMsg "Comprimento da p�gina em polegadas inv�lido."
        Imprime_Nota = 3
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
        Open Nome_Nota For Input As #nFileNum
        Input #nFileNum, Texto
      End If
    End If
  End If

  If Mid(Texto, 40, 3) = "SIM" Then
    If SetCompressPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel usar compress�o na impressora solicitada pelo arquivo de configura��o: """ & Nome_Nota & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      SetPrinterName ("REL")
      Close #nFileNum
      Imprime_Nota = 3
      Exit Function
    End If
  End If
  
  If Mid(Texto, 55, 3) = "SIM" Then
    If SetOitavoPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel ajustar a impressora para 1/8 solicitada pelo arquivo de configura��o: """ & Nome_Nota & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Close #nFileNum
      Imprime_Nota = 1
      SetPrinterName ("REL")
      Exit Function
    End If
  End If
  
  If sParte = "LIN" Then
    If SetComprimPagLinPrinter(Filial, nCtLin) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel alterar o comprimento de p�gina na impressora."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Imprime_Nota = 4
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
        Imprime_Nota = 4
        SetPrinterName ("REL")
        Close #nFileNum
        Exit Function
      End If
    End If
  End If
  
  Call SetPrinterCommand(gsInitPrinter)
  
  
  Rem Acha Sa�da
  rsSaidas.Index = "Sequ�ncia"
  rsSaidas.Seek "=", Filial, Sequ�ncia
  If rsSaidas.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "N�o � poss�vel imprimir a Nota, registro de sa�da n�o encontrado para Filial= " & Filial & ", " & "Seq��ncia= " & Sequ�ncia
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Nota = 1
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
    Imprime_Nota = 2
    Close #nFileNum
    Exit Function
  End If
  
    
  rsOp_Sa�da.Index = "C�digo"
  rsOp_Sa�da.Seek "=", rsSaidas("Opera��o")
  If rsOp_Sa�da.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Opera��o de Sa�da referida pelo registro de Sa�das n�o foi localizada: Opera��o=" & rsSaidas("Opera��o")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Nota = 3
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
    Imprime_Nota = 4
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
    Imprime_Nota = 5
    Close #nFileNum
    Exit Function
  End If
  
  
  Call Limpa_Vari�veis_Nota
  
  Glob_gnCodFilial = Filial
  Glob_Nome_Filial = rsParametros("Nome") & ""
  Glob_Data = Format(rsSaidas("Data"), "dd/mm/yyyy")
  Glob_Data_Sa�da = Format(Date, "dd/mm/yyyy")
  Glob_Hora_Sa�da = Format(Time, "hh:mm:ss")
  Glob_Cod_Opera��o = rsSaidas("Opera��o")
  Glob_Nome_Opera��o = rsOp_Sa�da("Nome") & ""
  
  '17/05/2004 - Daniel
  'Populando g_strObsDiferimento
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
            If Mid((.Fields("Estado").Value), 1, 2) = "PR" And .Fields("F�sica_Jur�dica").Value = "J" Then
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
  '-------------------------
  
  '30/01/2004 - Daniel
  'Populando as vars de tratamento de impostos sobre Servi�os
  '02/04/2004 - Busca da tabela Sa�das e n�o mais de Par�metros
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
  
  '27/04/2005 - Tratamento para o Seguro vindo da table Sa�das
  If IsNumeric(rsSaidas.Fields("Seguro").Value) Then
    g_dblSeguro = Format(rsSaidas.Fields("Seguro").Value, FORMAT_VALUE)
  Else
    g_dblSeguro = 0
  End If
  
  '30/10/2003 - mpdea
  'Corrigido RT-94
  Glob_C�digo_Fiscal = rsOp_Sa�da("C�digo Fiscal") & ""
  
  
  '06/05/2007 - Anderson
  'Implementa��o da impress�o para CFOP's de Produtos e servi�os
  If Len(Trim(rsOp_Sa�da("C�digo Fiscal"))) > 0 Then
    'Glob_C�digo_Fiscal_Completo = Mid(Trim(rsOp_Sa�da("C�digo Fiscal")), 1, 4)
    '16/10/2008 - Anderson
    'Informa o primeiro CFOP
    Call VerificaCFOP(Mid(Trim(rsOp_Sa�da("C�digo Fiscal")), 1, 4))
  End If
  
  '24/04/2008 - mpdea
  'Inclu�do op��o somente leitura
  Set rstCFOP = db.OpenRecordset("SELECT TOP 4 CFOP From [Sa�das - Produtos] GROUP BY CFOP, Sequ�ncia, Filial HAVING Filial = " & Filial & " AND Sequ�ncia=" & rsSaidas("Sequ�ncia"), dbOpenDynaset, dbReadOnly)
  
  Do Until rstCFOP.EOF
    
    '16/10/2007 - Anderson
    'Coloca CFOP de Servi�o
    Call VerificaCFOP(Trim(Mid("" & rstCFOP("CFOP"), 1, 4)))
    'Glob_C�digo_Fiscal_Item(rstCFOP.AbsolutePosition) = Trim(Mid(rstCFOP("CFOP"), 1, 4))
    
    'If Len(Glob_C�digo_Fiscal_Completo) > 0 Then
    '  Glob_C�digo_Fiscal_Completo = Glob_C�digo_Fiscal_Completo & "/" & Glob_C�digo_Fiscal_Item(rstCFOP.AbsolutePosition)
    'End If
    rstCFOP.MoveNext
  Loop
  
  rstCFOP.Close
  
  '24/04/2008 - mpdea
  'Inclu�do op��o somente leitura
  Set rstCFOP = db.OpenRecordset("SELECT TOP 4 CFOP From [Sa�das - Servi�os] GROUP BY CFOP, Sequ�ncia, Filial HAVING Filial = " & Filial & " AND Sequ�ncia=" & rsSaidas("Sequ�ncia"), dbOpenDynaset, dbReadOnly)
  
  Do Until rstCFOP.EOF
    
    '16/10/2007 - Anderson
    'Coloca CFOP de Servi�o
    Call VerificaCFOP(Trim(Mid("" & rstCFOP("CFOP"), 1, 4)))
    'Glob_C�digo_Fiscal_Item(rstCFOP.AbsolutePosition) = Trim(Mid(rstCFOP("CFOP"), 1, 4))
    
    'If Len(Glob_C�digo_Fiscal_Completo) > 0 Then
    '  Glob_C�digo_Fiscal_Completo = Glob_C�digo_Fiscal_Completo & "/" & Glob_C�digo_Fiscal_Item(rstCFOP.AbsolutePosition)
    'End If
    rstCFOP.MoveNext
  Loop
  
  rstCFOP.Close
  
  'Cria campo de CFOP Completo
  For intContadorCFOP = 0 To 4
    If Len(Glob_C�digo_Fiscal_Completo) = 0 Then
      Glob_C�digo_Fiscal_Completo = Glob_C�digo_Fiscal_Item(intContadorCFOP)
    Else
      If Len(Glob_C�digo_Fiscal_Item(intContadorCFOP)) > 0 Then
        Glob_C�digo_Fiscal_Completo = Glob_C�digo_Fiscal_Completo & "/" & Glob_C�digo_Fiscal_Item(intContadorCFOP)
      End If
    End If
  
    '24/04/2008 - mpdea
    'Descri��o e total por CFOP relacionado a movimenta��o
    If Len(Glob_C�digo_Fiscal_Item(intContadorCFOP)) > 0 Then
      'Descri��o
      Set rstCFOP = db.OpenRecordset("SELECT Nome FROM [Opera��es Sa�da] WHERE [C�digo Fiscal] = '" & Glob_C�digo_Fiscal_Item(intContadorCFOP) & "'", dbOpenDynaset, dbReadOnly)
      With rstCFOP
        If Not (.BOF And .EOF) Then
          Glob_Nome_Operacao_CFOP(intContadorCFOP) = .Fields("Nome").Value & ""
        End If
        .Close
      End With
      'Valor total
      dbl_total_cfop_produtos = 0
      Set rstCFOP = db.OpenRecordset("SELECT SUM([Pre�o Final]) AS Total From [Sa�das - Produtos] WHERE Filial = " & Filial & " AND Sequ�ncia = " & Sequ�ncia & " AND CFOP = '" & Glob_C�digo_Fiscal_Item(intContadorCFOP) & "'", dbOpenDynaset, dbReadOnly)
      With rstCFOP
        If Not (.BOF And .EOF) Then
          Call IsDataType(dtDouble, .Fields("Total").Value, dbl_total_cfop_produtos)
        End If
        .Close
      End With
      dbl_total_cfop_servicos = 0
      Set rstCFOP = db.OpenRecordset("SELECT SUM(Tempo * Pre�o) AS Total From [Sa�das - Servi�os] WHERE Filial = " & Filial & " AND Sequ�ncia = " & Sequ�ncia & " AND CFOP = '" & Glob_C�digo_Fiscal_Item(intContadorCFOP) & "'", dbOpenDynaset, dbReadOnly)
      With rstCFOP
        If Not (.BOF And .EOF) Then
          Call IsDataType(dtDouble, .Fields("Total").Value, dbl_total_cfop_servicos)
        End If
        .Close
      End With
      Glob_Total_CFOP(intContadorCFOP) = Format(dbl_total_cfop_produtos + dbl_total_cfop_servicos, FORMAT_VALUE)
    End If
  Next
  
  Set rstDiferimento = db.OpenRecordset("SELECT ObsDiferimento, EstadoCorrente FROM Diferimento WHERE Filial = " & Filial, dbOpenDynaset)
  
  Glob_Sequ�ncia = rsSaidas("Sequ�ncia")
  Glob_Cod_Vendedor = rsSaidas("Digitador")
  
  '20/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '
  'O sistema Julga se a nota fiscal foi criada
  'automaticamente ou manualmente a partir da opera��o escolhida
  'Nota: Caso seja manualmente (notas de bloquinho), o sistema n�o
  'incrementou o contador pois o sistema estava fora do ar
  If Not (gbNotaManual(rsSaidas.Fields("Opera��o").Value, "SAIDA")) Then
    Glob_Nota_Impressa = rsSaidas("Nota Impressa")
  Else
    Glob_Nota_Impressa = CLng("0" & rsSaidas("Nota Fiscal"))
  End If

  '26/08/2002 - mpdea
  'Inclu�do o campo de informa��es sobre o or�amento (n�mero do or�amento e terminal)
  gstrInfoNrOrcamento = rsSaidas.Fields("InfoNrOrcamento").Value & ""
  
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
  'Inclu�do var g_sngQtdeItens para a Soma da Qtde
  'Populando g_sngQtdeItens
  Set rsSomaQtde = db.OpenRecordset("SELECT SUM(Qtde) AS Soma FROM [Sa�das - Produtos] WHERE Filial =" & Glob_gnCodFilial & " AND Sequ�ncia =" & Glob_Sequ�ncia, dbOpenDynaset)
  'g_sngQtdeItens = rsSomaQtde.Fields("Soma")
  Call IsDataType(dtSingle, rsSomaQtde.Fields("Soma"), g_sngQtdeItens)
  '09/01/2004 - Daniel
  rsSomaQtde.Close
  Set rsSomaQtde = Nothing
  '----------------------------------------------
  
  
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", Glob_Cod_Vendedor
  If rsFuncionarios.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Digitador referido pelo registro de Sa�das n�o foi localizado: Digitador=" & rsSaidas("Digitador")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Imprime_Nota = 6
    Close #nFileNum
    Exit Function
  End If
  
  '20/11/2006 - Anderson
  'Inclu�do o campo apelido do funcion�rio.
  'Solicitante: Technomax
  Glob_Nome_Vendedor = rsFuncionarios("Nome") & ""
  Glob_Apelido = rsFuncionarios("Apelido") & ""
  
  
  '-------------------------------------------------------------------
  '12/08/2003 - mpdea
  'Corrigido busca com c�digo igual a zero
  '
  '07/08/2003 - mpdea
  'Inclu�do C�digo e Nome do T�cnico
  Call IsDataType(dtInteger, rsSaidas.Fields("T�cnico").Value, g_intCodigoTecnico)
  
  If g_intCodigoTecnico > 0 Then
    rsFuncionarios.Seek "=", g_intCodigoTecnico
    If rsFuncionarios.NoMatch Then
      gsTitle = LoadResString(201)
      gsMsg = "T�cnico referido pelo registro de Sa�das n�o foi localizado: T�cnico = " & g_intCodigoTecnico
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      SetPrinterName ("REL")
      Imprime_Nota = 7
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
  'Populando g_strObsIsentoIPI
  g_strObsIsentoIPI = Format((rsCliFor("ObsIsentoIPI").Value), 70)
  '------------------------------------------------------------
  Glob_Cod_Caixa = rsSaidas("Caixa")
  Glob_Nome_Caixa = rsCaixas("Descri��o") & ""
  Glob_Tab_Pre�o = rsSaidas("Tabela") & ""
  Glob_RefrmInterna = rsSaidas("Refer�ncia") & ""
  Glob_Obs_Mov = rsSaidas("Observa��es") & ""
  
  Glob_Base_ICM = rsSaidas("Base ICM")
  Glob_Valor_ICM = rsSaidas("Valor ICM")
  Glob_Base_ICM_Sub = rsSaidas("Base ICM Subs")
  Glob_Valor_ICM_Sub = rsSaidas("Valor ICM Subs")
  Glob_Total_Produto = rsSaidas("Produtos")
  Glob_Total_Desconto = rsSaidas("Desconto")
  Glob_Frete = rsSaidas("Frete")
  Glob_IPI = rsSaidas("IPI")
  Glob_Total_Pagar = rsSaidas("Total")
  
  Glob_Total_Servi�o = rsSaidas("Servi�os")
  Glob_Total_ISS = rsSaidas("Valor ISS")
  
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
  Call IsDataType(dtDouble, rsSaidas.Fields("DescontoSubTotal").Value, g_dblDescontoSubTotal)
  
  
  '19/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos com desconto no subtotal
  g_dblTotalProdutosDST = Glob_Total_Produto - g_dblDescontoSubTotal
  
  
  '26/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos menos total de descontos
  g_dblTotalProdutosMenosDescontos = Glob_Total_Produto - Glob_Total_Desconto
  
  
  '01/02/2006 - mpdea
  'Carrega as mensagens para Nota Fiscal
  Set m_clcMensagens = New Collection
  Call GetMensagemNotaFiscal(Filial, Sequ�ncia, m_clcMensagens)
  
  
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
    
  '19/12/2007 - Anderson
  'Implementa��o do NSU
  NSU = Format("" & rsSaidas("NSU"), "0000000000")
  NSU_Data = Format("" & rsSaidas("NSU_Data"), "dd/mm/yy")
  NSU_Hora = Format("" & rsSaidas("NSU_Hora"), "hh:nn")
  
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
  DoEvents
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
  
  '20/08/2004 - Daniel
  'Inclu�do Tratamento para imprimir a Descri��o Adicional quando
  'necess�rio
  If rsProdutos("UsaDescrAdic").Value Then
    Tab_Prod(gnCtItemProd).Nome = rsSaidas_Prod("Descricao Adicional") & ""
  Else
    Tab_Prod(gnCtItemProd).Nome = rsProdutos("Nome Nota") & ""
  End If
  
  If rsProdutos("Nome Nota") = "" Then Tab_Prod(gnCtItemProd).Nome = rsProdutos("Nome") & ""
  
  '04/09/2002 - mpdea
  'Inclu�do os campos para impress�o espec�fica do nome do produto como
  'est� no campo Nome do cadastro ou o campo Nome para nota (Fixo)
  Tab_Prod(gnCtItemProd).NomeCadastro = rsProdutos("Nome") & ""
  Tab_Prod(gnCtItemProd).NomeNota = rsProdutos("Nome Nota") & ""
  
  Tab_Prod(gnCtItemProd).C_Fiscal = rsProdutos("Classifica��o Fiscal") & ""
  
  '29/04/2008 - mpdea
  'CFOP do produto
  Tab_Prod(gnCtItemProd).CFOP = rsSaidas_Prod.Fields("CFOP").Value & ""
  
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
  
  '05/05/2011 - mpdea
  'NBM/NCM do produto
  Tab_Prod(gnCtItemProd).CodigoNbmNcm = rsProdutos.Fields("CodigoNBM").Value & ""
  
  Tab_Prod(gnCtItemProd).Unid = rsProdutos("Unidade Venda") & ""
  Tab_Prod(gnCtItemProd).Qtde = rsSaidas_Prod("Qtde")
  
  If g_bln5CasasDecimais Then
    '04/05/2004 - Daniel
    'Personaliza��o Embalavi
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
  
  Tab_Prod(gnCtItemProd).Desconto_Perc = Format(rsSaidas_Prod("Desconto"), "############0.00")
  Tab_Prod(gnCtItemProd).Aliq_ICM = Format(rsSaidas_Prod("ICM"), "############0.00")
  
  Val_Imposto = rsSaidas_Prod("Pre�o") - (rsSaidas_Prod("Desconto") * rsSaidas_Prod("Pre�o") / 100)
  Val_Imposto = Val_Imposto * rsSaidas_Prod("Qtde")
  Val_Imposto = Val_Imposto * rsSaidas_Prod("ICM") / 100
  Val_Imposto = Format(Val_Imposto, "############0.00")
  Tab_Prod(gnCtItemProd).Valor_ICM = Val_Imposto
  Tab_Prod(gnCtItemProd).Local = rsProdutos("Localiza��o") & ""
  Tab_Prod(gnCtItemProd).Aliq_IPI = Format(rsSaidas_Prod("IPI"), "############0.00")
  
  '27/04/2005 - Daniel
  'Tratamento para o campo Produtos.Fabricante
  Tab_Prod(gnCtItemProd).Fabricante = rsProdutos("Fabricante").Value & ""
  '----------------------------------------------------------------------
  
  Val_Imposto = rsSaidas_Prod("Pre�o") - (rsSaidas_Prod("Desconto") * rsSaidas_Prod("Pre�o") / 100)
  Val_Imposto = Val_Imposto * rsSaidas_Prod("Qtde")
  Val_Imposto = Val_Imposto * rsSaidas_Prod("IPI") / 100
  Val_Imposto = Format(Val_Imposto, "############0.00")
  Tab_Prod(gnCtItemProd).Valor_IPI = Val_Imposto
  
  Tab_Prod(gnCtItemProd).Valor_Final = Format(rsSaidas_Prod("Pre�o Final"), FORMAT_VALUE)
  Tab_Prod(gnCtItemProd).Pesq1 = rsProdutos("Pesquisa 1")
  Tab_Prod(gnCtItemProd).Pesq2 = rsProdutos("Pesquisa 2")
  Tab_Prod(gnCtItemProd).Pesq3 = rsProdutos("Pesquisa 3")
  '29/11/2004 - Daniel
  'Inclu�do os campos Lote e Data de Validade
  If Len(rsProdutos("Lote").Value) > 0 Then Tab_Prod(gnCtItemProd).Lote = rsProdutos("Lote").Value
  If IsDate(rsProdutos("DataValidade").Value) Then Tab_Prod(gnCtItemProd).DataValidade = CStr(rsProdutos("DataValidade").Value)
  
  Tab_Prod(gnCtItemProd).Nome_Pesq1 = Nome_Pesq1
  Tab_Prod(gnCtItemProd).Nome_Pesq2 = Nome_Pesq2
  Tab_Prod(gnCtItemProd).Nome_Pesq3 = Nome_Pesq3
  Tab_Prod(gnCtItemProd).Descr_Adicional = rsSaidas_Prod("Descricao Adicional") & ""
   
  sAuxGrade = ""
  nCor = 0
  nTamanho = 0
  
  sNomeCor = ""
  sNomeTamanho = ""
   
  If rsSaidas_Prod("C�digo") <> UCase(rsSaidas_Prod("C�digo sem Grade")) Then
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
  DoEvents
  rsSaidas_Serv.Seek ">", rsSaidas("Filial"), rsSaidas("Sequ�ncia"), Linha
  If rsSaidas_Serv.NoMatch Then GoTo Ve_Fatura
  If rsSaidas("Filial") <> rsSaidas_Serv("Filial") Then GoTo Ve_Fatura
  If rsSaidas("Sequ�ncia") <> rsSaidas_Serv("Sequ�ncia") Then GoTo Ve_Fatura
  Linha = rsSaidas_Serv("Linha")

  Tab_Serv(gnCtItemServ).C�digo = rsSaidas_Serv("C�digo")

  '29/04/2008 - mpdea
  'CFOP do servi�o
  Tab_Serv(gnCtItemServ).CFOP = rsSaidas_Serv.Fields("CFOP").Value & ""

  If rsSaidas_Serv("Completo") = True Then Tab_Serv(gnCtItemServ).Conclu�do = "Sim"
  If rsSaidas_Serv("Completo") = False Then Tab_Serv(gnCtItemServ).Conclu�do = "N�o"
  Tab_Serv(gnCtItemServ).Descri��o = rsSaidas_Serv("Descri��o") & ""
  Tab_Serv(gnCtItemServ).Qtde = CSng(gsHandleNull(rsSaidas_Serv("Tempo") & ""))
  Tab_Serv(gnCtItemServ).Pre�o_Unit = gsHandleNull(rsSaidas_Serv("Pre�o"))
  Tab_Serv(gnCtItemServ).Pre�o_Total = Format(Tab_Serv(gnCtItemServ).Qtde * rsSaidas_Serv("Pre�o"), "##############0.00")
  '27/07/2005 - Daniel
  'CST (C�digo de Situa��o Tribut�ria)
  'Finalidade: Atender a realidade da empresa W.V. Hidroan�lise Ltda (J.R. Hidroqu�mica)
  Tab_Serv(gnCtItemServ).CST = rsSaidas_Serv("CST").Value & ""
  
  gnCtItemServ = gnCtItemServ + 1
  
  GoTo Lp_Serv


Ve_Fatura:
  
  Erase Tab_Fat
  
  Conta_Fat = 0
  Linhas = 0
  
  Glob_Conta_Fat = 0
  gnCtParcFat = 0
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
  GoTo Lp_Fat




Final:

  
  '----------------------------------------------------------------------------
  '18-19/02/2004 - mpdea
  '
  'Leitura do arquivo de layout e impress�o de acordo com o mesmo
  '----------------------------------------------------------------------------
  Set clcLayoutFile = New Collection
  '
  'Realiza a leitura do arquivo de configura��o
  Do
    Input #nFileNum, strLayoutLinha
    'Verifica final do arquivo de configura��o
    If strLayoutLinha = "*** Fim de arquivo ***" Then Exit Do
    'Remove aspas
    strLayoutLinha = Apaga_Aspas(strLayoutLinha)
    'Adiciona linha de layout a cole��o
    clcLayoutFile.Add strLayoutLinha
  Loop
  Close #nFileNum
  '
  'Realiza a impress�o da nota baseada no arquivo de configura��o
  Call modPrintNotaMN.PrintNotaFiscalByColl(clcLayoutFile)
  '----------------------------------------------------------------------------



'  Final = False
'  Do
'    Input #nFileNum, Texto
'    If Texto = "*** Fim de arquivo ***" Then Final = True
'    If Final = False Then
'      Texto = Apaga_Aspas(Texto)
'      Final_Linha = False
'      If Len(Texto) < 3 Then
'        MsgBox "Arquivo de configura��o inv�lido."
'        Imprime_Nota = 1
'        Exit Function
'      End If
'      Especial2 = False
'      If Left(Texto, 13) = "[LINHA_BRANCO" Then
'        Especial2 = True
'        Linhas = Val(Mid(Texto, 15))
'        Do
'
'          '27/05/2003 - mpdea
'          'C�digo comentado para uso futuro
''          '21/05/2003 - mpdea
''          'Corrige a impress�o de linhas em branco para plataformas NT
''          If blnWindowsNT Then
''            Printer.Print vbCrLf
''          Else
'            Printer.Print
''          End If
'
'          Linhas = Linhas - 1
'        Loop Until Linhas = 0
'      End If
'      If Especial2 = False Then
'        If InStr(Texto, "Obs") > 0 Then
'          Texto = Texto
'        End If
'        Str_Impre = Retorna_Texto(Texto)
'
'        '16/08/2002 - mpdea
'        'Inclu�do in�cio da formata��o em negrito
'        If InStr(Texto, "LINHA_EM_NEGRITO") > 0 Then
'          Printer.FontBold = True
'        End If
'
'        Printer.Print Str_Impre
'
'        '16/08/2002 - mpdea
'        'T�rmino da formata��o em negrito
'        If InStr(Texto, "LINHA_EM_NEGRITO") > 0 Then
'          Printer.FontBold = False
'        End If
'
'      End If
'    End If
'  Loop Until Final = True

      

'  Close #nFileNum
  
  
  Imprime_Nota = 0
  
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
  rsCor.Close
  Set rsCor = Nothing
  rsTamanho.Close
  Set rsTamanho = Nothing
  
  '06/05/2007 - Anderson
  'Implementa��o da impress�o para CFOP's de servi�os
  Set rstCFOP = Nothing
      
  '----------------------------------------------------------------------------
  'Avisos
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
  '----------------------------------------------------------------------------
  
  
  Exit Function
 
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao imprimir Nota usando o arquivo de configura��o: """ & Nome_Nota & """."
  If Err.Number = 53 Then
    gsMsg = gsMsg & vbCrLf & "Arquivo de configura��o n�o encontrado."
  Else
    gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  End If
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetPrinterName ("REL")
  Resume Next
  Imprime_Nota = 9
  On Error Resume Next
  Close '#nFileNum
  Exit Function
  
End Function

Public Function Imprime_Nota_Entrada(ByVal Nome_Nota As String, ByVal Filial As Integer, ByVal Sequ�ncia As Long) As Integer
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
  Dim Extenso_Tot As String
  Dim Conta_Fat As Integer
  Dim Val_Imposto As Double
  Dim nFileNum As Integer
  Dim nCtLin As Integer
  Dim nComprPag As Integer
  Dim sParte As String
  Dim sAuxGrade As String
  Dim nCor As Long
  Dim sNomeCor As String
  Dim nTamanho As Long
  Dim sNomeTamanho As String
  Dim intVolumagem As Integer
  
  
  Set rsEntradas = db.OpenRecordset("Entradas", , dbReadOnly)
  Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
  Set rsFuncionarios = db.OpenRecordset("Funcion�rios", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  Set rsOp_Entrada = db.OpenRecordset("Opera��es Entrada", , dbReadOnly)
  Set rsCaixas = db.OpenRecordset("Caixas em Uso", , dbReadOnly)
  Set rsEntradas_Prod = db.OpenRecordset("Entradas - Produtos", , dbReadOnly)
  Set rsPesquisa1 = db.OpenRecordset("Pesquisa 1", , dbReadOnly)
  Set rsPesquisa2 = db.OpenRecordset("Pesquisa 2", , dbReadOnly)
  Set rsPesquisa3 = db.OpenRecordset("Pesquisa 3", , dbReadOnly)
  Set rsCor = db.OpenRecordset("Cores", , dbReadOnly)
  Set rsTamanho = db.OpenRecordset("Tamanhos", , dbReadOnly)
  
  '20/11/2006 - Anderson
  'Vari�vel setada para evitar que ocorra erro quando a procedure buscar as mensagens da nota.
  Set m_clcMensagens = New Collection
  
  On Error GoTo ErrHandler
  
  SetPrinterName ("NOTA")
  
  gsInitPrinter = ""
'  Call ResetPrinter
  
  nFileNum = FreeFile
  Open Nome_Nota For Input As #nFileNum
  
  Input #nFileNum, Texto
  If Left(Texto, 22) <> "*** Configura��es Nota" Then
    gsTitle = LoadResString(201)
    gsMsg = "Layout do cabe�alho do arquivo de configura��o """ & Nome_Nota & """ diferente do esperado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Close #nFileNum
    Imprime_Nota_Entrada = 1
    Exit Function
  End If
  
  nComprPag = 0
  sParte = UCase(Trim(Mid(Texto, 74, 4)))
  If Len(sParte) > 0 Then
    If sParte <> "N�O" And sParte <> "LIN" Then
      If Not IsNumeric(sParte) Then
        DisplayMsg "Valor para par�metro de comprimento da p�gina pode ser: N�O, LIN ou <99> (inteiro dois digitos)."
        Imprime_Nota_Entrada = 3
        Close #nFileNum
        Exit Function
      End If
      If Val(sParte) <= 0 Or Val(sParte) > 20 Then
        DisplayMsg "Comprimento da p�gina em polegadas inv�lido."
        Imprime_Nota_Entrada = 3
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
        Open Nome_Nota For Input As #nFileNum
        Input #nFileNum, Texto
      End If
    End If
  End If

  If Mid(Texto, 40, 3) = "SIM" Then
    If SetCompressPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel usar compress�o na impressora solicitada pelo arquivo de configura��o: """ & Nome_Nota & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Close #nFileNum
      Imprime_Nota_Entrada = 1
      SetPrinterName ("REL")
      Exit Function
    End If
  End If
  
  If Mid(Texto, 55, 3) = "SIM" Then
    If SetOitavoPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel ajustar a impressora para 1/8 solicitada pelo arquivo de configura��o: """ & Nome_Nota & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Close #nFileNum
      Imprime_Nota_Entrada = 1
      SetPrinterName ("REL")
      Exit Function
    End If
  End If
  
  If sParte = "LIN" Then
    'Calcule o comprimento pagina Em polegadas
    If Mid(Texto, 55, 3) = "SIM" Then
      nComprPag = nCtLin \ 8
    Else
      nComprPag = nCtLin \ 6
    End If
  End If
  If nComprPag > 0 Then
    If SetComprimPagPrinter(Filial, nComprPag) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel alterar o comprimento de p�gina na impressora."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Imprime_Nota_Entrada = 4
      SetPrinterName ("REL")
      Close #nFileNum
      Exit Function
    End If
  End If
  
  Call SetPrinterCommand(gsInitPrinter)
  
  rsEntradas.Index = "Sequ�ncia"
  rsEntradas.Seek "=", Filial, Sequ�ncia
  If rsEntradas.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "N�o � poss�vel imprimir a nota, registro de entrada n�o encontrado para Filial= " & Filial & ", " & "Seq��ncia= " & Sequ�ncia
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Close #nFileNum
    Exit Function
  End If
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Filial
  If rsParametros.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "N�o � poss�vel imprimir a nota, par�metros n�o encontrados para Filial=" & Filial
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Close #nFileNum
    Exit Function
  End If
  
  rsOp_Entrada.Index = "C�digo"
  rsOp_Entrada.Seek "=", rsEntradas("Opera��o")
  If rsOp_Entrada.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Opera��o de Entrada referida pelo registro de Entradas n�o foi localizada: Opera��o=" & rsEntradas("Opera��o")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Close #nFileNum
    Exit Function
  End If
  
  rsCliFor.Index = "C�digo"
  rsCliFor.Seek "=", rsEntradas("Fornecedor")
  If rsCliFor.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Fornecedor referido pelo registro de Entradas n�o foi localizado: Fornecedor=" & rsEntradas("Fornecedor")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Close #nFileNum
    Exit Function
  End If
  
  Glob_gnCodFilial = Filial
  Glob_Nome_Filial = rsParametros("Nome") & ""
  Glob_Data = Format(rsEntradas("Data"), "dd/mm/yyyy")
  Glob_Data_Sa�da = Format(Date, "dd/mm/yyyy")
  Glob_Hora_Sa�da = Format(Time, "hh:mm:ss")
  Glob_Cod_Opera��o = rsEntradas("Opera��o")
  Glob_Nome_Opera��o = rsOp_Entrada("Nome") & ""
  Glob_C�digo_Fiscal = rsOp_Entrada("C�digo Fiscal") & ""
  Glob_Sequ�ncia = rsEntradas("Sequ�ncia")
  Glob_Cod_Vendedor = rsEntradas("Digitador")
  '19/05/2005 - Daniel
  '
  'Solicitante: Ped�gio - Esta otimiza��o est� dispon�vel
  '             para todos usu�rios do Quick Store
  '
  'O sistema dever� julgar se a nota fiscal foi criada
  'automaticamente ou manualmente a partir da opera��o escolhida
  'Nota: Caso tenha sido manualmente (bloquinhos) mostraremos o
  'campo Entradas.[Nota Fiscal] ao inv�s de Entradas.[Nota Impressa]
  If gbNotaManual(rsEntradas("Opera��o"), "ENTRADA") Then
    Glob_Nota_Impressa = CLng("0" & rsEntradas("Nota Fiscal") & "")
  Else
    Glob_Nota_Impressa = rsEntradas("Nota Impressa")
  End If
  
  rsFuncionarios.Index = "C�digo"
  rsFuncionarios.Seek "=", Glob_Cod_Vendedor
  If rsFuncionarios.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Digitador referido pelo registro de Entradas n�o foi localizado: Digitador=" & rsEntradas("Digitador")
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    SetPrinterName ("REL")
    Close #nFileNum
    Exit Function
  End If
  
  '20/11/2006 - Anderson
  'Inclu�do o campo apelido do funcion�rio.
  'Solicitante: Technomax
  Glob_Nome_Vendedor = rsFuncionarios("Nome") & ""
  Glob_Apelido = rsFuncionarios("Apelido") & ""
  
  Glob_C�digo_Cli = rsEntradas("Fornecedor")
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
  Glob_Cod_Caixa = rsEntradas("Caixa")
  Glob_Nome_Caixa = rsCaixas("Descri��o") & ""
  Glob_Tab_Pre�o = ""
  Glob_RefrmInterna = ""
  Glob_Obs_Mov = rsEntradas("Observa��es") & ""
  
  
  '  gsObsDoc(0) = frmObsNota.Obs1.Text & ""
  '  gsObsDoc(1) = frmObsNota.Obs2.Text & ""
  '  gsObsDoc(2) = frmObsNota.Obs3.Text & ""
  '  gsObsDoc(3) = frmObsNota.Obs4.Text & ""
  '  gsObsDoc(4) = frmObsNota.Obs5.Text & ""
  '  gsObsDoc(5) = frmObsNota.Obs6.Text & ""
  '  gsObsDoc(6) = frmObsNota.Obs7.Text & ""
  '  gsObsDoc(7) = frmObsNota.Obs8.Text & ""
  '  gsTransportadora = frmObsNota.Nome_Transp.Text & ""
  '  gsPlaca = frmObsNota.Placa.Text & ""
  '  gsUfrmPlaca = frmObsNota.UfrmPlaca.Text & ""
  '  gsQtdeTrans = frmObsNota.Qtde.Text & ""
  '  gsMarcaTrans = frmObsNota.Marca.Text & ""
  '  gsEspecieTrans = frmObsNota.Esp�cie.Text & ""
  '  gsPesoBruto = frmObsNota.Bruto.Text & ""
  '  gsPesoLiquido = frmObsNota.L�quido.Text & ""
  '  If frmObsNota.O_Destinat�rio.Value = True Then gsFretePago = "2"
  '  If frmObsNota.O_Destinat�rio.Value = False Then gsFretePago = "1"
  '
  
  '
  Glob_Base_ICM = rsEntradas("Base ICM")
  Glob_Valor_ICM = rsEntradas("Valor ICM")
  Glob_Base_ICM_Sub = rsEntradas("Base ICM Subs")
  Glob_Valor_ICM_Sub = rsEntradas("Valor ICM Subs")
  Glob_Total_Produto = rsEntradas("Produtos")
  Glob_Total_Desconto = rsEntradas("Desconto")
  Glob_Frete = rsEntradas("Frete")
  Glob_IPI = rsEntradas("IPI")
  Glob_Total_Pagar = rsEntradas("Total")
  
  Glob_Total_Servi�o = 0
  Glob_Total_ISS = 0
  
  '30/01/2004 - Daniel
  'Tratando as vars de impostos requeridos
  'g_dblTotalCSLL = Format((CDbl((Glob_Total_Servi�o * g_dblPercentualCSLL) / 100)), FORMAT_VALUE)
  'g_dblTotalCOFINS = Format((CDbl((Glob_Total_Servi�o * g_dblPercentualCOFINS) / 100)), FORMAT_VALUE)
  'g_dblTotalPIS = Format((CDbl((Glob_Total_Servi�o * g_dblPercentualPIS) / 100)), FORMAT_VALUE)
  'g_dblTotalIRRF = Format((CDbl((Glob_Total_Servi�o * g_dblPercentualIRRF) / 100)), FORMAT_VALUE)
  '------------------------------------------------------------------------------------------------
  
  
  Extenso_Tot = Extenso(rsEntradas("Total"))
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
  
  '19/12/2007 - Anderson
  'Implementa��o do NSU
  NSU = Format("" & rsEntradas("NSU"), "0000000000")
  NSU_Data = Format("" & rsEntradas("NSU_Data"), "dd/mm/yy")
  NSU_Hora = Format("" & rsEntradas("NSU_Hora"), "hh:nn")
    
  Rem Monta tabela dos produtos
  Erase Tab_Prod
  
  i = 0
  Linha = 0
  Glob_Conta_Prod = 0
  
  rsPesquisa1.Index = "C�digo"
  rsPesquisa2.Index = "C�digo"
  rsPesquisa3.Index = "C�digo"
  rsProdutos.Index = "C�digo"
  rsEntradas_Prod.Index = "Sequ�ncia"
  
Lp_Prod:
  rsEntradas_Prod.Seek ">", rsEntradas("Filial"), rsEntradas("Sequ�ncia"), Linha
  If rsEntradas_Prod.NoMatch Then GoTo Ve_Servi�os
  If rsEntradas("Filial") <> rsEntradas_Prod("Filial") Then GoTo Ve_Servi�os
  If rsEntradas("Sequ�ncia") <> rsEntradas_Prod("Sequ�ncia") Then GoTo Ve_Servi�os
  Linha = rsEntradas_Prod("Linha")
  
  Nome_Pesq1 = ""
  Nome_Pesq2 = ""
  Nome_Pesq3 = ""
  
  
  '----------------------------------------------------------------------------
  '28/11/2003 - mpdea
  'Corrigido RT-94 (Invalid use of null)
  rsProdutos.Seek "=", rsEntradas_Prod("C�digo Sem Grade") & ""
  If rsProdutos.NoMatch Then GoTo Lp_Prod
  
  If rsProdutos("Pesquisa 1") <> 0 Then
    rsPesquisa1.Seek "=", rsProdutos("Pesquisa 1")
    If Not rsPesquisa1.NoMatch Then Nome_Pesq1 = rsPesquisa1("Nome") & ""
  End If
  If rsProdutos("Pesquisa 2") <> 0 Then
    rsPesquisa1.Seek "=", rsProdutos("Pesquisa 2")
    If Not rsPesquisa2.NoMatch Then Nome_Pesq2 = rsPesquisa2("Nome") & ""
  End If
  If rsProdutos("Pesquisa 3") <> 0 Then
    rsPesquisa3.Seek "=", rsProdutos("Pesquisa 3")
    If Not rsPesquisa1.NoMatch Then Nome_Pesq3 = rsPesquisa3("Nome") & ""
  End If
  '----------------------------------------------------------------------------
  
  
  Tab_Prod(i).C�digo = rsProdutos("C�digo")
  Tab_Prod(i).C�digo_Prod_Forn = rsProdutos("C�digo do Fornecedor") & ""
  Tab_Prod(i).Nome = rsProdutos("Nome Nota") & ""
  If rsProdutos("Nome Nota") = "" Then Tab_Prod(i).Nome = rsProdutos("Nome") & ""
  
  '06/09/2002 - mpdea
  'Inclu�do os campos para impress�o espec�fica do nome do produto como
  'est� no campo Nome do cadastro ou o campo Nome para nota (Fixo)
  '
  '06/07/2005 - Daniel
  'Corre��o: A vari�vel correta para a incrementa��o do Array (Tab_Prod)
  '� "i" e n�o a global "gnCtItemProd" da Imprime_Nota (Sa�das)
  Tab_Prod(i).NomeCadastro = rsProdutos("Nome") & ""
  Tab_Prod(i).NomeNota = rsProdutos("Nome Nota") & ""
  
  Tab_Prod(i).C_Fiscal = rsProdutos("Classifica��o Fiscal") & ""
  
  '22/09/2005 - mpdea
  'Tratamento da impress�o da Descri��o da Classifica��o Fiscal
  If IsNumeric(rsProdutos("Classifica��o Fiscal")) And rsProdutos("Classifica��o Fiscal") <> 0 Then
    Set rstClassFiscal = db.OpenRecordset("SELECT * FROM [Classifica��o Fiscal] WHERE C�digo = " & rsProdutos("Classifica��o Fiscal"), dbOpenDynaset)
    
    With rstClassFiscal
      If Not (.BOF And .EOF) Then
        .MoveFirst
        '20/02/2009 - mpdea
        'Corrigido uso de �ndice incorreto
        Tab_Prod(i).DescricaoClassificaoFiscal = .Fields("Nome").Value & ""
      End If
      .Close
    End With
    
    Set rstClassFiscal = Nothing
  End If
  
  
  Tab_Prod(i).S_Tribut�ria = rsProdutos("Situa��o Tribut�ria") & ""
  Tab_Prod(i).Unid = rsProdutos("Unidade Venda") & ""
  Tab_Prod(i).Qtde = rsEntradas_Prod("Qtde")
  Tab_Prod(i).Valor_Unit = rsEntradas_Prod("Pre�o")
  Tab_Prod(i).Valor_Total = rsEntradas_Prod("Pre�o") * rsEntradas_Prod("Qtde")
  Tab_Prod(i).Desconto_Perc = rsEntradas_Prod("Desconto")
  Tab_Prod(i).Aliq_ICM = rsEntradas_Prod("ICM")
  
  Val_Imposto = rsEntradas_Prod("Pre�o") - (rsEntradas_Prod("Desconto") * rsEntradas_Prod("Pre�o") / 100)
  Val_Imposto = Val_Imposto * rsEntradas_Prod("Qtde")
  Val_Imposto = Val_Imposto * rsEntradas_Prod("ICM") / 100
  Val_Imposto = Format(Val_Imposto, "############0.00")
  Tab_Prod(i).Valor_ICM = Val_Imposto
  
  Tab_Prod(i).Aliq_IPI = rsEntradas_Prod("IPI")
  
  Val_Imposto = rsEntradas_Prod("Pre�o") - (rsEntradas_Prod("Desconto") * rsEntradas_Prod("Pre�o") / 100)
  Val_Imposto = Val_Imposto * rsEntradas_Prod("Qtde")
  Val_Imposto = Val_Imposto * rsEntradas_Prod("IPI") / 100
  Val_Imposto = Format(Val_Imposto, "############0.00")
  Tab_Prod(i).Valor_IPI = Val_Imposto
  
  Tab_Prod(i).Valor_Final = rsEntradas_Prod("Pre�o Final")
  Tab_Prod(i).Pesq1 = rsProdutos("Pesquisa 1")
  Tab_Prod(i).Pesq2 = rsProdutos("Pesquisa 2")
  Tab_Prod(i).Pesq3 = rsProdutos("Pesquisa 3")
  Tab_Prod(i).Nome_Pesq1 = Nome_Pesq1
  Tab_Prod(i).Nome_Pesq2 = Nome_Pesq2
  Tab_Prod(i).Nome_Pesq3 = Nome_Pesq3
  Tab_Prod(i).Local = rsProdutos("Localiza��o") & ""
  
 sAuxGrade = ""
 nCor = 0
 nTamanho = 0
 sNomeCor = ""
 sNomeTamanho = ""
 
  If rsEntradas_Prod("C�digo") <> rsEntradas_Prod("C�digo Sem Grade") Then
     sAuxGrade = Right(rsEntradas_Prod("C�digo"), 6)
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
         
     Tab_Prod(i).Cor = nCor & ""
     Tab_Prod(i).Nome_Cor = sNomeCor & ""
     Tab_Prod(i).Tamanho = nTamanho & ""
     Tab_Prod(i).Nome_Tamanho = sNomeTamanho & ""
  End If
    
 
  '27/09/2004 - mpdea
  'Inclu�do o campo de Volumagem por Quantidade
  With Tab_Prod(i)
    Call IsDataType(dtInteger, rsProdutos.Fields("Volumagem").Value, intVolumagem)
    If intVolumagem > 0 Then
      .VolumagemQtde = "(" & Format(.Qtde \ intVolumagem, "000") & "/" & Format(.Qtde Mod intVolumagem, "000") & ")"
    End If
  End With
      
  
  i = i + 1
  
  GoTo Lp_Prod
  
Ve_Servi�os:
  Erase Tab_Serv
  Glob_Conta_Serv = 0
  
  
Ve_Fatura:
  
  Erase Tab_Fat
  
  Conta_Fat = 0
  Linhas = 0
  
  
Final:
  Final = False
  Do
    DoEvents
    Input #nFileNum, Texto
    If Texto = "*** Fim de arquivo ***" Then Final = True
    If Final = False Then
      Texto = Apaga_Aspas(Texto)
      Final_Linha = False
      If Len(Texto) < 3 Then
        MsgBox "Arquivo de configura��o """ & Nome_Nota & """ inv�lido."
        Imprime_Nota_Entrada = 1
        Exit Function
      End If
      Especial2 = False
      If Left(Texto, 13) = "[LINHA_BRANCO" Then
        Especial2 = True
        Linhas = Val(Mid(Texto, 15))
        Do
          Printer.Print
          Linhas = Linhas - 1
        Loop Until Linhas = 0
      End If
      If Especial2 = False Then
        Str_Impre = Retorna_Texto(Texto)
        
        '16/08/2002 - mpdea
        'Inclu�do in�cio da formata��o em negrito
        If InStr(Texto, "LINHA_EM_NEGRITO") > 0 Then
          Printer.FontBold = True
        End If
        
        Printer.Print Str_Impre
      
        '16/08/2002 - mpdea
        'T�rmino da formata��o em negrito
        If InStr(Texto, "LINHA_EM_NEGRITO") > 0 Then
          Printer.FontBold = False
        End If
        
      End If
    End If
  Loop Until Final = True
  
  Close #nFileNum
  
  Printer.Print
  Printer.EndDoc
  SetPrinterName ("REL")
  
  rsEntradas.Close
  Set rsEntradas = Nothing
  rsCliFor.Close
  Set rsCliFor = Nothing
  rsFuncionarios.Close
  Set rsFuncionarios = Nothing
  rsProdutos.Close
  Set rsProdutos = Nothing
  rsParametros.Close
  Set rsParametros = Nothing
  rsOp_Entrada.Close
  Set rsOp_Entrada = Nothing
  rsCaixas.Close
  Set rsCaixas = Nothing
  rsEntradas_Prod.Close
  Set rsEntradas_Prod = Nothing
  rsPesquisa1.Close
  Set rsPesquisa1 = Nothing
  rsPesquisa2.Close
  Set rsPesquisa2 = Nothing
  rsPesquisa3.Close
  Set rsPesquisa3 = Nothing
  rsCor.Close
  Set rsCor = Nothing
  rsTamanho.Close
  Set rsTamanho = Nothing

  
  
  
  
  
  Exit Function
 
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao imprimir Nota usando o arquivo de configura��o: """ & Nome_Nota & """."
  If Err.Number = 53 Then
    gsMsg = gsMsg & vbCrLf & "Arquivo de configura��o n�o encontrado."
  Else
    gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  End If
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetPrinterName ("REL")
  Imprime_Nota_Entrada = 9
  On Error Resume Next
  Close #nFileNum
  Exit Function

End Function

Function Retorna_Texto(Texto As String) As String
  Dim Aux, Parte, Letra, Campo As String
  Dim Tam, Pos As Integer
  Dim Texto_Final As String
  Dim Aspas As Integer
  
  Aux = Texto
  
  Tam = Len(Aux)
  Pos = 1
  Parte = ""
  Aspas = False
  Do
    Letra = Mid(Aux, Pos, 1)
    If Letra = "}" Then
      Texto_Final = Texto_Final + Parte
      Parte = ""
    End If
    If Letra = "]" Then
      Campo = Pega_Campo(CStr(Parte))
      Texto_Final = Texto_Final + Campo
      Parte = ""
    End If
    If Letra <> "[" And Letra <> "]" And Letra <> "{" And Letra <> "}" Then
       Parte = Parte & Letra
    End If
    Pos = Pos + 1
  Loop Until Pos > Tam
    
  If Left(Texto_Final, 1) = "-" And Len(Trim(Texto_Final)) > 1 Then
    Texto_Final = " " & Mid(Texto_Final, 2, Len(Texto_Final) - 1)
  End If
  Retorna_Texto = Texto_Final
  
End Function

Function Pega_Campo(Parte As String) As String
  Dim Retorno As String
  Dim Tamanho As Integer
  Dim �_N�mero As Integer
  Dim �_Valor As Integer
  Dim �_Texto As Integer
  Dim �_Zeros As Integer
  Dim Pos, Tam, Pega_Tam As Integer
  Dim Letra, Texto, Texto2, Aux As String
 
 Const Brancos = "                                                                               "
 
 Texto = ""
 Texto2 = ""
 Retorno = ""
 Pega_Tam = False
 �_N�mero = False
 �_Valor = False
 �_Texto = False
 �_Zeros = False
 
 Tam = Len(Parte)
 For Pos = 1 To Tam
   Letra = Mid(Parte, Pos, 1)
   If Pega_Tam = True Then
     Texto2 = Texto2 + Letra
   End If
   If Pega_Tam = False Then
     If Letra <> "," Then Texto = Texto + Letra
     If Letra = "," Then Pega_Tam = True
   End If
 Next Pos
 
 Tamanho = Val(Texto2)
 
 
 Texto = UCase(Texto)
 
 If Texto = "PROXIMO_PRODUTO" Then
   Retorno = " "       'LINHA DA ALTERA��O
   Glob_Conta_Prod = Glob_Conta_Prod + 1
   If Glob_Conta_Prod > 500 Then Glob_Conta_Prod = 500
   �_Texto = True
   Tamanho = 0
 End If
 If Texto = "PROXIMA_FATURA" Then
   Retorno = " "       'LINHA DA ALTERA��O
   Glob_Conta_Fat = Glob_Conta_Fat + 1
   If Glob_Conta_Fat > 50 Then Glob_Conta_Fat = 50
   �_Texto = True
   Tamanho = 0
 End If
 If Texto = "PROXIMO_SERVI�O" Then
   Retorno = " "       'LINHA DA ALTERA��O
   Glob_Conta_Serv = Glob_Conta_Serv + 1
   If Glob_Conta_Serv > 50 Then Glob_Conta_Serv = 50
   �_Texto = True
   Tamanho = 0
 End If


 
  
 
 If Texto = "RESUMO DO PAGAMENTO" Then
   Retorno = GLOB_RESUMO_PAGTO
   �_Texto = True
   Tamanho = 0
 End If
 
 
 If Texto = "C�DIGO FILIAL" Then
   Retorno = Glob_gnCodFilial
   �_N�mero = True
 End If
  
 If Texto = "NOME FILIAL" Then
   Retorno = Glob_Nome_Filial
   �_Texto = True
 End If
 
 If Texto = "DATA" Then
   Retorno = Glob_Data
   �_Texto = True
 End If
 
 If Texto = "DATA SA�DA" Then
   Retorno = Glob_Data_Sa�da
   �_Texto = True
 End If
 
 If Texto = "DATA EMISS�O CONTA" Then
   Retorno = Glob_Data_Emiss�o
   �_Texto = True
 End If
 
 If Texto = "HORA SA�DA" Then
   Retorno = Glob_Hora_Sa�da
   �_Texto = True
 End If
 
 If Texto = "C�DIGO OPERA��O" Then
   Retorno = Glob_Cod_Opera��o
   �_Texto = True
 End If
   
 If Texto = "NOME OPERA��O" Then
   Retorno = Glob_Nome_Opera��o
   �_Texto = True
 End If
   
 If Texto = "C�DIGO FISCAL" Then
   Retorno = Glob_C�digo_Fiscal
   �_Texto = True
 End If
 
 '----------------------------------------------------
 '06/05/2007 - Anderson
 'Implementa��o da impress�o para CFOP's de servi�os
 'C�digo Fiscal Completo (Opera��o + Itens)
 If Texto = "C�DIGO FISCAL COMPLETO (OPERA��O + ITENS)" Then
   Retorno = Glob_C�digo_Fiscal_Completo
   �_Texto = True
 End If
 
 If Mid(Texto, 1, 18) = "C�DIGO FISCAL ITEM" Then
   Retorno = Glob_C�digo_Fiscal_Item(Mid(Texto, 20, 1) - 1)
   �_Texto = True
 End If
 '----------------------------------------------------
 
  '----------------------------------------------------
  '24/04/2008 - mpdea
  'Descri��o e total por CFOP relacionado a movimenta��o
  If Mid(Texto, 1, 29) = "NOME OPERA��O - C�DIGO FISCAL" Then
    Retorno = Glob_Nome_Operacao_CFOP(Mid(Texto, 31, 1) - 1)
    �_Texto = True
  End If
  If Mid(Texto, 1, 27) = "VALOR TOTAL - C�DIGO FISCAL" Then
    Retorno = Glob_Total_CFOP(Mid(Texto, 29, 1) - 1)
    �_Valor = True
  End If
  '----------------------------------------------------
 
 If Texto = "SEQ��NCIA" Then
   Retorno = Glob_Sequ�ncia
   �_Texto = True
 End If
 
 If Texto = "C�DIGO VENDEDOR" Then
   Retorno = Glob_Cod_Vendedor
   �_N�mero = True
 End If
 
 If Texto = "NOME VENDEDOR" Then
   Retorno = Glob_Nome_Vendedor & ""
   �_Texto = True
 End If
 
 '20/11/2006 - Anderson
 'Inclu�do o campo apelido do funcion�rio.
 'Solicitante: Technomax
 If Texto = "APELIDO" Then
   Retorno = Glob_Apelido & ""
   �_Texto = True
 End If
 
 '-------------------------------------------------------------------
 '07/08/2003 - mpdea
 'Inclu�do C�digo e Nome do T�cnico
 If Texto = "C�DIGO T�CNICO" Then
   Retorno = g_intCodigoTecnico
   �_N�mero = True
 End If
 
 If Texto = "NOME T�CNICO" Then
   Retorno = g_strNomeTecnico
   �_Texto = True
 End If
 '-------------------------------------------------------------------

 '-------------------------------------------------------------------
 '08/01/2004 - Daniel
 'Inclu�do Valor Recebido e Troco da
 'tabela de Sa�das
  If Texto = "VALOR RECEBIDO VENDA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Format(g_dblValorRecebido, FORMAT_VALUE)
      �_N�mero = True
    End If
  End If
 
  If Texto = "TROCO" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Format(g_dblTroco, FORMAT_VALUE)
      �_N�mero = True
    End If
  End If
  '-------------------------------------------------------------------
 
  '-------------------------------------------------------------------
  '28/01/2004 - mpdea
  'Formatado o valor
  '
  '09/01/2004 - Daniel
  'Inclu�do a soma da Qtde de Itens
  If Texto = "SOMA DA QTDE DE ITENS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      If CLng(g_sngQtdeItens) = CSng(g_sngQtdeItens) Then
        Retorno = g_sngQtdeItens
      Else
        Retorno = Format(g_sngQtdeItens, "#,###,###,##0.###")
      End If
    '    Retorno = Format(g_sngQtdeItens, FORMAT_VALUE)
      �_N�mero = True
    End If
  End If
  '-------------------------------------------------------------------
  
  '-------------------------------------------------------------------
  '30/01/2004 - Daniel
  'Inclu�do Vars do tratamento de
  'Impostos
  If Texto = "PERCENTUAL CSLL" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Format(g_dblPercentualCSLL, FORMAT_VALUE)
      �_N�mero = True
    End If
  End If
  
  If Texto = "PERCENTUAL COFINS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Format(g_dblPercentualCOFINS, FORMAT_VALUE)
      �_N�mero = True
    End If
  End If
  
  If Texto = "PERCENTUAL PIS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Format(g_dblPercentualPIS, FORMAT_VALUE)
      �_N�mero = True
    End If
  End If
  
  If Texto = "PERCENTUAL IRRF" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Format(g_dblPercentualIRRF, FORMAT_VALUE)
      �_N�mero = True
    End If
  End If
  
  '-------------------------------------------------------------------
  If Texto = "TOTAL CSLL" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblTotalCSLL
      �_Valor = True
      �_Texto = False
    End If
  End If
  
  If Texto = "TOTAL COFINS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblTotalCOFINS
      �_Valor = True
      �_Texto = False
    End If
  End If
  
  If Texto = "TOTAL PIS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblTotalPIS
      �_Valor = True
      �_Texto = False
    End If
  End If
  
  If Texto = "TOTAL IRRF" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblTotalIRRF
      �_Valor = True
      �_Texto = False
    End If
  End If
  '-------------------------------------------------------------------
  
  '27/04/2005 - Daniel
  'Tratamento para o campo Seguro da table Sa�das
  If Texto = "SEGURO" Then
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblSeguro
      �_Valor = True
      �_Texto = False
    End If
  End If
  
  '-------------------------------------------------------------------
  '13/04/2004 - Daniel
  'Tratamento para: N�mero da Autoriza��o, M�s X, Endere�o Cob,
  'Complemento Cob, Bairro Cob, Cidade Cob, Estado Cob e CEP Cob
  If Texto = "N�MERO DA AUTORIZA��O" Then
    Retorno = g_lngNumAutorizacao
    �_Valor = False
    �_Texto = True
  End If
  
  If Texto = "M�S X" Then
    Retorno = g_intMesX
    �_Valor = False
    �_Texto = True
  End If
  
  If Texto = "ENDERE�O COB" Then
    Retorno = g_strEnderecoCob
    �_Valor = False
    �_Texto = True
  End If
  
  If Texto = "COMPLEMENTO COB" Then
    Retorno = g_strComplementoCob
    �_Valor = False
    �_Texto = True
  End If
  
  If Texto = "BAIRRO COB" Then
    Retorno = g_strBairroCob
    �_Valor = False
    �_Texto = True
  End If
  
  If Texto = "CIDADE COB" Then
    Retorno = g_strCidadeCob
    �_Valor = False
    �_Texto = True
  End If
  
  If Texto = "ESTADO COB" Then
    Retorno = g_strEstadoCob
    �_Valor = False
    �_Texto = True
  End If
  
  If Texto = "CEP COB" Then
    Retorno = g_strCepCob
    �_Valor = False
    �_Texto = True
  End If
  '-------------------------------------------------------------------
  
  '06/05/2004 - Daniel
  'Adi��o do campo ObsIsentoIPI da tabela Cli_For
  If Texto = "OBS ISEN��O IPI" Then
    Retorno = g_strObsIsentoIPI
    �_Valor = False
    �_Texto = True
  End If
  
  '17/05/2004 - Daniel
  'Adi��o do campo Diferimento.ObsDiferimento
  If Texto = "OBS DIFERIMENTO" Then
    Retorno = g_strObsDiferimento
    �_Valor = False
    �_Texto = True
  End If
  
  '31/05/2007 - Anderson
  If Texto = "OR�AMENTO APROVADO POR" Then
    Retorno = Glob_Aprovado
    �_Texto = True
  End If
  
  '31/05/2007 - Anderson
  If Texto = "PROMETIDO PARA" Then
    Retorno = Glob_Prometido
    �_Texto = True
  End If
  
  If Texto = "C�DIGO CLIENTE" Then
    Retorno = Glob_C�digo_Cli
    �_N�mero = True
  End If
  
  If Texto = "NOME CLIENTE" Then
    Retorno = Glob_Nome
    �_Texto = True
  End If
  
  If Texto = "FANTASIA" Then
    Retorno = Glob_Fantasia
    �_Texto = True
  End If
  
  If Texto = "ENDERE�O" Then
    Retorno = Glob_Endere�o
    �_Texto = True
  End If
  
  '23/10/2009 - mpdea
  'N�mero do endere�o
  If Texto = "N�MERO ENDERE�O" Then
    Retorno = Glob_NumeroEndereco
    �_Texto = True
  End If
  
  If Texto = "COMPLEMENTO" Then
    Retorno = Glob_Complemento
    �_Texto = True
  End If
  
  If Texto = "BAIRRO" Then
    Retorno = Glob_Bairro
    �_Texto = True
  End If
  
  If Texto = "CEP" Then
    Retorno = Glob_CEP
    �_Texto = True
  End If
  
  If Texto = "CIDADE" Then
    Retorno = Glob_Cidade
    �_Texto = True
  End If
  
  If Texto = "FONE1" Then
    Retorno = Glob_Fone1
    �_Texto = True
  End If
  
  If Texto = "FONE2" Then
    Retorno = Glob_Fone2
    �_Texto = True
  End If
  
  If Texto = "ESTADO" Then
    Retorno = Glob_Estado
    �_Texto = True
  End If
  
  If Texto = "C�DIGO CAIXA" Then
    Retorno = Glob_Cod_Caixa
    �_N�mero = True
  End If
  
  If Texto = "NOME CAIXA" Then
    Retorno = Glob_Nome_Caixa
    �_Texto = True
  End If
  
  If Texto = "TABELA PRE�O" Then
    Retorno = Glob_Tab_Pre�o
    �_Texto = True
  End If
  
  If Texto = "REFER�NCIA INTERNA" Then
    Retorno = Glob_RefrmInterna
    �_Texto = True
  End If
  
  If Texto = "OBSERVA��ES" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = Glob_Obs_Mov
      �_Texto = True
    End If
  End If
  
  '15/08/2002 - mpdea
  'Inclu�do o campo de informa��es sobre o or�amento (n�mero do or�amento e terminal)
  If Texto = "N�MERO DO OR�AMENTO" Then
    Retorno = gstrInfoNrOrcamento
    �_Texto = True
  End If
  
  If Texto = "N�MERO NOTA" Then
    Retorno = Glob_Nota_Impressa & ""
    �_Zeros = True
  End If
  If Texto = "FRETE CONTA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsFretePago
      �_Texto = True
    End If
  End If
  If Texto = "OBS1" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(0)
      �_Texto = True
    End If
  End If
  If Texto = "OBS2" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(1)
      �_Texto = True
    End If
  End If
  If Texto = "OBS3" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(2)
      �_Texto = True
    End If
  End If
  If Texto = "OBS4" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(3)
      �_Texto = True
    End If
  End If
  If Texto = "OBS5" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(4)
      �_Texto = True
    End If
  End If
  If Texto = "OBS6" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(5)
      �_Texto = True
    End If
  End If
  If Texto = "OBS7" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(6)
      �_Texto = True
    End If
  End If
  If Texto = "OBS8" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsObsDoc(7)
      �_Texto = True
    End If
  End If
  
  
  '----------------------------------------------------------------------------
  '01/02/2006 - mpdea
  'Mensagem para Nota Fiscal
  If Texto = "MENSAGEMNOTAFISCAL1" Then
    'Tratamento para impress�o multinota
    �_Texto = True
    If g_blnPrintMNF Then
      Retorno = ""
    Else
      If m_clcMensagens.Count > 0 Then
        Retorno = m_clcMensagens.Item(1)
      End If
    End If
  End If
  If Texto = "MENSAGEMNOTAFISCAL2" Then
    'Tratamento para impress�o multinota
    �_Texto = True
    If g_blnPrintMNF Then
      Retorno = ""
    Else
      If m_clcMensagens.Count > 1 Then
        Retorno = m_clcMensagens.Item(2)
      End If
    End If
  End If
  If Texto = "MENSAGEMNOTAFISCAL3" Then
    'Tratamento para impress�o multinota
    �_Texto = True
    If g_blnPrintMNF Then
      Retorno = ""
    Else
      If m_clcMensagens.Count > 2 Then
        Retorno = m_clcMensagens.Item(3)
      End If
    End If
  End If
  '----------------------------------------------------------------------------
  
  
  If Texto = "NOME TRANSPORTADORA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsTransportadora
      �_Texto = True
    End If
  End If
  
  If Texto = "CNPJ TRANSPORTADORA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsCNPJTransportadora
      �_Texto = True
    End If
  End If
  
  If Texto = "IE TRANSPORTADORA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsIETransportadora
      �_Texto = True
    End If
  End If
  
  If Texto = "ENDER TRANSPORTADORA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsEnderTransportadora
      �_Texto = True
    End If
  End If
  
  If Texto = "MUNICIPIO TRANSPORTADORA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsMunicipioTransportadora
      �_Texto = True
    End If
  End If
  
  If Texto = "UF TRANSPORTADORA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsUFTransportadora
      �_Texto = True
    End If
  End If
  
  If Texto = "PLACA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsPlaca
      �_Texto = True
    End If
  End If
  
  If Texto = "UF PLACA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsUfrmPlaca
      �_Texto = True
    End If
  End If
  
  If Texto = "QTDE TRANS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsQtdeTrans
      �_Texto = True
    End If
  End If
  
  If Texto = "ESP�CIE TRANS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsEspecieTrans
      �_Texto = True
    End If
  End If
  
  If Texto = "MARCA TRANS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsMarcaTrans
      �_Texto = True
    End If
  End If
  
  If Texto = "PESO BRUTO" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsPesoBruto
      �_Texto = True
    End If
  End If
  
  If Texto = "PESO L�QUIDO" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gsPesoLiquido
      �_Texto = True
    End If
  End If
  
  If Texto = "QTDE ITENS PRODUTO" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gnCtItemProd
      �_Texto = True
    End If
  End If
  
  If Texto = "QTDE ITENS SERVI�O" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = ""
      �_Texto = True
    Else
      Retorno = gnCtItemServ
      �_Texto = True
    End If
  End If
  
  If Texto = "QTDE PARCELAS FATURA" Then
    Retorno = gnCtParcFat
    �_Texto = True
  End If
    
  If Texto = "FANTASIA" Then
    Retorno = Glob_Fantasia
    �_Texto = True
  End If
  
  If Texto = "CGC" Then
    Retorno = Glob_CGC
    �_Texto = True
  End If
  
  If Texto = "INSCRI��O ESTADUAL" Then
    Retorno = Glob_Inscri��o
    �_Texto = True
  End If
  
  If Texto = "C�DIGO PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).C�digo
    End If
  End If
  If Texto = "C�DIGO PRODUTO FORNECEDOR" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).C�digo_Prod_Forn
    End If
  End If
  If Texto = "NOME PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Nome
    End If
  End If
  
  
  '04/09/2002 - mpdea
  'Inclus�o de novos campos
  If Texto = "NOME PRODUTO (CADASTRO)" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).NomeCadastro
    End If
  End If
  If Texto = "NOME PRODUTO (NOTA)" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).NomeNota
    End If
  End If
  
  '29/04/2008 - mpdea
  'CFOP do produto
  If Texto = "CFOP DO PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).CFOP
    End If
  End If

  If Texto = "CLASSIFICA��O FISCAL" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).C_Fiscal
    End If
  End If
  
  If Texto = "DESCRI��O DA CLASS. FISCAL" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      '22/09/2005 - mpdea
      'Corrigido descri��o da classifica��o fiscal
      'Estava armazenando somente a do �ltimo produto
      Retorno = Tab_Prod(Glob_Conta_Prod).DescricaoClassificaoFiscal
      'Retorno = g_strDescrClassFiscal
    End If
  End If
  
  If Texto = "SITUA��O TRIBUT�RIA" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).S_Tribut�ria
    End If
  End If
  If Texto = "UNIDADE VENDA" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Unid
    End If
  End If
  
  '05/05/2011 - mpdea
  'NBM/NCM do produto
  If Texto = "C�DIGO NBM/NCM" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).CodigoNbmNcm <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).CodigoNbmNcm
    End If
  End If
  
  '27/04/2005 - Daniel
  'Tratamento para o Produtos.Fabricante
  If Texto = "FABRICANTE" Then
    Retorno = ""
    �_Texto = True
    Retorno = Tab_Prod(Glob_Conta_Prod).Fabricante & ""
  End If
  
  '29/11/2004 - Daniel
  'Adicionado os campos Lote e Data de Validade
  If Texto = "LOTE" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Lote <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Lote
    End If
  End If
  
  If Texto = "DATA DE VALIDADE" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).DataValidade <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).DataValidade
    End If
  End If
  
  If Texto = "LOCAL" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Local <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Local
    End If
  End If
  
  If Texto = "QTDE" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Qtde
      �_N�mero = True
      �_Texto = False
    End If
  End If
  
  If Texto = "PERC DESCONTO PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Desconto_Perc <> 0 Then
      '17/08/2007 - Anderson
      'altera��o realizada para formatar descontos fracionados.
      'Retorno = Tab_Prod(Glob_Conta_Prod).Desconto_Perc
      Retorno = Format(Tab_Prod(Glob_Conta_Prod).Desconto_Perc, "0.00")
      �_N�mero = True
      �_Texto = False
    End If
  End If
   
  If Texto = "PRE�O UNIT�RIO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      '04/05/2004 - Daniel
      'Personaliza��o para a Embalavi
      If g_bln5CasasDecimais Then
        Retorno = Format(gsHandleNull(Tab_Prod(Glob_Conta_Prod).Valor_Unit & ""), "##,###,##0.00000")
        'Acendo Flag de Tratamento de 5 casas
        m_blnEmbalavi = True
      '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
      ElseIf g_bln3CasasDecimais Then
        Retorno = Format(gsHandleNull(Tab_Prod(Glob_Conta_Prod).Valor_Unit & ""), "##,###,##0.000")
        m_blnEmbalavi = True
      Else
        Retorno = Format(gsHandleNull(Tab_Prod(Glob_Conta_Prod).Valor_Unit & ""), "##,###,##0.00")
      End If
      
      �_Valor = True
      �_Texto = False
    End If
  End If
  If Texto = "PRE�O PRODUTO TOTAL" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Format(gsHandleNull(Tab_Prod(Glob_Conta_Prod).Valor_Total & ""), "##,###,##0.00")
      �_Valor = True
      �_Texto = False
    End If
  End If
  If Texto = "PERC ICM PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Aliq_ICM
    End If
  End If
  If Texto = "VALOR ICM PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Valor_ICM
    End If
  End If
  If Texto = "PERC IPI PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Aliq_IPI
    End If
  End If
  If Texto = "VALOR IPI PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).C�digo <> "" Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Valor_IPI
      �_Valor = True
      �_Texto = False
    End If
  End If
  If Texto = "PRE�O FINAL PRODUTO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Valor_Final <> 0 Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Valor_Final
      �_Valor = True
      �_Texto = False
    End If
  End If
  
  If Texto = "C�DIGO PESQUISA 1" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Pesq1 <> 0 Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Pesq1
      �_N�mero = True
    End If
  End If
  If Texto = "C�DIGO PESQUISA 2" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Pesq2 <> 0 Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Pesq2
      �_N�mero = True
    End If
  End If
  If Texto = "C�DIGO PESQUISA 3" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Pesq3 <> 0 Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Pesq3
      �_N�mero = True
    End If
  End If
  If Texto = "NOME PESQUISA 1" Then
    Retorno = Tab_Prod(Glob_Conta_Prod).Nome_Pesq1
    �_Texto = True
  End If
  If Texto = "NOME PESQUISA 2" Then
    Retorno = Tab_Prod(Glob_Conta_Prod).Nome_Pesq2
    �_Texto = True
  End If
  If Texto = "NOME PESQUISA 3" Then
    Retorno = Tab_Prod(Glob_Conta_Prod).Nome_Pesq3
    �_Texto = True
  End If
  
  If Texto = "DESCRI��O ADICIONAL" Then
    Retorno = Tab_Prod(Glob_Conta_Prod).Descr_Adicional
    �_Texto = True
  End If
  
  If Texto = "COR" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Cor <> 0 Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Cor
      �_N�mero = True
    End If
  End If
  
  If Texto = "NOME COR" Then
    Retorno = Tab_Prod(Glob_Conta_Prod).Nome_Cor
    �_Texto = True
  End If
  
  If Texto = "TAMANHO" Then
    Retorno = ""
    �_Texto = True
    If Tab_Prod(Glob_Conta_Prod).Tamanho <> 0 Then
      Retorno = Tab_Prod(Glob_Conta_Prod).Tamanho
      �_N�mero = True
    End If
  End If
  
  If Texto = "NOME TAMANHO" Then
    Retorno = Tab_Prod(Glob_Conta_Prod).Nome_Tamanho
    �_Texto = True
  End If
  
  '27/09/2004 - mpdea
  'Volumagem por Quantidade
  If Texto = "VOLUMAGEM POR QTDE" Then
    Retorno = Tab_Prod(Glob_Conta_Prod).VolumagemQtde
    �_Texto = True
  End If
    
  Rem Servicos
  If Texto = "C�DIGO SERVI�O" Then
    Retorno = Tab_Serv(Glob_Conta_Serv).C�digo
    �_Texto = True
    If Retorno = "0" Then Retorno = ""
  End If
  
  '29/04/2008 - mpdea
  'CFOP do servi�o
  If Texto = "CFOP DO SERVI�O" Then
    Retorno = ""
    �_Texto = True
    �_Valor = False
    If Tab_Serv(Glob_Conta_Serv).C�digo <> 0 Then
      Retorno = Tab_Serv(Glob_Conta_Serv).CFOP
    End If
  End If
  
  If Texto = "NOME SERVI�O" Then
    Retorno = Tab_Serv(Glob_Conta_Serv).Descri��o
    �_Texto = True
  End If
  If Texto = "QTDE SERVI�O" Then
    Retorno = ""
    �_Texto = True
    �_Valor = False
    If Tab_Serv(Glob_Conta_Serv).Qtde <> 0 Then
      Retorno = Tab_Serv(Glob_Conta_Serv).Qtde
    End If
  End If
  If Texto = "PRE�O UNIT�RIO SERVI�O" Then
    Retorno = ""
    �_Texto = True
    If Tab_Serv(Glob_Conta_Serv).Pre�o_Unit <> 0 Then
      Retorno = Tab_Serv(Glob_Conta_Serv).Pre�o_Unit
      �_Valor = True
      �_Texto = False
    End If
  End If
  If Texto = "PRE�O TOTAL SERVI�O" Then
    Retorno = ""
    �_Texto = True
    If Tab_Serv(Glob_Conta_Serv).Pre�o_Total <> 0 Then
      Retorno = Tab_Serv(Glob_Conta_Serv).Pre�o_Total
      �_Valor = True
      �_Texto = False
    End If
  End If
  If Texto = "SERVI�O CONCLU�DO" Then
    Retorno = Tab_Serv(Glob_Conta_Serv).Conclu�do
    �_Texto = True
  End If
  If Texto = "VALOR TOTAL SERVI�O" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = ""
      �_Texto = True
      If Glob_Total_Servi�o <> 0 Then
        Retorno = Glob_Total_Servi�o
        �_Valor = True
        �_Texto = False
      End If
    End If
  End If
  If Texto = "CST" Then
    '27/07/2005 - Daniel
    'CST (C�digo de Situa��o Tribut�ria)
    'Finalidade: Atender a realidade da empresa W.V. Hidroan�lise Ltda (J.R. Hidroqu�mica)
    Retorno = Tab_Serv(Glob_Conta_Serv).CST
    �_Texto = True
  End If
  
  If Texto = "BASE C�LCULO ICM" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Base_ICM
      �_Valor = True
    End If
  End If
  If Texto = "VALOR ICM" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Valor_ICM
      �_Valor = True
    End If
  End If
  If Texto = "BASE C�LCULO ICM SUBS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Base_ICM_Sub
      �_Valor = True
    End If
  End If
  If Texto = "VALOR ICM SUBS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Valor_ICM_Sub
      �_Valor = True
    End If
  End If
  If Texto = "VALOR TOTAL DOS PRODUTOS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Total_Produto
      �_Valor = True
    End If
  End If
  
  '19/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos com desconto no subtotal
  If Texto = "VALOR TOTAL DOS PRODUTOS COM DESCONTO NO SUBTOTAL" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblTotalProdutosDST
      �_Valor = True
    End If
  End If
  
  
  '26/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos menos total de descontos
  If Texto = "VALOR TOTAL DOS PRODUTOS MENOS TOTAL DE DESCONTOS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblTotalProdutosMenosDescontos
      �_Valor = True
    End If
  End If
  
  
  '06/09/2002 - mpdea
  'Inclu�do o campo para exibi��o do Desconto no SubTotal
  If Texto = "DESCONTO NO SUBTOTAL" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = g_dblDescontoSubTotal
      �_Valor = True
    End If
  End If
  
  
  If Texto = "VALOR TOTAL DE DESCONTOS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Total_Desconto
      �_Valor = True
    End If
  End If
  
  If Texto = "VALOR FRETE" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Frete
      �_Valor = True
    End If
  End If
  If Texto = "VALOR TOTAL IPI" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_IPI
      �_Valor = True
    End If
  End If
  If Texto = "VALOR TOTAL DA NOTA" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Total_Pagar
      �_Valor = True
    End If
  End If
  
  If Texto = "VALOR TOTAL SERVI�OS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Total_Servi�o
      �_Valor = True
    End If
  End If
  If Texto = "VALOR ISS" Then
    '08/03/2004 - mpdea
    'Tratamento para impress�o multinota
    If g_blnPrintMNF Then
      Retorno = String(13, "*")
      �_Texto = True
    Else
      Retorno = Glob_Total_ISS
      �_Valor = True
    End If
  End If
  
  If Texto = "FATURA" Then
    Retorno = Tab_Fat(Glob_Conta_Fat).N�mero
    �_Texto = True
  End If
  
  If Texto = "VALOR FATURA" Then
    If Tab_Fat(Glob_Conta_Fat).Valor = 0 Then
      Texto = "ESPACO_BRANCO"  '-- LINHA DA ALTERA��O
      �_Texto = True
      �_Valor = False
    Else
      Retorno = Tab_Fat(Glob_Conta_Fat).Valor
      �_Valor = True
      �_Texto = False
    End If
  End If
  If Texto = "DATA FATURA" Then
    If Left(Tab_Fat(Glob_Conta_Fat).Vencimento, 5) = "00:00" Then
      Texto = "ESPACO_BRANCO"   '-- LINHA DA ALTERA��O
    Else
      Retorno = Tab_Fat(Glob_Conta_Fat).Vencimento
    End If
    �_Texto = True
  End If
  
    
  
  If Texto = "ESPACO_BRANCO" Then
    Retorno = ""
    �_Texto = True
  End If
  
  
  
  
  If Texto = "DESCRI��O" Then
   Retorno = Glob_Descri��o
   �_Texto = True
  End If
  
  If Texto = "VALOR" Then
    Retorno = Glob_Valor
    �_Valor = True
    �_Texto = False
  End If
  
  '15/01/2004 - Daniel
  'Valor Recebido do [Contas a Receber]
  If Texto = "VALOR RECEBIDO" Then
    Retorno = g_dblValorRecebidoCR
    �_Valor = True
    �_Texto = False
  End If
  
  'Valor Total do Ticket
  If Texto = "VALOR TOTAL DO TICKET" Then
    Retorno = g_dblValorRecebidoCR
    �_Valor = True
    �_Texto = False
  End If
  '-----------------------------------
   
  If Texto = "DESCONTO" Then
    Retorno = Glob_Desconto
    �_Valor = True
    �_Texto = False
  End If
  
  If Texto = "ACR�SCIMO" Then
    Retorno = Glob_Acr�scimo
    �_Valor = True
    �_Texto = False
  End If
  
  If Texto = "VENCIMENTO" Then
    Retorno = Glob_Vencimento
    �_Texto = True
  End If
  
  If Texto = "MENSAGEM CLIENTE" Then
    Retorno = Glob_Mensagem_Cli
    �_Texto = True
  End If
  
  If Texto = "FATURA RECEBER" Then
    Retorno = Glob_Fatura
    �_Texto = True
  End If
  
  
  Rem Campos de Extenso
  
  
  If Texto = "EXTENSO1_60" Then
    Retorno = Extenso1_60
    �_Texto = True
  End If
  If Texto = "EXTENSO61_120" Then
    Retorno = Extenso61_120
    �_Texto = True
  End If
  If Texto = "EXTENSO121_180" Then
    Retorno = Extenso121_180
    �_Texto = True
  End If
   
  If Texto = "EXTENSO1_45" Then
    Retorno = Extenso1_45
    �_Texto = True
  End If
  If Texto = "EXTENSO46_90" Then
    Retorno = Extenso46_90
    �_Texto = True
  End If
  If Texto = "EXTENSO91_135" Then
    Retorno = Extenso91_135
    �_Texto = True
  End If
  If Texto = "EXTENSO136_180" Then
    Retorno = Extenso136_180
    �_Texto = True
  End If
  
  If Texto = "EXTENSO1_30" Then
    Retorno = Extenso1_30
    �_Texto = True
  End If
  If Texto = "EXTENSO31_60" Then
    Retorno = Extenso31_60
    �_Texto = True
  End If
  If Texto = "EXTENSO61_90" Then
    Retorno = Extenso61_90
    �_Texto = True
  End If
  If Texto = "EXTENSO91_120" Then
    Retorno = Extenso91_120
    �_Texto = True
  End If
  If Texto = "EXTENSO121_150" Then
    Retorno = Extenso121_150
    �_Texto = True
  End If
  If Texto = "EXTENSO151_180" Then
    Retorno = Extenso151_180
    �_Texto = True
  End If
  
  '19/12/2007 - Anderson
  'Implementa��o do NSU
  If Texto = "NSU" Then
    Retorno = NSU
    �_Texto = True
  End If
  
  If Texto = "NSU (DATA GERA��O)" Then
    Retorno = NSU_Data
    �_Texto = True
  End If
  
  If Texto = "NSU (HORA GERA��O)" Then
    Retorno = NSU_Hora
    �_Texto = True
  End If
  
  If �_Texto = True Then
    If Tamanho > 0 Then
      Retorno = Retorno + Brancos
      Retorno = Left(Retorno, Tamanho)
    End If
  End If
  
  If �_N�mero = True Then
    Retorno = "                                    " + Retorno
    Retorno = Right(Retorno, Tamanho)
  End If
  
  If �_Zeros = True Then
    Retorno = "0000000000000000000000" + Retorno
    Retorno = Right(Retorno, Tamanho)
  End If
  
  If �_Valor = True Then
    '04/05/2004 - Daniel
    'Personaliza��o Embalavi
    If m_blnEmbalavi Then
    
    '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
      If g_bln3CasasDecimais Then
        Aux = Format$(CDbl(Retorno), "###,###,###,##0.000")
        Retorno = "                      " + Aux
        Retorno = Right(Retorno, 12)
      Else
        Aux = Format$(CDbl(Retorno), "###,###,###,##0.00000")
        Retorno = "                      " + Aux
        Retorno = Right(Retorno, 12)
      End If
      m_blnEmbalavi = False
    
    Else
      'N�o Embalavi
      Aux = Format$(CDbl(Retorno), "###,###,###,##0.00")
      Retorno = "                      " + Aux
      Retorno = Right(Retorno, Tamanho)
    End If
    
  End If

  Pega_Campo = Retorno

End Function

Sub Imprime_Cheque(ByVal Favorecido As String, ByVal Banco As Integer, ByVal Dia As String, ByVal Valor As Double)
  Dim rsParametros As Recordset
  Dim Cidade As String
  Dim Str_Impre As String
  Dim Str_In�cio As String
  Dim Str_Favorecido As String
  Dim Str_Localidade As String
  Dim Str_Banco As String
  Dim Str_Valor As String
  Dim Str_Data As String
  Dim Num_cod As Integer
  Dim Num_Cod2 As Integer
  Dim Str_Fim As String
  Dim Resposta As Integer
  Dim Str_Valor2 As String
  
  On Error GoTo ErrHandle
  
  Set rsParametros = db.OpenRecordset("Par�metros Filial", , dbReadOnly)
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then Exit Sub
  
  If Valor = 0 Then
    Exit Sub
  ElseIf Banco = 0 Then
    Exit Sub
  End If
  If Favorecido = "()" Then
    Favorecido = rsParametros("Cheque Favorecido")
  End If
  Cidade = rsParametros("Cheque Cidade")
  
  Str_Valor2 = Trim(str(Valor))
  If rsParametros("Imprimir Centavos") = True Then
    Str_In�cio = Trim(Format(Valor, "##########0.00"))
    Str_Valor2 = Left(Str_In�cio, (Len(Str_In�cio) - 3))
    Str_Valor2 = Str_Valor2 + "," + Right(Str_In�cio, 2)
  End If
     
  'String de In�cio da Impressora
  Num_cod = 0
  Str_In�cio = ""
  If rsParametros("In�cio Cheque 1") <> "" Then
    Num_cod = 1
    If rsParametros("In�cio Cheque 2") <> "" Then
      Num_cod = 2
    End If
  End If
  If Num_cod = 1 Then
    Str_In�cio = Chr$(Val(rsParametros("In�cio Cheque 1"))) + Chr$(27)
  End If
  If Num_cod = 2 Then
    Str_In�cio = Chr$(Val(rsParametros("In�cio Cheque 1")))
    Str_In�cio = Str_In�cio + Chr$(Val(rsParametros("In�cio Cheque 2"))) + Chr$(13)
  End If
  
  'String do Favorecido
  Num_cod = 0
  Str_Favorecido = ""
  If rsParametros("C�digo Favorecido 1") <> "" Then
    Num_cod = 1
    If rsParametros("C�digo Favorecido 2") <> "" Then
      Num_cod = 2
    End If
  End If
  If Num_cod = 1 Then
    Str_Favorecido = Chr$(Val(rsParametros("C�digo Favorecido 1"))) + Favorecido
  End If
  If Num_cod = 2 Then
    Str_Favorecido = Chr$(Val(rsParametros("C�digo Favorecido 1")))
    Str_Favorecido = Str_Favorecido + Chr$(Val(rsParametros("C�digo Favorecido 2"))) + Favorecido
  End If
  If rsParametros("C�digo Favorecido 3") <> "" Then
    Str_Favorecido = Str_Favorecido + Chr$(Val(rsParametros("C�digo Favorecido 3")))
  End If
  
  'String da Localidade
  Num_cod = 0
  Str_Localidade = ""
  If rsParametros("C�digo Cidade 1") <> "" Then
    Num_cod = 1
    If rsParametros("C�digo Cidade 2") <> "" Then
      Num_cod = 2
    End If
  End If
  If Num_cod = 1 Then
    Str_Localidade = Chr$(Val(rsParametros("C�digo Cidade 1"))) + Cidade
  End If
  If Num_cod = 2 Then
    Str_Localidade = Chr$(Val(rsParametros("C�digo Cidade 1")))
    Str_Localidade = Str_Localidade + Chr$(Val(rsParametros("C�digo Cidade 2"))) + Cidade
  End If
  If rsParametros("C�digo Cidade 3") <> "" Then
    Str_Localidade = Str_Localidade + Chr$(Val(rsParametros("C�digo Cidade 3")))
  End If
  
  'String do Banco
  Num_cod = 0
  Str_Banco = ""
  If rsParametros("C�digo Banco 1") <> "" Then
    Num_cod = 1
    If rsParametros("C�digo Banco 2") <> "" Then
       Num_cod = 2
    End If
  End If
  If Num_cod = 1 Then
     Str_Banco = Chr$(Val(rsParametros("C�digo Banco 1"))) + Trim(str(Banco))
  End If
  If Num_cod = 2 Then
     Str_Banco = Chr$(Val(rsParametros("C�digo Banco 1")))
     Str_Banco = Str_Banco + Chr$(Val(rsParametros("C�digo Banco 2"))) + Trim(str(Banco))
  End If
  If rsParametros("C�digo Banco 3") <> "" Then
    Str_Banco = Str_Banco + Chr$(Val(rsParametros("C�digo Banco 3")))
  End If
  
  'String do Valor
  Num_cod = 0
  Str_Valor = ""
  If rsParametros("C�digo Valor 1") <> "" Then
    Num_cod = 1
    If rsParametros("C�digo Valor 2") <> "" Then
       Num_cod = 2
    End If
  End If
  If Num_cod = 1 Then
     Str_Valor = Chr$(Val(rsParametros("C�digo Valor 1"))) + Str_Valor2
  End If
  If Num_cod = 2 Then
     Str_Valor = Chr$(Val(rsParametros("C�digo Valor 1")))
     Str_Valor = Str_Valor + Chr$(Val(rsParametros("C�digo Valor 2"))) + Str_Valor2
  End If
  If rsParametros("C�digo Valor 3") <> "" Then
    Str_Valor = Str_Valor + Chr$(Val(rsParametros("C�digo Valor 3")))
  End If
  
  'String da Data
  Dia = Format(CDate(Dia), "dd/mm/yy")
  Num_cod = 0
  Str_Data = ""
  If rsParametros("C�digo Data 1") <> "" Then
    Num_cod = 1
    If rsParametros("C�digo Data 2") <> "" Then
      Num_cod = 2
    End If
  End If
  If Num_cod = 1 Then
    Str_Data = Chr$(Val(rsParametros("C�digo Data 1"))) + Dia
  End If
  If Num_cod = 2 Then
    Str_Data = Chr$(Val(rsParametros("C�digo Data 1")))
    Str_Data = Str_Data + Chr$(Val(rsParametros("C�digo Data 2"))) + Dia
  End If
  If rsParametros("C�digo Data 3") <> "" Then
    Str_Data = Str_Data + Chr$(Val(rsParametros("C�digo Data 3")))
  End If
  
  'String de Final da Impressora
  Num_cod = 0
  Str_Fim = ""
  If rsParametros("Imprime Cheque 1") <> "" Then
    Num_cod = 1
    If rsParametros("Imprime Cheque 2") <> "" Then
      Num_cod = 2
    End If
  End If
  If Num_cod = 1 Then
    Str_In�cio = Chr$(Val(rsParametros("Imprime Cheque 1"))) + Chr$(13)
  End If
  If Num_cod = 2 Then
    Str_Fim = Chr$(Val(rsParametros("Imprime Cheque 1")))
    Str_Fim = Str_Fim + Chr$(Val(rsParametros("Imprime Cheque 2"))) + Chr$(13)
  End If
  
  Call SetPrinterName("CHEQUE")
  
  Str_Impre = Str_In�cio + Str_Favorecido + Str_Localidade + Str_Banco + Str_Valor + Str_Data + Str_Fim
  Str_Impre = Chr$(Len(Str_Impre) Mod 256) + Chr$(Len(Str_Impre) \ 256) + Str_Impre
  Printer.Print ""
  If Not IsWindowsNT() Then
    Resposta = Escape(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
  Else
    Resposta = Escape32(Printer.hdc, PASSTHROUGH, 0, Str_Impre$, 0&)
  End If
  Printer.EndDoc
  
  Call SetPrinterName("REL")
  
  Exit Sub
  
ErrHandle:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao imprimir Cheque."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gsMsg = gsMsg & vbCrLf & "A impressora de Cheques est� corretamente definida?"
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Call SetPrinterName("REL")
  Exit Sub

End Sub

Public Sub Limpa_Vari�veis_Nota()

  Dim intContador As Integer '06/05/2007 - Anderson Contador Auxiliar
  
  Glob_Sequ�ncia = 0
  Glob_Nota_Impressa = 0
  Glob_Nome_Opera��o = ""
  Glob_C�digo_Fiscal = ""
  Glob_C�digo_Cli = 0
  Glob_Nome = ""
  Glob_Fantasia = ""
  Glob_CGC = ""
  Glob_Endere�o = ""
  Glob_NumeroEndereco = "" '23/10/2009 - mpdea
  Glob_Complemento = ""
  Glob_Bairro = ""
  Glob_CEP = ""
  Glob_Cidade = ""
  Glob_Fone1 = ""
  Glob_Fone2 = ""
  Glob_Estado = ""
  Glob_Inscri��o = ""
  Glob_Base_ICM = 0
  Glob_Valor_ICM = 0
  Glob_Base_ICM_Sub = 0
  Glob_Valor_ICM_Sub = 0
  Glob_Total_Produto = 0
  Glob_Frete = 0
  Glob_IPI = 0
  Glob_Total_Desconto = 0
  Glob_Total_Pagar = 0
  Glob_Nome_Vendedor = 0
  Glob_Apelido = 0
  Glob_Data_Emiss�o = ""
  Glob_Fatura = ""
  Glob_Descri��o = ""
  Glob_Vencimento = ""
  Glob_Valor = ""
  Glob_Desconto = ""
  Glob_Acr�scimo = ""
  Glob_Mensagem_Cli = ""
  
  '31/05/2007 - Anderson
  'Implementa��o dos campos Prometido para e or�amento aprovado por
  Glob_Aprovado = ""
  Glob_Prometido = ""
  
  '--------------------------------------------------
  '06/05/2007 - Anderson
  'Implementa��o da impress�o para CFOP's de servi�os
  Glob_C�digo_Fiscal_Completo = ""
  
  For intContador = 0 To 4
    Glob_C�digo_Fiscal_Item(intContador) = ""
  Next
  '--------------------------------------------------
    
  '15/08/2002 - mpdea
  'Inclu�do o campo de informa��es sobre o or�amento (n�mero do or�amento e terminal)
  gstrInfoNrOrcamento = ""
  
  '06/09/2002 - mpdea
  'Inclu�do o campo para exibi��o do Desconto no SubTotal
  g_dblDescontoSubTotal = 0
  
  '07/08/2003 - mpdea
  'Inclu�do C�digo e Nome do T�cnico
  g_intCodigoTecnico = 0
  g_strNomeTecnico = ""

  '19/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos com desconto no subtotal
  g_dblTotalProdutosDST = 0
  
  '26/08/2003 - mpdea
  'Inclu�do campo para totalizador de produtos menos total de descontos
  g_dblTotalProdutosMenosDescontos = 0
  
  
  '01/02/2006 - mpdea
  'Mensagens para Nota Fiscal
  Set m_clcMensagens = Nothing
  
  
  '----------------------------------------------
  '08/01/2004 - Daniel
  'Inclu�do vars para o campo Valor Recebido e
  'Troco da tabela de Sa�das
  g_dblValorRecebido = 0
  g_dblTroco = 0
  '----------------------------------------------
  
  '----------------------------------------------
  '09/01/2004 - Daniel
  'Inclu�do var para soma da Qtde de Itens
  g_sngQtdeItens = 0
  '----------------------------------------------
  
  '----------------------------------------------
  '30/01/2004 - Daniel
  'Inclu�do vars para tratamento de impostos
  g_dblPercentualCSLL = 0
  g_dblPercentualCOFINS = 0
  g_dblPercentualPIS = 0
  g_dblPercentualIRRF = 0
  
  g_dblTotalCSLL = 0
  g_dblTotalCOFINS = 0
  g_dblTotalPIS = 0
  g_dblTotalIRRF = 0
  '----------------------------------------------
  
  '27/04/2005 - Daniel
  'Tratamento para a global do Seguro
  g_dblSeguro = 0
  
  '----------------------------------------------
  '13/04/2004 - Daniel
  'Vars adicionadas devido solicita��o do cliente
  'STC de Caxias do Sul
  g_lngNumAutorizacao = 0
  g_intMesX = 0
  g_strEnderecoCob = ""
  g_strComplementoCob = ""
  g_strBairroCob = ""
  g_strCidadeCob = ""
  g_strEstadoCob = ""
  g_strCepCob = ""
  '----------------------------------------------
  
  '06/05/2004 - Daniel
  'Adi��o do campo ObsIsentoIPI da tabela Cli_For
  g_strObsIsentoIPI = ""
  '17/05/2004 - Daniel
  'Adi��o do campo ObsDiferimento da tabela Diferimento
  g_strObsDiferimento = ""
  '----------------------------------------------
    
  Limpa_Faturas
  Limpa_Produtos
  Limpa_Servi�os
  
End Sub

'16/10/2007 - Anderson
'Fun��o utilizada para criar CFOP
Private Function VerificaCFOP(ByVal valorCFOP As String) As Boolean
  Dim intContador As Integer
  
  'Se n�o houver CFOP, retorna false
  If Len(Trim(valorCFOP)) = 0 Then
    VerificaCFOP = False
    Exit Function
  End If
  
  'Verifica se o CFOP existe
  For intContador = 0 To 4
    If Glob_C�digo_Fiscal_Item(intContador) = valorCFOP Then
      VerificaCFOP = False
      Exit Function
    End If
  Next
  
  'Coloca o CFOP
  For intContador = 0 To 4
    If Len(Glob_C�digo_Fiscal_Item(intContador)) = 0 Then
      Glob_C�digo_Fiscal_Item(intContador) = valorCFOP
      VerificaCFOP = True
      Exit Function
    End If
  Next
  
End Function

'18/12/2007 - Anderson
'Fun��o para implementa��o do NSU
Public Sub GerarNSU(ByRef Tabela As Recordset, ByVal NomeTabela As String)

  Dim strSQL As String
  Dim rsParametros As Recordset
  Dim dblNSU As Double
  Dim strMotivo As String
  
  If NomeTabela = "Entradas" Then
    strMotivo = "Entrada"
  Else
    strMotivo = "Sa�da"
  End If
  
  
  strSQL = "SELECT NSU FROM [Par�metros Filial] WHERE Filial=" & gnCodFilial
  
  Set rsParametros = db.OpenRecordset(strSQL, dbOpenSnapshot)
 
  With rsParametros
  
    If .RecordCount > 0 Then
      If .Fields("NSU") = 9999999999# Then
        dblNSU = 1
        strMotivo = "Reinicializa��o do NSU"
      Else
        dblNSU = .Fields("NSU") + 1
      End If
      db.Execute "UPDATE [Par�metros Filial] SET NSU=" & dblNSU & " WHERE Filial=" & gnCodFilial
      strSQL = "UPDATE " & NomeTabela & " SET NSU_Data = #"
      strSQL = strSQL & Format(Date, "mm/dd/yyyy") & "#, "
      strSQL = strSQL & " NSU_Hora =#"
      strSQL = strSQL & Format(Now(), "hh:nn") & "# "
      strSQL = strSQL & "WHERE Filial = " & Tabela.Fields("Filial").Value
      strSQL = strSQL & " AND Sequ�ncia = " & Tabela.Fields("Sequ�ncia").Value
      db.Execute strSQL, dbFailOnError
      
      strSQL = "INSERT INTO NSU (Filial, NSU, Movimento, Motivo, Sequencia, NotaFiscal,Data_Hora,Total) "
      strSQL = strSQL & "VALUES (" & gnCodFilial & ",'" & Format(dblNSU, "0000000000") & "','" & NomeTabela & "','" & strMotivo & "',"
      strSQL = strSQL & Tabela("Sequ�ncia") & ","
      strSQL = strSQL & Tabela("Nota Impressa") & ","
      strSQL = strSQL & "#" & Format(Now, "MM/DD/YYYY HH:NN:SS") & "#,"
      strSQL = strSQL & Replace(Tabela("Total"), ",", ".") & ")"
      db.Execute strSQL, dbFailOnError
    End If
  
    .Close
  
  End With
  
  Set rsParametros = Nothing
  
  With Tabela
    .LockEdits = True
    .Edit
    .Fields("NSU").Value = dblNSU
    .Update
    .LockEdits = False
  End With
  
End Sub

