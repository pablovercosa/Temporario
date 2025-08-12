Attribute VB_Name = "modPrintBoleto"
Option Explicit

Public Type ParcelasFatura
  pfDataVencimento As String
  pfValor As String
End Type

Public pfParcelasFatura() As ParcelasFatura

Public Function Imprime_Boleto(ByVal Tipo As String, ByVal Filial As Integer, ByVal Vencimento As Date, ByVal Contador As Long, ByVal Nome_Boleto As String) As Integer
  Dim rsReceber As Recordset
  Dim rsClientes As Recordset
  Dim Aux As Variant
  Dim Nome_Arq As String
  Dim Texto As String
  Dim Final As Integer
  Dim Str_Impre As String
  Dim Num_cod As Integer
  Dim Resposta As Long
  Dim Final_Linha As Integer
  Dim Linhas As Integer
  Dim Especial2 As Integer
  Dim Extenso_Tot As String
  Dim nFileNum As Integer
  Dim nCtLin As Integer
  Dim nComprPag As Integer
  Dim sParte As String
  
  On Error GoTo ErrHandler
  
  SetPrinterName ("BOLETO")
  
  gsInitPrinter = ""
  Call ResetPrinter
  
  nFileNum = FreeFile
  Open Nome_Boleto For Input As #nFileNum
  
  Input #nFileNum, Texto
  If Left(Texto, 24) <> "*** Configura��es Boleto" Then
    gsTitle = LoadResString(201)
    gsMsg = "Layout do cabe�alho do arquivo de configura��o """ & Nome_Boleto & """ diferente do esperado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Boleto = 3
    Close #nFileNum
    Exit Function
  End If
  
  nComprPag = 0
  sParte = UCase(Mid(Texto, 75, 3))
  If Len(sParte) > 0 Then
    If sParte <> "N�O" And sParte <> "LIN" Then
      If Not IsNumeric(sParte) Then
        DisplayMsg "Valor para par�metro de comprimento da p�gina pode ser: N�O, LIN ou <99> (inteiro dois digitos)."
        Imprime_Boleto = 3
        Close #nFileNum
        Exit Function
      End If
      If Val(sParte) <= 0 Or Val(sParte) > 20 Then
        DisplayMsg "Comprimento da p�gina em polegadas inv�lido."
        Imprime_Boleto = 3
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
        Open Nome_Boleto For Input As #nFileNum
        Input #nFileNum, Texto
      End If
    End If
  End If

  
  If Mid(Texto, 40, 3) = "SIM" Then
    If SetCompressPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel usar compress�o na impressora solicitada pelo arquivo de configura��o: """ & Nome_Boleto & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Imprime_Boleto = 5
      SetPrinterName ("REL")
      Close #nFileNum
      Exit Function
    End If
  End If
  
  If Mid(Texto, 55, 3) = "SIM" Then
    If SetOitavoPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel ajustar a impressora para 1/8 solicitada pelo arquivo de configura��o: """ & Nome_Boleto & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Imprime_Boleto = 6
      SetPrinterName ("REL")
      Close #nFileNum
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
      Imprime_Boleto = 4
      SetPrinterName ("REL")
      Close #nFileNum
      Exit Function
    End If
  End If
  
  Call SetPrinterCommand(gsInitPrinter)
  
  Set rsReceber = db.OpenRecordset("Contas a Receber", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  
  rsReceber.Index = "Vencimento"
  rsReceber.Seek "=", Tipo, Filial, Vencimento, Contador
  If rsReceber.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Conta do Contas a Receber n�o foi localizada: Tipo=" & Tipo & ", Filial=" & Filial & ", Vencimento=" & Vencimento & ", Contador=" & Contador
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Boleto = 1
    Close #nFileNum
    Exit Function
  End If
  
  rsClientes.Index = "C�digo"
  rsClientes.Seek "=", rsReceber("Cliente")
  If rsClientes.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Cliente referido pela Conta do Contas a Receber n�o foi localizado: Cliente=" & rsReceber("Cliente") & ", Vencimento=" & Vencimento
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Boleto = 2
    Close #nFileNum
    Exit Function
  End If
  
    
  Rem Inicializa vari�veis nota
  Limpa_Vari�veis_Boleto
  
  
  Glob_Nota_Impressa = rsReceber("Nota")
  Glob_Nome = rsClientes("Nome") & ""
  Glob_Fantasia = rsClientes("Fantasia") & ""
  Glob_CGC = rsClientes("CGC") & ""
  Glob_Inscri��o = rsClientes("Inscri��o") & ""
  Glob_Data_Emiss�o = rsReceber("Data Emiss�o") & ""
  
  If IsNull(rsClientes("Endere�o Cob")) Or rsClientes("Endere�o Cob") = "" Then
    Glob_Endere�o = rsClientes("Endere�o") & ""
    Glob_NumeroEndereco = rsClientes.Fields("Endere�o N�mero").Value & "" '23/10/2009 - mpdea
    
    Glob_Complemento = rsClientes("Complemento") & ""
    Glob_Bairro = rsClientes("Bairro") & ""
    Glob_CEP = rsClientes("Cep") & ""
    Glob_Cidade = rsClientes("Cidade") & ""
    Glob_Estado = rsClientes("Estado") & ""
  Else
    Glob_Endere�o = rsClientes("Endere�o Cob") & ""
    Glob_Complemento = rsClientes("Complemento Cob") & ""
    Glob_Bairro = rsClientes("Bairro Cob") & ""
    Glob_CEP = rsClientes("Cep Cob") & ""
    Glob_Cidade = rsClientes("Cidade Cob") & ""
    Glob_Estado = rsClientes("Estado Cob") & ""
  End If
  
  Glob_Data_Sa�da = str(Date)
  Glob_Fatura = rsReceber("Fatura") & ""
  Glob_Descri��o = rsReceber("Descri��o") & ""
  Glob_Vencimento = rsReceber("Vencimento") & ""
  Glob_Valor = rsReceber("Valor")
  Glob_Desconto = rsReceber("Desconto")
  Glob_Acr�scimo = rsReceber("Acr�scimo")
  Glob_Mensagem_Cli = rsClientes("Mensagem Boleto") & ""
  gsObsDoc(0) = gsObsDoc(0) & ""
  gsObsDoc(1) = gsObsDoc(1) & ""
  gsObsDoc(2) = gsObsDoc(2) & ""
  Glob_C�digo_Cli = rsClientes("C�digo")
  Glob_Sequ�ncia = rsReceber("Sequ�ncia")
  '15/01/2004 - Daniel
  'Populando g_dblValorRecebidoCR do campo
  'Valor Recebido do [Contas a Receber]
  g_dblValorRecebidoCR = rsReceber("Valor Recebido")
  
  Extenso_Tot = Extenso(rsReceber("Valor"))
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
  
   
  Final = False
  Do
    DoEvents
    If gbToCancel = True Then
      Exit Do
    End If
    Input #nFileNum, Texto
    If Texto = "*** Fim de arquivo ***" Then Final = True
    If Final = False Then
      Texto = Apaga_Aspas(Texto)
      Final_Linha = False
      If Len(Texto) < 3 Then
        Imprime_Boleto = 8
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
      
  Imprime_Boleto = 0
  
  Close #nFileNum
  
  Printer.Print
  Printer.EndDoc
  SetPrinterName ("REL")
  
  Exit Function
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao imprimir Nota usando o arquivo de configura��o: """ & Nome_Boleto & """."
  If Err.Number = 53 Then
    gsMsg = gsMsg & vbCrLf & "Arquivo de configura��o n�o encontrado."
  Else
    gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  End If
  On Error Resume Next
  Close #nFileNum
  Exit Function
  
  SetPrinterName ("REL")
  Imprime_Boleto = 4
  On Error Resume Next
  Close #nFileNum
  Exit Function
  
End Function

Public Function Imprime_Carn�(ByVal Tipo As String, ByVal Filial As Integer, ByVal Vencimento As Date, ByVal Contador As Long, ByVal Nome_Carne As String) As Integer
  Dim rsReceber As Recordset
  Dim rsClientes As Recordset
  Dim Aux As Variant
  Dim Nome_Arq As String
  Dim Texto As String
  Dim Final As Integer
  Dim Str_Impre As String
  Dim Num_cod As Integer
  Dim Resposta As Long
  Dim Final_Linha As Integer
  Dim Linhas As Integer
  Dim Especial2 As Integer
  Dim Extenso_Tot As String
  Dim nFileNum As Integer
  Dim nCtLin As Integer
  Dim nComprPag As Integer
  Dim sParte As String
  
  On Error GoTo ErrHandler
  
  SetPrinterName ("CARN�")
  'SetPrinterNameCARNE_TESTE ("CARN�")
  
  gsInitPrinter = ""
  Call ResetPrinter
  
  nFileNum = FreeFile
  Open Nome_Carne For Input As #nFileNum
  
  Input #nFileNum, Texto
  If Left(Texto, 24) <> "*** Configura��es Carn� " Then
    gsTitle = LoadResString(201)
    gsMsg = "Layout do cabe�alho do arquivo de configura��o """ & Nome_Carne & """ diferente do esperado."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Carn� = 3
    Close #nFileNum
    Exit Function
  End If
  
  nComprPag = 0
  sParte = UCase(Mid(Texto, 75, 3))
  If Len(sParte) > 0 Then
    If sParte <> "N�O" And sParte <> "LIN" Then
      If Not IsNumeric(sParte) Then
        DisplayMsg "Valor para par�metro de comprimento da p�gina pode ser: N�O, LIN ou <99> (inteiro dois digitos)."
        Imprime_Carn� = 3
        Close #nFileNum
        Exit Function
      End If
      If Val(sParte) <= 0 Or Val(sParte) > 20 Then
        DisplayMsg "Comprimento da p�gina em polegadas inv�lido."
        Imprime_Carn� = 3
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
        Open Nome_Carne For Input As #nFileNum
        Input #nFileNum, Texto
      End If
    End If
  End If

  
  If Mid(Texto, 40, 3) = "SIM" Then
    If SetCompressPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel usar compress�o na impressora solicitada pelo arquivo de Nome_Carne: """ & Nome_Carne & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Imprime_Carn� = 5
      SetPrinterName ("REL")
      Close #nFileNum
      Exit Function
    End If
  End If
  
  If Mid(Texto, 55, 3) = "SIM" Then
    If SetOitavoPrinter(Filial) <> 0 Then
      gsTitle = LoadResString(201)
      gsMsg = "N�o foi poss�vel ajustar a impressora para 1/8 solicitada pelo arquivo de Nome_Carne: """ & Nome_Carne & """."
      gnStyle = vbOKOnly + vbExclamation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Imprime_Carn� = 6
      SetPrinterName ("REL")
      Close #nFileNum
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
      Imprime_Carn� = 6
      SetPrinterName ("REL")
      Close #nFileNum
      Exit Function
    End If
  End If
  
  Call SetPrinterCommand(gsInitPrinter)
  
  Set rsReceber = db.OpenRecordset("Contas a Receber", , dbReadOnly)
  Set rsClientes = db.OpenRecordset("Cli_For", , dbReadOnly)
  
  rsReceber.Index = "Vencimento"
  rsReceber.Seek "=", Tipo, Filial, Vencimento, Contador
  If rsReceber.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Conta do Contas a Receber n�o foi localizada: Tipo=" & Tipo & ", Filial=" & Filial & ", Vencimento=" & Vencimento & ", Contador=" & Contador
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Carn� = 1
    Close #nFileNum
    Exit Function
  End If
  
  rsClientes.Index = "C�digo"
  rsClientes.Seek "=", rsReceber("Cliente")
  If rsClientes.NoMatch Then
    gsTitle = LoadResString(201)
    gsMsg = "Cliente referido pela Conta do Contas a Receber n�o foi localizado: Cliente=" & rsReceber("Cliente") & ", Vencimento=" & Vencimento
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Imprime_Carn� = 2
    Close #nFileNum
    Exit Function
  End If
  
    
  Rem Inicializa vari�veis nota
  Limpa_Vari�veis_Boleto
  
  
  Glob_Nota_Impressa = rsReceber("Nota")
  Glob_Nome = rsClientes("Nome") & ""
  Glob_Fantasia = rsClientes("Fantasia") & ""
  Glob_CGC = rsClientes("CGC") & ""
  Glob_Inscri��o = rsClientes("Inscri��o") & ""
  Glob_Data_Emiss�o = rsReceber("Data Emiss�o") & ""
  
  If IsNull(rsClientes("Endere�o Cob")) Or rsClientes("Endere�o Cob") = "" Then
    Glob_Endere�o = rsClientes("Endere�o") & ""
    Glob_NumeroEndereco = rsClientes.Fields("Endere�o N�mero").Value & "" '23/10/2009 - mpdea
    
    Glob_Complemento = rsClientes("Complemento") & ""
    Glob_Bairro = rsClientes("Bairro") & ""
    Glob_CEP = rsClientes("Cep") & ""
    Glob_Cidade = rsClientes("Cidade") & ""
    Glob_Estado = rsClientes("Estado") & ""
  Else
    Glob_Endere�o = rsClientes("Endere�o Cob") & ""
    Glob_Complemento = rsClientes("Complemento Cob") & ""
    Glob_Bairro = rsClientes("Bairro Cob") & ""
    Glob_CEP = rsClientes("Cep Cob") & ""
    Glob_Cidade = rsClientes("Cidade Cob") & ""
    Glob_Estado = rsClientes("Estado Cob") & ""
  End If
  
  Glob_Data_Sa�da = str(Date)
  Glob_Fatura = rsReceber("Fatura") & ""
  Glob_Descri��o = rsReceber("Descri��o") & ""
  Glob_Vencimento = rsReceber("Vencimento") & ""
  Glob_Valor = rsReceber("Valor")
  Glob_Desconto = rsReceber("Desconto")
  Glob_Acr�scimo = rsReceber("Acr�scimo")
  Glob_Mensagem_Cli = rsClientes("Mensagem Boleto") & ""
  gsObsDoc(0) = gsObsDoc(0) & ""
  gsObsDoc(1) = gsObsDoc(1) & ""
  gsObsDoc(2) = gsObsDoc(2) & ""
  Glob_C�digo_Cli = rsClientes("C�digo")
  Glob_Sequ�ncia = rsReceber("Sequ�ncia")
  '15/01/2004 - Daniel
  'Populando g_dblValorRecebidoCR do campo
  'Valor Recebido do [Contas a Receber]
  g_dblValorRecebidoCR = rsReceber("Valor Recebido")
  
  
  Extenso_Tot = Extenso(rsReceber("Valor"))
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
  
  
  
   
  Final = False
  Do
    DoEvents
    If gbToCancel = True Then
      Exit Do
    End If
    Input #nFileNum, Texto
    If Texto = "*** Fim de arquivo ***" Then Final = True
    If Final = False Then
      Texto = Apaga_Aspas(Texto)
      Final_Linha = False
      If Len(Texto) < 3 Then
        Imprime_Carn� = 8
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
      
  Imprime_Carn� = 0
  
  Close #nFileNum
  
  Printer.Print
  Printer.EndDoc
  SetPrinterName ("REL")
  Exit Function
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao imprimir Carn� usando o arquivo de configura��o: """ & Nome_Carne & """."
  If Err.Number = 53 Then
    gsMsg = gsMsg & vbCrLf & "Arquivo de configura��o n�o encontrado."
  Else
    gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  End If
  On Error Resume Next
  Imprime_Carn� = 4
  Close #nFileNum
  Exit Function
  
End Function

Sub Limpa_Vari�veis_Boleto()
  Glob_Nota_Impressa = 0
  Glob_Nome = ""
  Glob_Fantasia = ""
  Glob_CGC = ""
  Glob_Inscri��o = ""
  Glob_Data_Emiss�o = ""
  Glob_Endere�o = ""
  Glob_NumeroEndereco = "" '23/10/2009 - mpdea
  Glob_Complemento = ""
  Glob_Bairro = ""
  Glob_CEP = ""
  Glob_Cidade = ""
  Glob_Estado = ""
  Glob_Data_Sa�da = Empty
  Glob_Fatura = ""
  Glob_Descri��o = ""
  Glob_Vencimento = ""
  Glob_Valor = 0
  Glob_Desconto = 0
  Glob_Acr�scimo = 0
  Glob_Mensagem_Cli = ""
  Glob_C�digo_Cli = 0
  Glob_Sequ�ncia = 0
  '15/01/2004 - Daniel
  g_dblValorRecebidoCR = 0
  
End Sub


