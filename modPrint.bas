Attribute VB_Name = "modPrint"
Option Explicit

Private Declare Function WTSGetActiveConsoleSessionId Lib "Kernel32.dll" () As Long
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetCurrentProcessId Lib "Kernel32" () As Long

Private Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessId As Long
    pProcessName As Long
    pUserSid As Long
End Type

Function TerminalServerSessionId() As String
    Dim lRetVal As Long, lCount As Long, lThisProcess As Long, lThisProcessId  As Long
    Dim lpBuffer As Long, lp As Long, udtProcessInfo As WTS_PROCESS_INFO
    Const WTS_CURRENT_SERVER_HANDLE = 0&
    
    On Error GoTo ErrNotTerminalServer
    'Set Default Value
    TerminalServerSessionId = "0"
    lThisProcessId = GetCurrentProcessId
    lRetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, lCount)
    If lRetVal Then
        'Successful
        lp = lpBuffer
        For lThisProcess = 1 To lCount
            CopyMemory udtProcessInfo, ByVal lp, LenB(udtProcessInfo)
            If lThisProcessId = udtProcessInfo.ProcessId Then
                TerminalServerSessionId = CStr(udtProcessInfo.SessionID)
                Exit For
            End If
            lp = lp + LenB(udtProcessInfo)
        Next
        'Free memory buffer
        WTSFreeMemory lpBuffer
    End If
    
    Exit Function
    
ErrNotTerminalServer:
    'The machine is not a Terminal Server
    On Error GoTo 0
End Function

'Public Sub SetPrinterModeloPwd1(ByRef objReport As CrystalReport)
'  ' Modelo 1
'  objReport.LogonInfo(0) = "Provider=Microsoft.Jet.OLEDB.4.0;dsn=;uid=;pwd=" & gsGetPValue()
'
'  Exit Sub
'
'Erro:
'    MsgBox "Erro na função 'SetPrinterModeloPwd2 " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
'End Sub
'
'Public Sub SetPrinterModeloPwd2(ByRef objReport As CrystalReport)
'  ' Modelo 2
'  objReport.Connect = "MS Access;pwd=" & gsGetPValue() & ";Provider=Microsoft.Jet.OLEDB.4.0"
'
'  'Este não funciona...
'  'objReport.Password = Chr$(10)  & gsGetPValue()
'
'  Exit Sub
'
'Erro:
'    MsgBox "Erro na função 'SetPrinterModeloPwd2 " & Err.Number & " " & Err.Description, vbInformation, "Atenção"
'End Sub


'28/07/2003 - mpdea
'Comentado configuração do objeto objReport devido que o report não permance
'com as propriedades padrões, por exemplo, alterando a orientação.
'
'25/07/2003 - mpdea
'Reestruturado código e implementado o funcionamento da opção para relatórios
Public Sub SetPrinterName(ByVal strPrinterType As String, Optional ByRef objReport As CrystalReport)

  
  '0-Não; 1-JáChamou
  If strPrinterType = "REL" And gSetPrinterName_jaChamou_REL = 1 And giQuick_viaRDP <> 1 And giQuick_viaRDP_ticket <> 1 Then
    'Set objReport = gObjReport_Global_REL
    Exit Sub
  ElseIf strPrinterType = "NOTA" And gSetPrinterName_jaChamou_NOTA = 1 And giQuick_viaRDP <> 1 And giQuick_viaRDP_ticket <> 1 Then
    'Set objReport = gObjReport_Global_NOTA
    Exit Sub
  ElseIf strPrinterType = "TICKET" And gSetPrinterName_jaChamou_TICKET = 1 And giQuick_viaRDP <> 1 And giQuick_viaRDP_ticket <> 1 Then
    'Set objReport = gObjReport_Global_TICKET
    Exit Sub
  ElseIf strPrinterType = "CHEQUE" And gSetPrinterName_jaChamou_CHEQUE = 1 And giQuick_viaRDP <> 1 And giQuick_viaRDP_ticket <> 1 Then
    'Set objReport = gObjReport_Global_CHEQUE
    Exit Sub
  ElseIf strPrinterType = "BOLETO" And gSetPrinterName_jaChamou_BOLETO = 1 And giQuick_viaRDP <> 1 And giQuick_viaRDP_ticket <> 1 Then
    'Set objReport = gObjReport_Global_BOLETO
    Exit Sub
  ElseIf strPrinterType = "CARNÊ" And gSetPrinterName_jaChamou_CARNE = 1 And giQuick_viaRDP <> 1 And giQuick_viaRDP_ticket <> 1 Then
    'Set objReport = gObjReport_Global_CARNE
    Exit Sub
  End If

  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome(6) As String
  Dim strNomeLPT(6) As String
  Dim strPortaLPT(6) As String
  Dim intX As Integer
  
  If Len(Trim(strPrinterType)) = 0 Then
    Exit Sub
  End If
  
  strNome(0) = "REL"
  strNomeLPT(0) = "NOME IMPRESSORA REL"
  strPortaLPT(0) = "PORTA IMPRESSORA REL"
  
  strNome(1) = "NOTA"
  strNomeLPT(1) = "NOME IMPRESSORA NOTA"
  strPortaLPT(1) = "PORTA IMPRESSORA NOTA"
  
  strNome(2) = "TICKET"
  strNomeLPT(2) = "NOME IMPRESSORA TICKET"
  strPortaLPT(2) = "PORTA IMPRESSORA TICKET"
  
  strNome(3) = "CHEQUE"
  strNomeLPT(3) = "NOME IMPRESSORA CHEQUE"
  strPortaLPT(3) = "PORTA IMPRESSORA CHEQUE"
  
  strNome(4) = "BOLETO"
  strNomeLPT(4) = "NOME IMPRESSORA BOLETO"
  strPortaLPT(4) = "PORTA IMPRESSORA BOLETO"
  
  strNome(5) = "CARNÊ"
  strNomeLPT(5) = "NOME IMPRESSORA CARNÊ"
  strPortaLPT(5) = "PORTA IMPRESSORA CARNÊ"
  
'  sMENSAGEM_LOG_TESTE_GERAL = sMENSAGEM_LOG_TESTE_GERAL & "STEP 2: " & Now & vbCrLf
  
  ' =======================================================================================================
  ' Tratamento para localizar o Driver e porta disponível nesta conexão deste USUARIO LOGADO no dataCenter
  Dim iIndice As Integer
  Dim sDeviceName As String
  Dim sDeviceNameOriginal As String
  Dim sPorta As String
  Dim sTS_SessionId As String
  Dim X As Printer

  If giQuick_viaRDP = 1 Then
      If strPrinterType = "REL" Then
          strImpressora = GetSetting("QuickStore", "ConfigLPT", "NOME IMPRESSORA REL", "")
      ElseIf strPrinterType = "NOTA" Then
          strImpressora = GetSetting("QuickStore", "ConfigLPT", "NOME IMPRESSORA NOTA", "")
      ElseIf strPrinterType = "TICKET" Then
          strImpressora = GetSetting("QuickStore", "ConfigLPT", "NOME IMPRESSORA TICKET", "")
      ElseIf strPrinterType = "CHEQUE" Then
          strImpressora = GetSetting("QuickStore", "ConfigLPT", "NOME IMPRESSORA CHEQUE", "")
      ElseIf strPrinterType = "BOLETO" Then
          strImpressora = GetSetting("QuickStore", "ConfigLPT", "NOME IMPRESSORA BOLETO", "")
      ElseIf strPrinterType = "CARNÊ" Then
          strImpressora = GetSetting("QuickStore", "ConfigLPT", "NOME IMPRESSORA CARNÊ", "")
      End If
      
      sTS_SessionId = TerminalServerSessionId

      iIndice = InStr(1, strImpressora, " (")
      If iIndice > 0 Then
          strImpressora = Mid(strImpressora, 1, iIndice - 1)
      End If
      '''MsgBox strImpressora

      For Each objPrinter In Printers
          sDeviceNameOriginal = objPrinter.DeviceName
          '''sDeviceNameOriginal = sDeviceNameOriginal & vbCrLf & objPrinter.DeviceName
          sDeviceName = objPrinter.DeviceName
          sPorta = objPrinter.Port

          iIndice = InStr(1, sDeviceName, " (")
          If iIndice > 0 Then
              sDeviceName = Mid(sDeviceName, 1, iIndice - 1)
          End If

          If sDeviceName = strImpressora Then
              If InStr(1, sDeviceNameOriginal, sTS_SessionId) > 0 Then
                  '''MsgBox sDeviceNameOriginal & " " & sPorta
                  
                  Set Printer = objPrinter

                  If Not objReport Is Nothing Then
                      'Seta a impressora para relatório
                      With objReport
                          .PrinterDriver = Printer.DriverName
                          .PrinterName = Printer.DeviceName
                          .PrinterPort = Printer.Port
                      End With
                  End If
                  
                  '''MsgBox objReport.PrinterDriver & " " & objReport.PrinterName & " " & objReport.PrinterPort
                  

                  '''Set objReport = objPrinter
                  Exit For
              End If
          End If
      Next objPrinter
  Else
  ' =======================================================================================================
      For intX = 0 To 5
        If strNome(intX) = strPrinterType Then
            strImpressora = GetSetting("QuickStore", "ConfigLPT", strNomeLPT(intX), "")
            strPorta = GetSetting("QuickStore", "ConfigLPT", strPortaLPT(intX), "")
    
            If Len(Trim(strImpressora)) > 0 And Len(Trim(strPorta)) > 0 Then
                For Each objPrinter In Printers
                    If objPrinter.DeviceName = strImpressora And _
                        objPrinter.Port = strPorta Then
            
                        Set Printer = objPrinter
                        Exit For
                    End If
                Next objPrinter
            End If
        End If
      Next intX
  End If
  
'  sMENSAGEM_LOG_TESTE_GERAL = sMENSAGEM_LOG_TESTE_GERAL & "STEP 3: " & Now & vbCrLf
  
  
'''  If strPrinterType = "REL" Then
'''    gSetPrinterName_jaChamou_REL = 1     '0-Não; 1-JáChamou
'''    'Set gObjReport_Global_REL = Printer
'''    Exit Sub
'''  ElseIf strPrinterType = "NOTA" Then
'''    gSetPrinterName_jaChamou_NOTA = 1     '0-Não; 1-JáChamou
'''    'Set gObjReport_Global_NOTA = Printer
'''    Exit Sub
'''  ElseIf strPrinterType = "TICKET" Then
'''    gSetPrinterName_jaChamou_TICKET = 1     '0-Não; 1-JáChamou
'''    'Set gObjReport_Global_TICKET = Printer
'''    Exit Sub
'''  ElseIf strPrinterType = "CHEQUE" Then
'''    gSetPrinterName_jaChamou_CHEQUE = 1     '0-Não; 1-JáChamou
'''    'Set gObjReport_Global_CHEQUE = Printer
'''    Exit Sub
'''  ElseIf strPrinterType = "BOLETO" Then
'''    gSetPrinterName_jaChamou_BOLETO = 1     '0-Não; 1-JáChamou
'''    'Set gObjReport_Global_BOLETO = Printer
'''    Exit Sub
'''  ElseIf strPrinterType = "CARNÊ" Then
'''    gSetPrinterName_jaChamou_CARNE = 1     '0-Não; 1-JáChamou
'''    'Set gObjReport_Global_CARNE = Printer
'''    Exit Sub
'''  End If
  
  
'  'Seta a impressora para relatório
'  If strPrinterType = "REL" Then
'    With objReport
'      .PrinterDriver = Printer.DriverName
'      .PrinterName = Printer.DeviceName
'      .PrinterPort = Printer.Port
'    End With
'  End If
  
End Sub

Public Sub SetPrinterNameCARNE_TESTE(ByVal strPrinterType As String, Optional ByRef objReport As CrystalReport)

  Dim objPrinter As Printer
  Dim strImpressora As String
  Dim strPorta As String
  
  Dim strNome(6) As String
  Dim strNomeLPT(6) As String
  Dim strPortaLPT(6) As String
  Dim intX As Integer
  
  If Len(Trim(strPrinterType)) = 0 Then
    Exit Sub
  End If
  
  strNome(0) = "REL"
  strNomeLPT(0) = "NOME IMPRESSORA REL"
  strPortaLPT(0) = "PORTA IMPRESSORA REL"
  
  strNome(1) = "NOTA"
  strNomeLPT(1) = "NOME IMPRESSORA NOTA"
  strPortaLPT(1) = "PORTA IMPRESSORA NOTA"
  
  strNome(2) = "TICKET"
  strNomeLPT(2) = "NOME IMPRESSORA TICKET"
  strPortaLPT(2) = "PORTA IMPRESSORA TICKET"
  
  strNome(3) = "CHEQUE"
  strNomeLPT(3) = "NOME IMPRESSORA CHEQUE"
  strPortaLPT(3) = "PORTA IMPRESSORA CHEQUE"
  
  strNome(4) = "BOLETO"
  strNomeLPT(4) = "NOME IMPRESSORA BOLETO"
  strPortaLPT(4) = "PORTA IMPRESSORA BOLETO"
  
  strNome(5) = "CARNÊ"
  strNomeLPT(5) = "NOME IMPRESSORA CARNÊ"
  strPortaLPT(5) = "PORTA IMPRESSORA CARNÊ"
  

  For intX = 0 To 5
    If strNome(intX) = strPrinterType Then
      strImpressora = GetSetting("QuickStore", "ConfigLPT", strNomeLPT(intX), "")
      strPorta = GetSetting("QuickStore", "ConfigLPT", strPortaLPT(intX), "")
      
      If Len(Trim(strImpressora)) > 0 And Len(Trim(strPorta)) > 0 Then
        For Each objPrinter In Printers
          If objPrinter.DeviceName = strImpressora And _
            objPrinter.Port = strPorta Then

            Set Printer = objPrinter
            Exit For
          End If
        Next objPrinter
      End If
    End If
  Next intX
  
'  'Seta a impressora para relatório
'  If strPrinterType = "REL" Then
'    With objReport
'      .PrinterDriver = Printer.DriverName
'      .PrinterName = Printer.DeviceName
'      .PrinterPort = Printer.Port
'    End With
'  End If
  
End Sub

Public Sub ResetPrinter()
  gsInitPrinter = Chr$(27) & Chr$(64)                              'RESET
  gsInitPrinter = gsInitPrinter & Chr$(27) & Chr$(81) & Chr$(160)  'MARGEM DIREITA = 160
  gsInitPrinter = gsInitPrinter & Chr$(27) & Chr$(79)              'CANCELA SALTO SOBRE O PICOTE
  'gsInitPrinter = Chr$(27) & Chr$(64) & Chr$(27) & Chr$(120) & Chr$(0)   'Esc @ (Reset)  + Esc x 0 (Modo Draft)
End Sub

Public Function SetOitavoPrinter(ByVal Filial As Integer) As Integer
  Dim rsParametros As Recordset
  Dim Num_cod As Integer
  Dim Resposta As Integer
  
  On Error GoTo ErrHandler
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Filial
  If rsParametros.NoMatch Then
    SetOitavoPrinter = 1
    Exit Function
  End If
  
  Rem Impressão em 1/8"
  Num_cod = 0
  If rsParametros("Cód Oitavo 1") <> "" Then
    Num_cod = 1
    If rsParametros("Cód Oitavo 2") <> "" Then
      Num_cod = 2
      If rsParametros("Cód Oitavo 3") <> "" Then
        Num_cod = 3
      End If
    End If
  End If
  If Num_cod = 1 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Oitavo 1")))
  End If
  If Num_cod = 2 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Oitavo 1")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Oitavo 2")))
  End If
  If Num_cod = 3 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Oitavo 1")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Oitavo 2")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Oitavo 3")))
  End If
  
  'SetOitavoPrinter = SetPrinterCommand(gsInitPrinter)
  SetOitavoPrinter = 0
  
  Exit Function
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ajustar o modo 1/8 para impressora."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetOitavoPrinter = 1
  Exit Function

End Function

Public Function SetComprimPagLinPrinter(ByVal Filial As Integer, ByVal nComprPag As Integer) As Integer
  Dim rsParametros As Recordset
  Dim Num_cod As Integer
  Dim Resposta As Integer
  
  On Error GoTo ErrHandler
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Filial
  If rsParametros.NoMatch Then
    SetComprimPagLinPrinter = 1
    Exit Function
  End If
  
  Num_cod = 0
  If rsParametros("Cód Comprim 1") <> "" Then
    Num_cod = 1
    If rsParametros("Cód Comprim 2") <> "" Then
      Num_cod = 2
    End If
  Else
    DisplayMsg "Nenhum comando padrão Epson (Códigos para impressão) definido em Parâmetros da Empresa/Filial."
    SetComprimPagLinPrinter = 1
    Exit Function
  End If
  If Num_cod = 1 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 1")))
  End If
  If Num_cod = 2 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 1")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 2")))
  End If
  
  gsInitPrinter = gsInitPrinter & Chr$(nComprPag)
  
  SetComprimPagLinPrinter = 0
  
  rsParametros.Close
  Set rsParametros = Nothing
  Exit Function
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ajustar o comprimento da página para impressora."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetComprimPagLinPrinter = 1
  Exit Function

End Function

Public Function SetComprimPagPrinter(ByVal Filial As Integer, ByVal nComprPag As Integer) As Integer
  Dim rsParametros As Recordset
  Dim Num_cod As Integer
  Dim Resposta As Integer
  
  On Error GoTo ErrHandler
  
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", Filial
  If rsParametros.NoMatch Then
    SetComprimPagPrinter = 1
    Exit Function
  End If
  
  ' Pegue seq para setting em polegadas
  Num_cod = 0
  If rsParametros("Cód Comprim 1") <> "" Then
    Num_cod = 1
    If rsParametros("Cód Comprim 2") <> "" Then
      Num_cod = 2
      If rsParametros("Cód Comprim 3") <> "" Then
        Num_cod = 3
      End If
    End If
  Else
    DisplayMsg "Nenhum comando padrão Epson (Códigos para impressão) definido em Parâmetros da Empresa/Filial."
    SetComprimPagPrinter = 1
    Exit Function
  End If
  If Num_cod = 1 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 1")))
  End If
  If Num_cod = 2 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 1")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 2")))
  End If
  If Num_cod = 3 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 1")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 2")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comprim 3")))
  End If
  
  gsInitPrinter = gsInitPrinter & Chr$(nComprPag)
  
  SetComprimPagPrinter = 0
  
  rsParametros.Close
  Set rsParametros = Nothing
  Exit Function
  
ErrHandler:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao ajustar o comprimento da página para impressora."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetComprimPagPrinter = 1
  Exit Function

End Function

Public Function SetCompressPrinter(Filial As Integer) As Integer
 Dim rsParametros As Recordset
 Dim Num_cod As Integer
 Dim Resposta As Integer
 
 On Error GoTo ErrComprime
 
 Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
 
 rsParametros.Index = "Filial"
 rsParametros.Seek "=", Filial
 If rsParametros.NoMatch Then
   rsParametros.Close
   Set rsParametros = Nothing
   SetCompressPrinter = 1
   Exit Function
 End If
 
   
 Rem Comprime a impressora
 Num_cod = 0
  If rsParametros("Cód Comp 1") <> "" Then
    Num_cod = 1
    If rsParametros("Cód Comp 2") <> "" Then
      Num_cod = 2
      If rsParametros("Cód Comp 3") <> "" Then
        Num_cod = 3
      End If
    End If
  End If
  
  'Lê o registro para verificar se utiliza configuração de modo compressão
  If CBool(GetSetting("QuickStore", "ConfigLPT", "ConfigCompressionPrinter", True)) Then
    'PADRAO EPSON: coloque a impressora em modo Rascunho para aceitar compressão
    gsInitPrinter = gsInitPrinter & Chr$(27) & Chr$(120) & Chr$(0)   'DRAFT
  End If
  
  If Num_cod = 1 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comp 1")))
  End If
  If Num_cod = 2 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comp 1")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comp 2")))
  End If
  If Num_cod = 3 Then
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comp 1")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comp 2")))
    gsInitPrinter = gsInitPrinter + Chr$(Val(rsParametros("Cód Comp 3")))
  End If
  
  SetCompressPrinter = 0
  
  rsParametros.Close
  Set rsParametros = Nothing
  Exit Function
  
ErrComprime:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao usar modo de compressão na impressora."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  SetCompressPrinter = 1
  rsParametros.Close
  Set rsParametros = Nothing
  Exit Function

End Function

Public Function SetPrinterCommand(ByVal sCommand As String) As Integer
  Dim Resposta As Integer
  If Len(sCommand) > 0 Then
    sCommand = Chr$(Len(sCommand) Mod 256) + Chr$(Len(sCommand) \ 256) + sCommand
    '*** ATENÇÃO: NÃO RETIRE O PRÓXIMO COMANDO OU VC. TERÁ UMA TELA AZUL DO NT
    Printer.Print ""
    DoEvents
    If Not IsWindowsNT() Then
      Resposta = Escape(Printer.hdc, PASSTHROUGH, 0, sCommand, 0&)
    Else
      Resposta = Escape32(Printer.hdc, PASSTHROUGH, 0, sCommand, 0&)
    End If
    If Resposta <= 0 Then
      SetPrinterCommand = 1
    Else
      SetPrinterCommand = 0
    End If
  Else
    SetPrinterCommand = 0
  End If
End Function
