Attribute VB_Name = "modGaveta"
Option Explicit

'14/06/2006 - mpdea
'Modificado utilização do componente Drawer.dll para o MSComm na gaveta Gerbô
'
'02/06/2006 - mpdea
'Incluído a Gaveta Gerbô
'
'22-23/03/2006 - mpdea
'Módulo para utilização com gavetas de dinheiro
'GAVETA MENNO MGI
'------------------------------------------------------------------------------

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

'DECLARACAO DAS FUNÇÕES DA GHDL32.DLL
'PARA INTERFACE COM A GAVETA MENNO MGI
Private Declare Function GavetaConfigura Lib "Ghdl32" (ByVal pulso As Integer, ByVal min As Integer) As Long
Private Declare Function DriverGaveta Lib "Ghdl32" (ByVal p As Integer, ByVal F As Integer) As Long

'Constantes de parametros da funcao DriverGaveta
Private Const GAVETA_INICIALIZA As Integer = 1
Private Const GAVETA_ABRE As Integer = 2
Private Const GAVETA_ESTADO As Integer = 3

'03/04/2006 - mpdea
'Variável teve que ser alterada de integer para string devido a problemas de chamada na dll
'que não executava o comando de abertura da gaveta
Private m_strSerial As String
Private m_blnInicializada As Boolean

'Verifica o uso da gaveta em Venda Rápida
Public Function g_blnUsaGavetaVendaRapida() As Boolean
  Dim strRet As String
  Dim blnUsaGaveta As Boolean
  
  
  On Error GoTo ErrHandler
  
  
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "GAVETA", "UtilizarEmVendaRapida")
    If strRet <> "" Then
      Call IsDataType(dtBoolean, strRet, blnUsaGaveta)
    End If
  End If
  
  g_blnUsaGavetaVendaRapida = blnUsaGaveta
  
  Exit Function
  
ErrHandler:
  MsgBox "Erro ao ler configuração [Gaveta]: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Function

'Realiza a abertura da gaveta
'Antes é necessário iniciá-la
Public Sub AbrirGaveta()
  Dim strRet As String
  
  
  On Error GoTo ErrHandler
  
  
  Screen.MousePointer = vbHourglass
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    '02/06/2006 - mpdea
    'Verifica a Marca
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "GAVETA", "Marca")
    Select Case UCase(strRet)
      Case "GERBO"
        '14/06/2006 - mpdea
        'Inicializa gaveta caso não tenha sido anteriormente
        'e aguarda um intervalo de 6 segundos para a abertura
        If Not m_blnInicializada Then InicializaGaveta: Sleep 6000
        'Modificado esquema na abertura da gaveta (Drawer.dll -> MSComm)
        With frmMain.MSComm1
          'Verifica se a gaveta não está aberta
          If Not .CTSHolding Then
            'Abre gaveta
            .RTSEnable = True
            Call Sleep(200)
            .RTSEnable = False
          End If
        End With
      
      Case Else 'Padrão: Menno MGI
        'Inicializa gaveta caso não tenha sido anteriormente
        If Not m_blnInicializada Then InicializaGaveta
        'Abre gaveta
        DriverGaveta m_strSerial, GAVETA_ABRE
    End Select
  End If
  Screen.MousePointer = vbDefault
  
    
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao abrir a gaveta: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

'Inicializa e configura gaveta
Public Sub InicializaGaveta()
  Dim strRet As String
  Dim intPulso As Integer
  Dim intMin As Integer
  
  
  On Error GoTo ErrHandler
  
  
  'Verifica inicialização da gaveta
  If m_blnInicializada Then Exit Sub
  
  
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    'Porta serial de comunicação
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "GAVETA", "Serial")
    If strRet <> "" Then
      Call IsDataType(dtInteger, strRet, m_strSerial)
    End If
    
    '14/06/2006 - mpdea
    'Verifica a Marca
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "GAVETA", "Marca")
    Select Case UCase(strRet)
      Case "GERBO"
        With frmMain.MSComm1
          'Verifica porta COM aberta
          If .PortOpen Then .PortOpen = False
          'Seta a porta COM
          .CommPort = CInt(m_strSerial)
          'Abre a porta COM
          .PortOpen = True
        End With
      
      Case Else 'Padrão: Menno MGI
        'Inicializa a gaveta
        DriverGaveta m_strSerial, GAVETA_INICIALIZA
        'Pulso
        strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "GAVETA", "Pulso")
        If strRet <> "" Then Call IsDataType(dtInteger, strRet, intPulso)
        'Min
        strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "GAVETA", "Min")
        If strRet <> "" Then Call IsDataType(dtInteger, strRet, intMin)
        'Configura gaveta
        GavetaConfigura intPulso, intMin
    
    End Select
    
  End If
  
  'Flag indicando inicialização da gaveta
  m_blnInicializada = True
  
  Exit Sub
  
ErrHandler:
  MsgBox "Erro ao inicializar a gaveta: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub
