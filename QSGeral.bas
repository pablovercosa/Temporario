Attribute VB_Name = "modQSGeral"
Option Explicit

'11/05/2004 - mpdea
'Objeto a ser desabilitado durante a exibição da tela de aviso de bloqueio (opcional)
Public g_objAvisoBloqueioDisabledObject As Object

'16/01/2006 - mpdea
'Tela de Venda Rápida
Public g_frmVendaRapida As Form

Public Const NMAXREGDEMO = 20

'Variável para checar a base de dados independente da versão
Private gbCheckDB As Boolean

'Em modo teste
Public gbTeste As Boolean

'Flag para verificação do sistema em modo completo ou limitado
Public gblnQuickFull As Boolean

Public gDB_SQLSERVER As New ADODB.Connection
Public ws_SQLSERVER As Workspace
Public db_SQLSERVER As Database
Public ws As Workspace
Public db As Database
Public dbFoo As Database
Public wsTemp As Workspace
Public dbTemp As Database

'Para uso com o DataReport
Private connRpt As ADODB.Connection
Public rsRdp As ADODB.Recordset

Public gbWin9X As Boolean

Public frmMain As Form

Public rsProdutos As Recordset
Public rsProdutosNome As Recordset

Public gbAppStarting As Boolean
Public gsAppVersion As String

Public gsCurrentUsers() As String
Public gnCtCurrentUsers As Integer
Public gnMaxUsers As Integer
Public gsMainCaption As String
Public gbDemoVersion As Boolean

Public gsNomeEmpresa As String
Public gsCGCCPF As String
Public gnCodFilial As Integer
Public gsNomeFilial As String
Public gsFilialEndereco As String
Public gsFilialBairro As String
Public gsFilialCep As String
Public gsFilialCidadeEstado As String
Public gsFilialFone As String
Public gnUserCode As Integer
Public gsUserName As String
Public gbSuperUser As Boolean
Public gsCurrencySymbol As String
Public gnCurrencyDecimals As Integer
Public gsCurrencyDecimal As String
Public gsCurrencySeparator As String
Public gsCurrencyFormat As String

Public gsCodCliente As String

Public gsCodProduto As String
Public gsTipoProduto As String

Public gsVendedorVR As String
Public gsTBLayOutFileName As String

Public gbLoginDone As Boolean

Public gnDeltaTime As Long

Public gsTitle As String
Public gsMsg As String
Public gnStyle As String
Public gnResponse As String

Public gsInitPrinter As String

Public gsBuffer As String

'31/07/2002 - mpdea
'Campo de utilização da Loja Virtual
Public gblnWorkWeb As Boolean

Public gbGrade As Boolean
Public gbAcertaGrade As Boolean

Public gsDefaultPath As String
Public gsConfigFileName As String
Public gsOldDBFileName As String
Public gsQuickDBFileName As String
Public gsTempDBFileName As String
Public gsQuickTMPFileName As String
Public gsQuickLicFileName As String
Public gsConsLicFileName As String
Public gsLayOutFileName As String
Public gsHelpFileName As String
Public gsHelpConv As String
Public gsTipFile As String
Public gsHelpContextFileName As String
Public gsGeradorReportFileName As String
Public gsGeradorRecEstadual As String
Public gsBackupFileName As String
'19/11/2009 - mpdea
Public gsCockpitFilename As String

'26/05/2004 - Daniel
'Var que controlará o cadastro de produtos com 0 'zero'
'a esquerda ou não, Parâmetros Filial.[Zero a Esquerda]
Public gbZeroEsquerda As Boolean


'Impressoras
Public gbIsEpson As Boolean

Public gsNumSerie As String
Public gnNumConvenio As Integer

Public gsConfigPath As String
Public gsReportPath As String
Public gsImagePath As String

Public gbPodeGravar As Boolean
Public gbPodeApagar As Boolean

Public gbProdutoRegistrado As Boolean
Public gbToCancel As Boolean
Public gbError As Boolean
Public gbLoading As Boolean
Public gbMyProgramIsRegistered As Boolean

Public gbFirstCFOP As Boolean

Public gnDesconto As Double

Public gbCopyPermissoes As Boolean
Public gsCodigoFrom As String

Public gnCurrHRes As Long
Public gnCurrVRes As Long
'Pega Num do Field //José
Public gnNum As Integer

'Maikel
Public nChamaConsulta As Single ' 1 para VendaRapida
                                ' 2 para Saida
                                ' 3 para Entrada
                                ' 0 Cadastro de produtos
                                ' 4 para cadastroProdutosCFOPs
                                ' 5 Produto Cestas
                                ' 6 para Tela TransferenciaEntreEmpresas
                                ' 7 para Tela de Cadastro da lista de Etiquetas
                                

Public Const CB_SHOWDROPDOWN = &H14F
Public Const BITSPIXEL As Long = 12
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10

' Declares de DLLs usados pela aplicacao
'Public Declare Function extenso Lib "Extens32.dll" _
'    (ByVal Valor As String, _
'     ByVal Retorno As String) As Integer

Public Declare Function GetTickCount Lib "Kernel32" () As Long

Public Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hdc As Long, _
    ByVal nIndex As Long) As Long
   
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Long) As Long

Public Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

' 32-bit Type declaration
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'

Public Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessID As Long
  dwThreadID As Long
End Type

Public Declare Function WaitForSingleObject Lib "Kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long
   
Public Declare Function CreateProcessA Lib "Kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long
   
Public Declare Function CloseHandle Lib "Kernel32" (ByVal _
   hObject As Long) As Long
   
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&


'<<<<<Início do antigo módulo Variáveis>>>>>>>

'Para Win 9.X
Public Declare Function Escape Lib "gdi32" (ByVal hdc%, _
      ByVal nEscape%, ByVal nCount%, _
      ByVal indata$, ByVal outdata As Any) As Long
      
'Para Win NT
Public Declare Function Escape32 Lib "gdi32" Alias "Escape" (ByVal hdc As Long, _
      ByVal nEscape As Long, ByVal nCount As Long, _
      ByVal lpInData As String, lpOutData As Any) As Long

Public Const PASSTHROUGH = 19


Public Formato_Preço As String

Public ESPECIAL As Integer
Public Data_Atual As Variant

Public Menu_Atual As String

Public Num_Usu As Integer
Public Reg_Nome As String

Public gsSenhaGerente As String
Public Filial_Liberada As Integer

Public Glob_Conta_Prod As Integer
Public Glob_Conta_Fat As Integer
Public Glob_Conta_Serv As Integer
Public Glob_Conta_Fatura As Integer
Public gnCtItemProd As Integer
Public gnCtItemServ As Integer
Public gnCtParcFat As Integer
Public Glob_Licenças As Integer

Public gsPesq1 As String
Public gsPesq2 As String
Public gsPesq3 As String

Public Glob_Cod_Alfa As Integer
Public Registrar As Boolean
Public Glob_Mostra_Reg As Boolean

'public gbGrade As Boolean
Public gbEdicao               As Boolean
Public gbServico              As Boolean
Public gbCaixas               As Boolean
Public gbSaldoAnterior        As Boolean '20/11/2006 - ANDERSON - Considerar saldo anterior
Public gbVendedorSenhaGerente As Boolean '17/01/2006 - Anderson - Solicitar senha do gerente ao alterar vendedor nas telas de cadastro de clientes, venda rápida, saídas e check-out

Public nReciboVALOR       As Double
Public nReciboACRESCIMO   As Double
Public nReciboDESCONTO    As Double
'<<<<<Fim do antigo módulo Variáveis>>>>>>>

Public Sub Main()
  Dim F As Form
  Dim dteDataCompilacao As Date
  Dim rstParametros As Recordset
  
  On Error GoTo ProcessErr
  
  If UCase(command$) = "CHECKDB" Then
    gbCheckDB = True
  End If
  
  'Inicializa a variável de teste do sistema
  #If QS_TESTE Then
    gbTeste = True
    Call DisplayMsg("Quick Store em modo teste. Uso exclusivo do Suporte INFOPAR.")
  #End If
  
  'Verifica se o executável é beta e se o período de testes está vencido
  '---------------------------------------------------------------------
  #If BETA = 1 Then
    dteDataCompilacao = FileDateTime(App.Path & "\" & App.EXEName & ".exe")
    
    If (CDate(Date) - CDate(dteDataCompilacao)) > 30 Then
      MsgBox "Atenção, esse executável que você está executando é uma versão beta e o tempo de testes dele expirou. Favor entrar em contato com a Infopar.", vbCritical, "Quick Store"
      End
    End If
  #End If
  '---------------------------------------------------------------------
  
  Call InitWorld
  
  'frmAbout.SplashOn 5000, gsNomeEmpresa, gsCGCCPF
  frmAbout2.SplashOn 5000, gsNomeEmpresa, gsCGCCPF
  '
  Screen.MousePointer = vbHourglass
  Call gnTimeElapsed(-1)
  Set rsProdutos = db.OpenRecordset("SELECT * FROM Produtos WHERE [Código] <> '0' ORDER BY [Código Ordenação]", dbOpenDynaset)
  Set rsProdutosNome = db.OpenRecordset("SELECT Nome, Código FROM Produtos WHERE [Código] <> '0' ORDER BY Nome", dbOpenDynaset)
  Screen.MousePointer = vbDefault
  gnDeltaTime = gnTimeElapsed(0)
  
  'frmAbout.SplashOff
  frmAbout2.SplashOff
  
  Set F = New frmLogin
  F.Show vbModal
  
  Set rstParametros = db.OpenRecordset(" SELECT [Verifica Agenda] FROM " & _
                                       " [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenSnapshot)
  With rstParametros
    If .Fields("Verifica Agenda") = True Then
      Call Verifica_Pendências
      If frmAgenda.lstPend.ListCount > 0 Then
        frmAgenda.Show vbModal
      End If
    End If
    .Close
    Set rstParametros = Nothing
  End With
  
  Set frmMain = New mdiMain
  
  Load frmMain
  DoEvents
'  Load frmProdutos
'  frmProdutos.Visible = False
'  frmProdutos.MoveData
'   Unload frmProdutos

  frmMain.Show
  Exit Sub
  
ProcessErr:
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao iniciar o programa."
  gsMsg = gsMsg & vbCrLf & CStr(Err.Number) & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  If Err.Number <> 3021 Then
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  End
End Sub
  
Public Sub InitWorld()
  Dim nRet As Integer
  Dim rs As Recordset
  Dim sTexto As String
  Dim strRet As String
  

  ' Flag para o comportamento duplo do botão "SAIR" da tela do Logon
  gbAppStarting = True
  
  ' Verifica se o Ano da Data do Sistema tem 4 digitos
  If Len(CStr(Date)) <> 10 Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(118)
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,4")
    End
  End If
  
  ' Verifica se existe ao menos uma impressora instalada no sistema
  If Not bPrinterIsInstalled Then
    gsTitle = LoadResString(201)
    gsMsg = LoadResString(119)
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    End
  End If
  
  Screen.MousePointer = vbHourglass
  
  
  '11/05/2004 - mpdea
  'Objeto a ser desabilitado durante a exibição
  'da tela de aviso de bloqueio (opcional)
  Set g_objAvisoBloqueioDisabledObject = frmMain
  
  
  'Path da aplicação
  gsDefaultPath = App.Path
  If Right(gsDefaultPath, 1) <> "\" Then
    gsDefaultPath = gsDefaultPath & "\"
  End If
  
  
  '---------------------------------------------------------------------------------
  '20/04/2005 - mpdea
  'Configurações adicionais do sistema
  '
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    'Path da aplicação
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "SISTEMA", "DefaultPath")
    If strRet <> "" Then gsDefaultPath = strRet
    
    '03/11/2005 - mpdea
    'KEY: ODBC
    'DSN para conexão ODBC com a base de dados Quick Store
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "SISTEMA", "DSN_QuickStore")
    If strRet <> "" Then
      g_str_dsn_quickstore = strRet
      'Flag indicando o uso de conexão ODBC com o sistema
      g_bln_odbc = True
    End If
  End If
  '---------------------------------------------------------------------------------
  
  
  gsOldDBFileName = gsDefaultPath & "Shop5.mdb"
  gsQuickDBFileName = gsDefaultPath & "QuickStore.mdb"
  gsQuickTMPFileName = gsDefaultPath & "QSTemp.tmp"
  gsTempDBFileName = gsDefaultPath & "Temp.mdb"
  gsQuickLicFileName = gsDefaultPath & "QuickStore.lic"
  gsConsLicFileName = gsDefaultPath & "QuickLic.exe"
  gsGeradorReportFileName = gsDefaultPath & "Gerador.exe"
  gsGeradorRecEstadual = gsDefaultPath & "InfoICMS.exe"
  gsBackupFileName = gsDefaultPath & "Backup.exe"
  '19/11/2009 - mpdea
  gsCockpitFilename = gsDefaultPath & "QuickCockpit.exe QS7QC.IA3"
  
  '---------------------------------------------------------------------------------
  '20/04/2005 - mpdea
  'Configurações adicionais do sistema
  '
  If Dir(gsDefaultPath & "CONFIG.INI") <> "" Then
    'Path das bases de dados
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "SISTEMA", "DatabasePath")
    If strRet <> "" Then
      gsOldDBFileName = strRet & "Shop5.mdb"
      gsQuickDBFileName = strRet & "QuickStore.mdb"
      gsQuickTMPFileName = strRet & "QSTemp.tmp"
      gsTempDBFileName = strRet & "Temp.mdb"
      gsQuickLicFileName = strRet & "QuickStore.lic"
      gsConsLicFileName = strRet & "QuickLic.exe"
      gsGeradorReportFileName = strRet & "Gerador.exe"
      gsGeradorRecEstadual = strRet & "InfoICMS.exe"
      gsBackupFileName = strRet & "Backup.exe"
      '19/11/2009 - mpdea
      gsCockpitFilename = strRet & "QuickCockpit.exe QS7QC.IA3"
    End If
    
    '20/06/2007 - Anderson
    'Configurações para exportação de dados em Excel
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "EXPORTAR_EXCEL", "SavePathEntrada")
    If strRet <> "" Then gsSaveExcelEntrada = strRet
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "EXPORTAR_EXCEL", "SavePathSaida")
    If strRet <> "" Then gsSaveExcelSaida = strRet
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "EXPORTAR_EXCEL", "ArquivoEntrada")
    If strRet <> "" Then gsArquivoExcelEntrada = strRet
    strRet = gstrReadIniFile(gsDefaultPath & "CONFIG.INI", "EXPORTAR_EXCEL", "ArquivoSaida")
    If strRet <> "" Then gsArquivoExcelSaida = strRet
  End If
  '---------------------------------------------------------------------------------
  
  
  gsHelpFileName = gsDefaultPath & "Ajuda\QuickStore.chm"
  gsHelpContextFileName = gsDefaultPath & "Ajuda\Context.txt"
  
  gsConfigFileName = GetWindowsDir() & "QSConfig.TXT"   'Antigo CNFSHOP5.TXT
  gsConfigPath = gsDefaultPath & "Config\"
  gsReportPath = gsDefaultPath & "Reports\"
  gsImagePath = gsDefaultPath & "Imagens\"
  
  
  '03/11/2005 - mpdea
  'KEY: ODBC
  'Somente em conexões Access realiza as alterações na base de dados
  If Not g_bln_odbc Then
    'Somente prossiga se o Quick Store tem as mudanças de Base necessárias
    If Not gbNewStringDB() Then
      Screen.MousePointer = vbDefault
      End
    End If
  End If
  
  '
  On Error GoTo 0
  
  nRet = gnOpenDB(gsQuickDBFileName, False, False)
  If nRet = -1 Then
    End
  End If
    
  Set rs = db.OpenRecordset("ZZZ", dbOpenDynaset)
  On Error Resume Next
  gsCGCCPF = rs("CGCCPF")
  gsNomeEmpresa = Trim(rs("Nome").Value)
  
  
  ' Pilatti 18/08/2017
  ' Tratamento para liberar uso do arquivo de licenças
  'Carrega Licenças do Produto e o último Número de Série do Produto
  'gnMaxUsers = gnReadQuickLic("QS")                          ---------- Pilatti comentou a linha
  gnMaxUsers = 20                                              ' Inclui estas 4 linhas
  gbDemoVersion = False
  gsNumSerie = "QS73063-54"
  gblnQuickFull = True
  'Pilatti 18/08/2017
  
  Call GetMDIMainCaption
  
  
  '22/01/2003 - mpdea
  'Verifica convênio e informações associadas
  Call GetGlobals
  
  
  '19/09/2005 - mpdea
  'Carrega os cases CheckSerialCaseMod globais
  Call LoadCases_CheckSerialCaseMod
  
  '10/09/2007 - Anderson
  'Informa o caminho para gerar arquivo log do sistema
  If g_bolSystemLog Then
    g_strArquivoSystemLog = gsDefaultPath & "System.rec"
  End If
  
  If IsProdutoRegistrado() Then
    If Not gbDemoVersion Then
      If InStr(1, gsCGCCPF, "/") > 0 Then
        sTexto = " (CNPJ: "
      Else
        sTexto = " (CPF: "
      End If
      gsNomeEmpresa = Trim(rs("Nome").Value) & sTexto & gsCGCCPF & ")"
    Else
      gsNomeEmpresa = Trim(rs("Nome").Value) & " (VERSÃO DE DEMONSTRAÇÃO)"
    End If
  Else
    If gbDemoVersion Then
      gsNomeEmpresa = Trim(rs("Nome").Value) & " (VERSÃO DE DEMONSTRAÇÃO)"
    Else
      gsNomeEmpresa = Trim(rs("Nome").Value) & " (CÓPIA NÃO REGISTRADA)"
    End If
  End If
  
  Screen.MousePointer = vbDefault
  
  nRet = gnOpenDB(gsQuickDBFileName, False, False)
  If nRet = -1 Then
    End
  End If
        
  nRet = gnOpenTempDB(gsTempDBFileName, False)
  If nRet = -1 Then
    End
  End If
     
End Sub

Public Function gsHandleNull(ByVal vField As Variant) As String
  On Error Resume Next
  vField = Trim(vField)
  If IsNull(vField) Or IsEmpty(vField) Then
    gsHandleNull = "0"
  Else
    If vField = "" Or vField = " " Then
      gsHandleNull = "0"
    Else
      If vField = Left(Format("1", "Currency"), 2) Or vField = "%" Then
        gsHandleNull = "0"
      Else
        If InStr(vField, gsCurrencySymbol) > 0 Then
          gsHandleNull = Trim(CDbl(vField))
        Else
          gsHandleNull = CStr(vField)
        End If
      End If
    End If
  End If
  On Error GoTo 0
End Function

'24/03/2003 - mpdea
'Alterado conversão CCur para CDbl devido a erro em alguns sistemas operacionais
'e retornar valores válidos como zero
Public Function gsFormatCurrency(ByVal sMoney As Variant, ByVal bZeroIsBlank As Boolean) As String
  Dim sValor As String
  Dim sValor2 As String
  
  sValor = gsHandleNull(sMoney & "")
  If Not IsNumeric(sValor) Then
    gsFormatCurrency = "0"
    Exit Function
  End If
  If bZeroIsBlank = True Then
    'If CCur(sValor) = 0 Then
    If CDbl(sValor) = 0 Then
      gsFormatCurrency = "0"
      Exit Function
    End If
  End If
  If sValor = "0" Then
    sValor = "0.00"
  End If
  
'  sValor2 = FormatCurrency(CCur(sValor), gnCurrencyDecimals)
  sValor2 = FormatCurrency(CDbl(sValor), gnCurrencyDecimals)
  gsFormatCurrency = Replace(sValor2, gsCurrencySymbol, "", 1)
  
End Function

Public Function bCloseAllForms() As Boolean
  Dim F As Form
  For Each F In Forms
    If F.Name <> "mdiMain" Then
      If F.MDIChild And F.Visible Then
        Unload F
      End If
    End If
  Next F
  For Each F In Forms
    If F.Name <> "mdiMain" Then
      If F.MDIChild And F.Visible Then
        bCloseAllForms = False
        Exit Function
      End If
    End If
  Next F
  bCloseAllForms = True
End Function

Public Sub ActiveBarLoadToolTips(ByRef F As Form)
  On Error Resume Next
  With F.ActiveBar1
    .Tools("miOpUpdate").Enabled = gbPodeGravar
    .Tools("miOpDelete").Enabled = gbPodeApagar
    .Bands("tbrOperacoes").Tools("miOpClear").ToolTipText = "Novo (CTRL+N)"
    .Bands("tbrOperacoes").Tools("miOpUpdate").ToolTipText = "Gravar (CTRL+G)"
    .Bands("tbrOperacoes").Tools("miOpDelete").ToolTipText = "Apagar (CTRL+A)"
    .Bands("tbrOperacoes").Tools("miOpFirst").ToolTipText = "Primeiro (F9)"
    .Bands("tbrOperacoes").Tools("miOpPrevious").ToolTipText = "Anterior (F10)"
    .Bands("tbrOperacoes").Tools("miOpNext").ToolTipText = "Próximo (F11)"
    .Bands("tbrOperacoes").Tools("miOpLast").ToolTipText = "Último (F12)"
    .RecalcLayout
    .Refresh
  End With
  On Error GoTo 0
End Sub

Public Function gsGetInvDate(ByVal sDate As String) As String
  On Error Resume Next
  gsGetInvDate = " #" & Mid(sDate, 4, 2) & "/" & Mid(sDate, 1, 2) & "/" & Mid(sDate, 7, 4) & "# "
  On Error GoTo 0
End Function

Public Function gsGetCurrencySymbol()
  gsGetCurrencySymbol = Replace(FormatCurrency(0, 0), "0", "")
End Function

Public Sub TryGetGridLayOut(ByVal sLayOutFileName As String, _
        ByVal sUserName As String, _
        ByVal grdGrid As SSDBGrid)
  '
  If Dir(sLayOutFileName) <> "" Then
    Call grdGrid.LoadLayout(sLayOutFileName)
  Else
    If Dir(App.Path & "\Config", vbDirectory) = "" Then
      ChDir App.Path
      MkDir "Config"
    End If
    If Not Dir(App.Path & "\Config\" & sUserName, vbDirectory) <> "" Then
      ChDir App.Path & "\Config"
      MkDir sUserName
      ChDir App.Path
    End If
  End If
  '
End Sub

Public Function bPrinterIsInstalled() As Boolean
  Dim sDummy As String

  On Error Resume Next
  sDummy = Printer.DeviceName
  
  If Err.Number Then
    bPrinterIsInstalled = False
  Else
    bPrinterIsInstalled = True
  End If
  
End Function

Sub MountStatusBar(F As Form)
  Dim pnlX As Panel
  
  F.stbStatusBar.Panels.Clear
  
  Set pnlX = F.stbStatusBar.Panels.Add()
  pnlX.MinWidth = 3800
  pnlX.Key = "pnlMSG"
  pnlX.Text = ""
  F.stbStatusBar.Panels("pnlMSG").AutoSize = sbrSpring
  
  Set pnlX = F.stbStatusBar.Panels.Add()
  pnlX.MinWidth = 1000
  pnlX.Key = "pnlFILIAL"
  pnlX.Text = "Filial: " & CStr(gnCodFilial)
  F.stbStatusBar.Panels("pnlFILIAL").AutoSize = sbrNoAutoSize

  Set pnlX = F.stbStatusBar.Panels.Add()
  pnlX.MinWidth = 2200
  pnlX.Key = "pnlUSER"
  pnlX.Text = "Usuário: " & CStr(gnUserCode) & "-" & gsUserName
  F.stbStatusBar.Panels("pnlUSER").AutoSize = sbrNoAutoSize
  
  Set pnlX = F.stbStatusBar.Panels.Add()
  pnlX.MinWidth = 1100
  pnlX.Key = "pnlVERSION"
  pnlX.Text = gsAppVersion
  F.stbStatusBar.Panels("pnlVERSION").AutoSize = sbrNoAutoSize
  
  Set pnlX = F.stbStatusBar.Panels.Add(, , , sbrDate, LoadResPicture(101, vbResBitmap))
  pnlX.Width = 1500
  pnlX.Key = "pnlDATE"
  pnlX.Bevel = sbrInset
  pnlX.Alignment = sbrLeft
  F.stbStatusBar.Panels("pnlDATE").AutoSize = sbrNoAutoSize

  Set pnlX = F.stbStatusBar.Panels.Add(, , , sbrTime, LoadResPicture(102, vbResIcon))
  pnlX.Width = 1100
  pnlX.Key = "pnlTIME"
  pnlX.Bevel = sbrInset
  pnlX.Alignment = sbrLeft
  F.stbStatusBar.Panels("pnlTIME").AutoSize = sbrNoAutoSize

End Sub

   
Sub CenterForm(frm As Form)
   Dim ClientRect As RECT           'Holds the area that the form is to be centered in
   Dim TaskBarRect As RECT           'Holds the TaskBar area if in Win95
   Dim X As Variant        'temp LeftPosition
   Dim y As Variant        'temp TopPosition
   '
   
  frm.KeyPreview = True
  
  ' Check if the form is a MDIChild.
  If frm.MDIChild Then
     '
     ' Center it in the MDIParent.
     GetClientRect GetParent(frm.hwnd), ClientRect
  Else  'Center it in the available desktop area.
     '
     ' Get the Desktop area
     Call GetClientRect(GetDesktopWindow(), ClientRect)
     '
     ' Check for the Task Bar.
     apiRetVal = FindWindow("Shell_TrayWnd", vbNullString)
     '
     ' If there is a taskbar, ie WIN95 then adjust the ClientRect.
     If apiRetVal Then
        Call GetWindowRect(apiRetVal, TaskBarRect)
        '
        If (TaskBarRect.Right - TaskBarRect.Left) > (TaskBarRect.Bottom - TaskBarRect.Top) Then
           '
           ' TaskBar at the Top of Screen.
           If TaskBarRect.Top <= 0 Then
              ClientRect.Top = ClientRect.Top + TaskBarRect.Bottom
           '
           ' TaskBar at the Bottom of Screen.
           Else
              ClientRect.Bottom = ClientRect.Bottom - (TaskBarRect.Bottom - TaskBarRect.Top)
           End If
        Else
           '
           ' TaskBar is on the Left side of the Screen.
           If TaskBarRect.Left <= 0 Then
              ClientRect.Left = ClientRect.Left + TaskBarRect.Right
           '
           ' TaskBar is on the Right side of the Screen.
           Else
              ClientRect.Right = ClientRect.Right - (TaskBarRect.Right - TaskBarRect.Left)
           End If
        End If   '[TaskBar on Top of Screen?]
     End If      '[if apiRetVal]
  End If
  '
  ' Center the Form
  With frm
    X = (((ClientRect.Right - ClientRect.Left) * Screen.TwipsPerPixelX) - .Width) / 2
    y = (((ClientRect.Bottom - ClientRect.Top) * Screen.TwipsPerPixelY) - .Height) / 2
    If .WindowState = vbNormal Then
     .Move X, y
    End If
  End With
End Sub

Public Sub ExecCmd(cmdline$)
  Dim proc As PROCESS_INFORMATION
  Dim Start As STARTUPINFO
  Dim ret&
  ' Initialize the STARTUPINFO structure:
  Start.cb = Len(Start)
  ' Start the shelled application:
  ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
     NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)
  ' Wait for the shelled application to finish:
  ret& = WaitForSingleObject(proc.hProcess, INFINITE)
  ret& = CloseHandle(proc.hProcess)
End Sub

Public Function WaitSeconds(ByVal nSecs As Integer, Optional ByVal bDoEvents As Boolean = True)
  Dim nSecIni As Long
  Dim nSec As Long
  nSecIni = GetTickCount()
  Do
    If bDoEvents Then DoEvents
    nSec = GetTickCount()
    nSec = (nSec - nSecIni) / 1000
    If nSec >= nSecs Then
      Exit Function
    End If
  Loop
End Function

Public Function WaitMilliSeconds(ByVal lngMilliSeconds As Long)
  Dim lngStart As Long
  
  lngStart = GetTickCount()
  Do
    If (GetTickCount() - lngStart) >= lngMilliSeconds Then Exit Function
    DoEvents
  Loop
End Function

Public Function gnTimeElapsed(ByVal nMode As Integer) As Single
  Static nVal As Long
  Dim nVal2 As Long
  If nMode = -1 Then
    nVal = GetTickCount()
    gnTimeElapsed = 0
  Else
    nVal2 = GetTickCount()
    gnTimeElapsed = (nVal2 - nVal) / 1000
  End If
  Exit Function
End Function

Public Function IsClear(ByRef sText As Variant) As Boolean
  IsClear = IsNull(sText)
  If Not IsClear Then
    IsClear = Len(Trim(sText)) = 0
  End If
End Function

Public Function sReplace(sText As String, sTextOut As String, sTextIn As String) As String
  Dim nPos As Integer
  Dim nLenOrig As Integer
  Dim nLen As Integer
  nPos = InStr(sText, sTextOut)
  If nPos <= 0 Then
    sReplace = sText
  Else
    nLenOrig = Len(sText)
    nLen = Len(sTextOut)
    sReplace = Left(sText, nPos - 1) & sTextIn & Right(sText, nLenOrig - (nPos + nLen) + 1)
  End If
End Function

Public Function sGetValueInDecimalPoint(ByVal sValue As String) As String
  sGetValueInDecimalPoint = sReplace(sValue, ",", ".")
End Function

Public Function vFieldVal(rvntFieldVal As Variant) As Variant
  If IsNull(rvntFieldVal) Then
    vFieldVal = vbNullString
  Else
    vFieldVal = CStr(rvntFieldVal)
  End If
End Function

Public Function nValidateDate(sDate As String, dtDateIni As Date, dtDateEnd As Date) As Integer
  Dim dtDate As Date
  If Not IsDate(sDate) Then
    nValidateDate = -1
    Exit Function
  End If
  dtDate = CDate(sDate)
  If dtDate > dtDateEnd Or dtDate < dtDateIni Then
    nValidateDate = -1
    Exit Function
  End If
  nValidateDate = 0
End Function

Public Function gsGetPValue() As String
  Dim A(19) As String
  A(10) = "m"
  A(0) = "x"
  A(1) = "a"
  A(12) = "k"
  A(2) = "y"
  A(3) = "i"
  A(14) = "n"
  A(4) = "Z"
  A(5) = "a"
  A(6) = "q"
  A(16) = "1"
  A(7) = "7"
  A(8) = "&"
  A(18) = "5"
  gsGetPValue = A(4) & A(0) & A(1) & A(12) & A(14) & A(3) & A(6) & A(7) & A(16) & A(18)
End Function

Public Function gsGetPValue2() As String
  Dim A(19) As String
  A(10) = "m"
  A(0) = "x"
  A(1) = "a"
  A(12) = "k"
  A(2) = "y"
  A(3) = "m"
  A(14) = "t"
  A(4) = "Z"
  A(5) = "a"
  A(6) = "p"
  A(16) = "1"
  A(7) = "7"
  A(8) = "&"
  A(18) = "5"
  gsGetPValue2 = A(4) & A(0) & A(1) & A(12) & A(14) & A(3) & A(6) & A(7) & A(16) & A(18)
End Function

Public Sub VerificaPeriodoValido(ByRef sDataIni As String, ByRef sDataFim As String)
         
  If CDate(sDataIni) > CDate(sDataFim) Then
    Dim sData As String
    sData = sDataIni
    sDataIni = sDataFim
    sDataFim = sData
  End If

End Sub

Public Sub SwitchGrid(ByRef ctlControl As Control, ByVal bWhat As Boolean)
  ctlControl.AllowAddNew = bWhat
  ctlControl.AllowDelete = bWhat
  ctlControl.AllowUpdate = bWhat
  ctlControl.Refresh
End Sub

Public Function sFormatZeros(ByVal sText As String, ByVal nSize As Integer) As String
  Dim nLen As Integer
  nLen = Len(Trim(sText))
  If nLen > nSize Then
    sFormatZeros = sText
    Exit Function
  End If
  
  nLen = nSize - nLen
  sFormatZeros = String(nLen, "0") & Trim(sText)
  
End Function

'Public Function sConverte(ByVal nValor As Long, ByVal bFlag As Boolean) As String
'  Dim sValor As String
'  If nValor < 0 Then
'    sConverte = ""
'    Exit Function
'  End If
'
'  On Error GoTo ErrExit
'  sValor = Space(255)
'  Call extenso(CStr(nValor), sValor)
'  sConverte = sValor
'  Exit Function
'
'ErrExit:
'  sConverte = "Valor para conversão com erro."
'  Exit Function
'End Function
'
Public Function bCheckCGC(ByVal sCGC As String) As Boolean
  Dim nI As Integer
  Dim nJ As Integer
  Dim nSum As Integer
  Dim nRem As Integer
  Dim nDig(2) As Integer
  Dim nVerif As Integer
  Dim sFat(2) As String
  Dim sDig As String
  Dim sValorCGC As String

  ' sCGC deve estar livre da mascara, no seguinte formato: NNNNNNNNXXXXDD
  ' Onde XXXX valor apos barra p.ex. 0001
  ' e DD valores dos digitos verificadores
  '
  sValorCGC = ""
  For nI = 1 To Len(sCGC)
    sDig = Mid(sCGC, nI, 1)
    If IsNumeric(sDig) Then
      sValorCGC = sValorCGC & sDig
    End If
  Next nI
  sValorCGC = Right(String(14, "0") & sValorCGC, 14)
  
  sFat(0) = "543298765432"
  sFat(1) = "6543298765432"

  For nI = 0 To 1
    nSum = 0
    For nJ = 1 To Len(sFat(nI))
      nSum = nSum + CInt(Mid(sValorCGC, nJ, 1)) * CInt(Mid(sFat(nI), nJ, 1))
    Next nJ
    nRem = nSum Mod 11
    nDig(nI) = IIf(nRem <= 1, 0, 11 - nRem)
  Next nI

  nVerif = nDig(0) * 10 + nDig(1)

  bCheckCGC = CInt(Mid(sValorCGC, 13, 2)) = nVerif

End Function

Public Function bCheckCPF(ByVal sCPF As String) As Boolean
  Dim nI As Integer
  Dim nJ As Integer
  Dim nN As Integer
  Dim nSum As Integer
  Dim nRem As Integer
  Dim nDig(2) As Integer
  Dim sFat(2) As String
  Dim nVerif As Integer
  Dim sValorCPF As String
  Dim sDig As String
  
  '15/07/2005 - Daniel
  'Adicionado tratamento de erro
  On Error GoTo TratarErro
    
  ' sCPF deve estar livre da mascara, no seguinte formato: NNNNNNNNNDD
  ' Onde DD são os valores dos digitos verificadores
  
  sValorCPF = ""
  For nI = 1 To Len(sCPF)
    sDig = Mid(sCPF, nI, 1)
    If IsNumeric(sDig) Then
      sValorCPF = sValorCPF & sDig
    End If
  Next nI
  sValorCPF = Right(String(11, "0") & sValorCPF, 11)
  
  sFat(0) = "100908070605040302"
  sFat(1) = "11100908070605040302"
  
  For nI = 0 To 1
    nSum = 0
    nN = Len(sFat(nI)) / 2
    For nJ = 1 To nN
      nSum = nSum + CInt(Mid(sCPF, nJ, 1)) * CInt(Mid(sFat(nI), 2 * nJ - 1, 2))
    Next
    nRem = (nSum * 10) Mod 11
    nDig(nI) = IIf(nRem = 10, 0, nRem)
  Next
  
  nVerif = nDig(0) * 10 + nDig(1)
  bCheckCPF = CInt(Mid(sCPF, 10, 2)) = nVerif
  
  Exit Function

TratarErro:
  MsgBox "Erro na execução da função bCheckCPF" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Function

Public Function sTranslateInvalidChar(ByVal sRecord As String) As String
  Dim nI As Integer
  Dim sChar As String * 1
  Dim sNextChar As String * 1
  Dim sTemp As String
  Dim nPos As Integer
  
  If Not IsNumeric(sRecord) And Len(Trim(sRecord)) > 0 Then
  
    For nI = 1 To Len(sRecord)
      
      sNextChar = Mid(sRecord, nI, 1)
      
      Select Case sNextChar
        Case "ç"
          '21/06/2007 - Anderson
          'Alteração na substituição do caractere
          'sChar = "sChar"
          sChar = "c"
        Case "Ç"
          '21/06/2007 - Anderson
          'Alteração na substituição do caractere
          'sChar = "sChar"
          sChar = "C"
        Case "ñ"
          sChar = "n"
        Case "Ñ"
          sChar = "N"
        Case "á", "à", "â", "ä", "ã"
          sChar = "a"
        Case "Á", "À", "Â", "Ä", "Ã"
          sChar = "A"
        Case "é", "è", "ê", "ë"
          sChar = "e"
        Case "É", "È", "Ê", "Ë"
          sChar = "E"
        Case "í", "ì", "î", "ï"
          '22/06/2007 - Anderson
          'Alteração na substituição do caractere
          'sChar = "nI"
          sChar = "i"
        Case "Í", "Ì", "Î", "Ï"
          '22/06/2007 - Anderson
          'Alteração na substituição do caractere
          'sChar = "nI"
          sChar = "I"
        Case "ó", "ò", "ô", "ö", "õ"
          sChar = "o"
        Case "Ó", "Ò", "Ô", "Ö", "Õ"
          sChar = "O"
        Case "ú", "ù", "û", "ü"
          sChar = "u"
        Case "Ú", "Ù", "Û", "Ü"
          sChar = "U"
        Case Else
          sChar = sNextChar
      End Select
      
      'Verifica se é uma letra maiúscula, minúscula ou dígito
      If (((Asc(sChar) <= 64) Or (Asc(sChar) >= 91)) And _
          ((Asc(sChar) <= 96) Or (Asc(sChar) >= 123)) And _
          ((Asc(sChar) <= 47) Or (Asc(sChar) >= 58))) Then
        'Verifica se é um caracter especial
        nPos = InStr("$%""'()*+,&#<>/|-_.;:=?", sChar)
        If nPos = 0 And Not IsNumeric(sChar) Then
          sChar = " "
        End If
      End If
      
      sTemp = sTemp + sChar
    
    Next nI
    
    sTranslateInvalidChar = UCase(sTemp)
  
  Else
  
    sTranslateInvalidChar = sRecord
  
  End If
  
End Function

Public Function sGetCurrencySymbol()
  sGetCurrencySymbol = Trim(Replace(FormatCurrency(0, 0), "0", ""))
End Function


Public Function gnOpenDB_SQLSERVER()
  On Error GoTo ErrHandler
  
  If gDB_SQLSERVER.State = 1 Then
    'DB SQL Server já esta aberto
    Exit Function
  End If

  'CONEXAO SERVIDOR PRODUCAO A3
  gDB_SQLSERVER.Open "PROVIDER = MSDASQL;driver={SQL Server};database=QuickStore;server=AMAZONA-F74E4RM\SQLEXPRESS;uid=sa;pwd=admin@A3;"
  
  'CONEXAO SERVIDOR DESENV QUICK A3
'  gDB_SQLSERVER.Open "PROVIDER = MSDASQL;driver={SQL Server};database=QuickStore;server=WIN2003VB\SQLEXPRESS;uid=sa;pwd=admin@A3;"
 
  Exit Function
ErrHandler:
  MsgBox "Erro na conexão com o DB SQL Server. Cod: " & Err.Number & " >> Desc: " & Err.Description, vbCritical, "Erro de Conexão"

End Function

Public Function gnCloseDB_SQLSERVER()
  On Error GoTo ErrHandler
  
  If gDB_SQLSERVER.State <> 1 Then
    'DB SQL Server já esta fechado
    Exit Function
  End If
  
  gDB_SQLSERVER.Close

  Exit Function
ErrHandler:
  MsgBox "Erro na conexão com o DB SQL Server. Cod: " & Err.Number & " >> Desc: " & Err.Description, vbCritical, "Erro de Conexão"

End Function

Public Function gnOpenDB( _
      ByVal sDB As String, _
      ByVal bExclusive As Boolean, _
      ByVal bCheckLic As Boolean) As Integer
  Dim rs As Recordset
  Dim F As Form
  Dim nI As Integer
  Dim nRet As Integer
  
  On Error GoTo ErrHandler
  
  If Dir(sDB) = "" Then
    gsMsg = "Base de Dados """ & sDB & """ não encontrada."
    gsMsg = gsMsg & " Revise a pasta onde esta aplicação executa e a presença dos arquivos básicos para execução."
    DisplayMsg gsMsg
    End
  End If
  
  Set ws = DBEngine.Workspaces(0)
  
  '03/11/2005 - mpdea
  'KEY: ODBC
  'Abertura do banco de dados
  If g_bln_odbc Then
    Set db = ws.OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DSN=" & g_str_dsn_quickstore)
  Else
    Set db = ws.OpenDatabase(sDB, bExclusive, False, ";pwd=" & gsGetPValue())
  End If
  
  '
  If Dir(gsQuickTMPFileName) = "" Then
    nRet = gnCreateTMPFileName(gsQuickTMPFileName)
    If nRet = -1 Then
      End
    End If
  End If
  Set dbFoo = ws.OpenDatabase(gsQuickTMPFileName, False, False)
  '
  If bCheckLic = True Then
    gsCurrentUsers = gsGetMDBUsers(gsQuickTMPFileName)
    gnCtCurrentUsers = 0
    For nI = LBound(gsCurrentUsers) To UBound(gsCurrentUsers)
      DoEvents
      If Len(gsCurrentUsers(nI)) = 0 Then
        Exit For
      End If
      gnCtCurrentUsers = gnCtCurrentUsers + 1
    Next nI
    If IsProdutoRegistrado() Then
      If (gnCtCurrentUsers <= gnMaxUsers) Or gbDemoVersion Then
        gnOpenDB = 0
      Else
        Screen.MousePointer = vbDefault
        gsTitle = LoadResString(201)
        gsMsg = "Número máximo de usuários atingiu o limite atual de licenças. Feche alguma seção em outra estação da rede ou tente mais tarde."
        gnStyle = vbOKOnly + vbCritical
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        gnOpenDB = -1
      End If
    Else
      If gnCtCurrentUsers <= 1 Then
        gnOpenDB = 0
      Else
        gnOpenDB = -2
      End If
    End If
  Else
    gnCtCurrentUsers = 0
    gnOpenDB = 0
  End If
  
  Screen.MousePointer = vbDefault
  On Error GoTo 0
  Exit Function

ErrHandler:
  Dim strError As String
  Dim errLoop As Error

  Screen.MousePointer = vbDefault

  If bExclusive = False Then
  
    ' Veja se alguem está mantendo o banco exclusivo...
    If Err.Number = 3261 Or Err.Number = 3045 Then
      gsMsg = "Banco de Dados Atualmente em Manutenção. Contate o Administrador."
      gnStyle = vbOKOnly + vbCritical
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      End
    End If
    
    If Err.Number = 3343 Then
      'MsgBox "O Banco de Dados do QuickStore precisa ser corrigido. Na próxima tela clique no botão COMPACTAR e depois REPARAR. Então se logue novamente. (Obs: Causa do problema, o Sistema QuickStore não foi encerrado adequadamente).", vbInformation, "Atenção"
      'Shell App.Path & "\Repara.exe", vbNormalFocus
    
      MsgBox "O Banco de Dados precisa ser corrigido. O QuickStore compactará a base de dados.", vbInformation, "Atenção"
      CompactarBancoDeDados
      
      MsgBox "Agora o QuickStore fará o reparo na base de dados.", vbInformation, "Atenção"
      RepararBancoDeDados
      
      MsgBox "Pronto. Agora se logue novamente no QuickStore.", vbInformation, "Atenção"
      End
    End If
  
    gsMsg = "Abertura do Banco de Dados '" & sDB & "' com Erro. " & vbCrLf
    
    ' Enumerate Errors collection and display properties of
    ' each Error object.
    For Each errLoop In Errors
      With errLoop
        strError = "Erro Número:" & .Number & vbCrLf
        strError = strError & "Descrição: " & .Description & vbCrLf
      End With
    Next
    gsMsg = gsMsg & vbCrLf & strError
    gnStyle = vbOKOnly + vbCritical
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    End

  End If
  
  gnOpenDB = -1

End Function

'
Private Sub RepararBancoDeDados()
  Dim nRet As Integer

  Screen.MousePointer = vbHourglass
    
  DBEngine.RepairDatabase gsQuickDBFileName
'  nRet = gnOpenDB(gsQuickDBFileName, False, True) + gnOpenTempDB(gsTempDBFileName, False)
'  If nRet <> 0 Then
'    End
'  End If
'  Call StatusMsg("")
  Screen.MousePointer = vbDefault

End Sub

Private Sub CompactarBancoDeDados()
  Dim sTempFileName As String
  Dim nRet As Integer
    
  Screen.MousePointer = vbHourglass
  sTempFileName = App.Path & "\TMP" & Format(Time, "HHMMSS") & ".MDB"
  On Error Resume Next
  Kill sTempFileName
  On Error GoTo 0
  DBEngine.CompactDatabase gsQuickDBFileName, sTempFileName, , , ";pwd=" & gsGetPValue()
  Kill gsQuickDBFileName
  Name sTempFileName As gsQuickDBFileName
 
  Screen.MousePointer = vbDefault
      
End Sub


Public Function gnOpenTempDB( _
      ByVal sDB As String, _
      ByVal bExclusive As Boolean) As Integer
  Dim rs As Recordset
  Dim F As Form
  
  On Error GoTo ErrHandler
  
  Set wsTemp = DBEngine.Workspaces(0)
  Set dbTemp = wsTemp.OpenDatabase(sDB, bExclusive, False, ";pwd=" & gsGetPValue2())
  If Err.Number Then
    gnStyle = vbOKOnly + vbCritical
    gsMsg = Err.Description
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    End
  End If
  '
  gnOpenTempDB = 0
  
  On Error GoTo 0
  Exit Function

ErrHandler:
  Dim strError As String
  Dim errLoop As Error

  If bExclusive = False Then
  
    ' Veja se alguem está mantendo o banco exclusivo...
    If Err.Number = 3261 Or Err.Number = 3045 Then
      gsMsg = "Banco de Dados '" & sDB & "' Atualmente em Manutenção. Contate o Administrador."
      gnStyle = vbOKOnly + vbCritical
      gsTitle = "Atenção"
      Call MsgBox(gsMsg, gnStyle, gsTitle)
      End
    End If
  
    gsMsg = "Abertura do Banco de Dados '" & sDB & "' com Erro. "
    gsMsg = gsMsg & vbCrLf & Err.Description
    
    ' Enumerate Errors collection and display properties of
    ' each Error object.
    For Each errLoop In Errors
      With errLoop
        strError = "Erro Número:" & .Number & vbCrLf
        strError = strError & "Descrição: " & .Description & vbCrLf
      End With
    Next
    gsMsg = gsMsg & vbCrLf & strError
    gnStyle = vbOKOnly + vbCritical
    gsTitle = "Atenção"
    Call MsgBox(gsMsg, gnStyle, gsTitle)
    End

  End If
  
  gnOpenTempDB = -1

End Function

Public Function gnCreateTMPFileName(ByVal sDB As String) As Integer
  On Error Resume Next
  Call ws.CreateDatabase(sDB, dbLangGeneral)
  If Err.Number <> 0 Then
    gnCreateTMPFileName = -1
  Else
    gnCreateTMPFileName = 0
  End If
  On Error GoTo 0
End Function

'03/04/2003 - mpdea
'Modificado grupo inicial do convênio [31] para 30 (antes era 31)
'devido a licenças já distribuídas com este inicial
'10/02/2003 - mpdea
'Obtém o convênio do nr. de série
Public Function gintGetConvenio(ByVal strNumSerie As String) As Integer
  Dim strSQL As String
  
  If Len(strNumSerie) = 11 Then
    Select Case CInt(Mid(strNumSerie, 3, 2))
      
      Case 30 To 40   'Clientes Infopar
        gintGetConvenio = 31
      
      Case 41 To 50   'Convenio CDL
        gintGetConvenio = 41
      
      'RESERVADO
      Case 51 To 60   'Convenio FACIAP (não utilizado)
        gintGetConvenio = 51
      
      Case 61 To 70   'Quick Posto
        gintGetConvenio = 61
      
      Case 71 To 80   'Clientes Infopar
        gintGetConvenio = 31 'Mantém ID já utilizado
      
      Case 81 To 90   'Outros Convenios
        gintGetConvenio = 81
      
      Case Else
        gintGetConvenio = 91
        
    End Select
    
  Else
    gintGetConvenio = 31
  End If
  
End Function

'16/01/2003 - mpdea
'Fixado gsHelpConv (arquivo de ajuda para convênios) para convênios não utilizados
'Fixado gsTipFile (arquivo de dicas) para os convênios diferentes de 31 e 41
'Adicionado verificação do modo do sistema (completo ou limitado)
'
'Função relacionada: gintGetConvenio
'
Public Sub GetGlobals()
  Dim strSQL As String
    
  'Versão limitada por padrão
  gblnQuickFull = False
  
  
  'Veja arquivo de Tips e Help segundo o Convênio
  gsTipFile = gsConfigPath & "QS31.dat"
  gsHelpConv = ""
  
'''  If Len(gsNumSerie) = 11 Then
'''    Select Case CInt(Mid(gsNumSerie, 3, 2))
'''
'''      Case 31 To 40   'Clientes Infopar
'''        gnNumConvenio = 31
'''        gblnQuickFull = True
'''
'''      Case 41 To 50   'Convenio CDL
'''        gnNumConvenio = 41
'''        gblnQuickFull = True
'''        gsTipFile = gsConfigPath & "QS41.dat"
'''        gsHelpConv = App.Path & "\Ajuda\" & "QS41.chm"
'''
'''      'RESERVADO
'''      Case 51 To 60   'Convenio FACIAP (não utilizado)
'''        gnNumConvenio = 51
'''
'''      Case 61 To 70   'Quick Posto
'''        gnNumConvenio = 61
'''        gblnQuickFull = True
'''
'''      Case 71 To 80   'Clientes Infopar
'''        gnNumConvenio = 31 'Mantém ID já utilizado
'''        gblnQuickFull = True
'''
'''      Case 81 To 90   'Outros Convenios
'''        gnNumConvenio = 81
'''
'''      Case Else
'''        gnNumConvenio = 91
'''
'''    End Select
'''
'''  Else
    gnNumConvenio = 31
'''  End If
  
  ' Pilatti 18/08/2017
  ' Adicionei a linha abaixo...setando TRUE para a variavel
  gblnQuickFull = True
  ' Pilatti
  
  '03/10/2003 - mpdea
  'Alterado para utilizar serviços
  '
  '22/01/2003 - mpdea
  'Ajustes para o Quick em modo limitado
  If Not gblnQuickFull Then
  
    'Flags
    gbSuperUser = False
    gbGrade = False
    gbEdicao = False
    gbServico = True
    gsPesq1 = ""
    gsPesq2 = ""
    gsPesq3 = ""
  
    
    'Filial padrão em modo limitado
    strSQL = "UPDATE [Parâmetros Filial] SET "
    strSQL = strSQL & "Nome = '" & gsNomeEmpresa & "', "
    strSQL = strSQL & "[Razão Social] = '', "
    strSQL = strSQL & "CGC = '" & gsCGCCPF & "', "
    strSQL = strSQL & "[VR Linhas Digitação] = 255, "
    strSQL = strSQL & "[Nota Saída] = '', "
    strSQL = strSQL & "[Nota Entrada] = '', "
    strSQL = strSQL & "[Consulta Tab1] = 'TABELA1', "
    strSQL = strSQL & "[Consulta Tab2] = 'TABELA2', "
    strSQL = strSQL & "[Consulta Tab3] = 'TABELA3', "
    strSQL = strSQL & "[Verifica Agenda] = False, "
    strSQL = strSQL & "[Usar Grade] = False, "
    strSQL = strSQL & "[Usar Edições] = False, "
    strSQL = strSQL & "[Usar Serviços] = True, "
    strSQL = strSQL & "[Código Banco Cheques] = 0, "
    strSQL = strSQL & "OpSaidaOrcVenda = 500, "
    strSQL = strSQL & "[Mensagem Troca] = '', "
    strSQL = strSQL & "[Mensagem Etiq 1] = '', "
    strSQL = strSQL & "[Mensagem Etiq 2] = '' "
    strSQL = strSQL & "WHERE Filial = 1"
    
    db.Execute strSQL, dbFailOnError
  End If
  
  
End Sub

Private Function gbSetNewStringDB(ByVal sDB As String) As Boolean
  Dim rs As Recordset

Again:
  On Error Resume Next
  db.Close
  dbFoo.Close
  ws.Close
  
  Err.Clear
  
  'Primeiro abra o DB como shareable sem senha...
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(sDB, False, False)
  
  'Se não houve erro, trata-se de um DB não protegido ainda...Proteja-o.
  If Err.Number = 0 Then
    
    On Error GoTo ErrHandler
    
    'Abra o DB em modo exclusivo
    db.Close
    ws.Close
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(sDB, True, False)
    '
    Call db.NewPassword("", gsGetPValue())
    db.Close
    ws.Close
    
    On Error Resume Next
    '
  End If
  '
  Err.Clear
  
  'Veja se existem alterações nesta versão a serem feitas no DB
  DBEngine.Idle dbRefreshCache
  Call WaitSeconds(2)
  db.Close
  ws.Close
  Set db = Nothing
  Set dbFoo = Nothing
  Set ws = Nothing
  DoEvents
  
  On Error GoTo ErrHandler
  
  Set ws = DBEngine.Workspaces(0)
  Set db = ws.OpenDatabase(sDB, True, False, ";pwd=" & gsGetPValue())
  '
  Call AlteraDB
  
  db.Close
  ws.Close
  DoEvents
  Set db = Nothing
  Set ws = Nothing
  
  'Renomeie o .MDB antigo para o novo nome
  If sDB <> gsQuickDBFileName Then
    On Error Resume Next
    Kill gsQuickDBFileName
    Name sDB As gsQuickDBFileName
    On Error GoTo 0
  End If
  
  gbSetNewStringDB = True
  
  Exit Function
  
ErrHandler:
  gsTitle = "Atenção"
  If Err.Number = 3621 Or Err.Number = 3356 Then
    Set db = Nothing
    Set dbFoo = Nothing
    Set ws = Nothing
    gsMsg = LoadResString(70)
    gnStyle = vbYesNo + vbQuestion
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbYes Then
      Err.Clear
      Resume Again
    Else
      End
    End If
  Else
    If Err.Number = 3031 Then
      Resume Next
    Else
      gnStyle = vbOKOnly + vbCritical
      gsMsg = "Erro: " & Err.Number & "-" & Err.Description
    End If
  End If
  Call MsgBox(gsMsg, gnStyle, gsTitle)
  Set db = Nothing
  Set dbFoo = Nothing
  Set ws = Nothing
  End
End Function

Public Sub GetConfigValues()
  Dim nFileNum As Integer
  Dim nPos As Integer
  Dim sRecord As String
  Dim sKeyName As String
  Dim sKeyValue As String
  
  nFileNum = FreeFile
  Open gsConfigFileName For Input As nFileNum
  
  Do While Not EOF(nFileNum)
    
    Line Input #nFileNum, sRecord
    
    nPos = InStr(1, sRecord, ";")
    If nPos > 0 Then
      sKeyName = UCase(Trim(Mid(sRecord, 1, nPos - 1)))
      sKeyValue = UCase(Trim(Mid(sRecord, nPos + 1, 200)))
      
      Select Case sKeyName
        Case "CLIENTES - MANTÉM CIDADE" '                         ;False
          Call UpdateArqConfig("Clientes", "Mantem Cidade", sKeyValue)
        Case "CLIENTES - MANTÉM ESTADO" '                         ;False
          Call UpdateArqConfig("Clientes", "Mantem Estado", sKeyValue)
        Case "CLIENTES - MANTÉM BOLETO" '                          ;False
          Call UpdateArqConfig("Clientes", "Mantem Boleto", sKeyValue)
        Case "NOME IMPRESSORA REL" '                               ;HP DeskJet 720C Series
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "PORTA IMPRESSORA REL" '                              ;LPT1:
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "NOME IMPRESSORA NOTA" '                              ;HP DeskJet 720C Series
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "PORTA IMPRESSORA NOTA" '                             ;LPT1:"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "NOME IMPRESSORA TICKET" '                            ;HP DeskJet 720C Series"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "PORTA IMPRESSORA TICKET" '                           ;LPT1:"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "NOME IMPRESSORA CHEQUE" '                            ;HP DeskJet 720C Series"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "PORTA IMPRESSORA CHEQUE" '                           ;LPT1:"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "NOME IMPRESSORA BOLETO" '                            ;HP DeskJet 720C Series"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "PORTA IMPRESSORA BOLETO" '                           ;LPT1:"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "NOME IMPRESSORA CARNÊ" '                             ;HP DeskJet 720C Series"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "PORTA IMPRESSORA CARNÊ" '                            ;LPT1:"
          Call UpdateArqConfig("ConfigLPT", sKeyName, sKeyValue)
        Case "VENDA RÁPIDA - USANDO SCANNER" '                     ;True"
          Call UpdateArqConfig("ConfigVR", "Scanner", sKeyValue)
        Case "VENDA RÁPIDA - ETIQUETA BALANÇA" '                   ;False"
          Call UpdateArqConfig("ConfigVR", "Etiqueta Balanca", sKeyValue)
        Case "VENDA RÁPIDA - MANTER VENDEDOR" '                    ;True"
          Call UpdateArqConfig("ConfigVR", "Mantem Vendedor", sKeyValue)
        Case "SAÍDAS - USANDO SCANNER" '                           ;True"
          Call UpdateArqConfig("ConfigSAIDAS", "Scanner", sKeyValue)
        Case "SAÍDAS - MANTÉM OPERAÇÃO" '                          ;False"
          Call UpdateArqConfig("ConfigSAIDAS", "Mantem Operacao", sKeyValue)
        Case "SAÍDAS - MANTÉM DIGITADOR" '                         ;False"
          Call UpdateArqConfig("ConfigSAIDAS", "Mantem Digitador", sKeyValue)
        Case "SAÍDAS - MANTÉM CLIENTE" '                           ;False"
          Call UpdateArqConfig("ConfigSAIDAS", "Mantem Cliente", sKeyValue)
        Case "SAÍDAS - MANTÉM TABELA DE PREÇOS" '                  ;True"
          Call UpdateArqConfig("ConfigSAIDAS", "Mantem TabPrecos", sKeyValue)
      
      End Select
    
    End If
  
  Loop
  
  Close #nFileNum
  Kill gsConfigFileName
  
End Sub

Public Function WriteOurDBVersion()
  Dim rs As Recordset
  On Error Resume Next
  Set rs = db.OpenRecordset("ZZZ", dbOpenDynaset)
  With rs
    .Edit
    .Fields("DBVersion") = App.Major & "." & App.Minor & "." & App.Revision
    .Update
  End With
  rs.Close
  Set rs = Nothing
  On Error GoTo 0
End Function

Public Function gbLoadToolID() As Boolean
  Dim rs As Recordset
  Dim rsZZZProgramas As Recordset
  
  On Error GoTo ErrLoad
  
  If gnOpenTempDB(gsTempDBFileName, False) <> 0 Then
    gbLoadToolID = False
    Exit Function
  End If
  
  Set rs = dbTemp.OpenRecordset("SELECT * FROM ActiveBar", dbOpenDynaset)
  Set rsZZZProgramas = db.OpenRecordset("SELECT * FROM ZZZProgramas", dbOpenDynaset)
  
  Do While Not rs.EOF
    With rsZZZProgramas
      .FindFirst "Número = " & rs("Numero")
      If Not .NoMatch Then
        .Edit
        .Fields("ToolID") = rs("ToolID")
        .Update
      End If
    End With
    rs.MoveNext
  Loop
  
  rs.Close
  rsZZZProgramas.Close
  Set rs = Nothing
  Set rsZZZProgramas = Nothing
  gbLoadToolID = True
  Exit Function
  
ErrLoad:
  On Error Resume Next
  rs.Close
  rsZZZProgramas.Close
  Set rs = Nothing
  Set rsZZZProgramas = Nothing
  gbLoadToolID = False
  Exit Function
  
End Function

Public Function gbLoadCodigosAcesso() As Boolean
  Dim rs As Recordset
  Dim rsProg As Recordset
  Dim sProg As String
  Dim sProgAnt As String
  Dim sCriteria As String
  
  On Error GoTo ErrLoad
  
  Call ws.BeginTrans
  Set rs = db.OpenRecordset("SELECT * FROM Acessos ORDER BY Programa", dbOpenDynaset)
  Set rsProg = db.OpenRecordset("SELECT * FROM ZZZProgramas ORDER BY [Nome Programa]", dbOpenDynaset)
  
  sProgAnt = ""
  Screen.MousePointer = vbHourglass
  Do While Not rs.EOF
    DoEvents
    sProg = rs("Programa").Value
    If sProg <> sProgAnt Then
      sProgAnt = sProg
      sCriteria = "[Nome Programa] = '" & sProg & "'"
      rsProg.FindFirst sCriteria
      If Not rsProg.NoMatch Then
        rs.Edit
        rs("Numero") = rsProg("Número").Value
        rs.Update
      End If
    Else
      rs.Edit
      rs("Numero") = rsProg("Número").Value
      rs.Update
    End If
    rs.MoveNext
  Loop
  rs.Close
  rsProg.Close
  Set rs = Nothing
  Set rsProg = Nothing
  Call ws.CommitTrans
  Screen.MousePointer = vbDefault
  gbLoadCodigosAcesso = True
  Exit Function
  
ErrLoad:
  Screen.MousePointer = vbDefault
  Call ws.Rollback
  gbLoadCodigosAcesso = False
  Exit Function
End Function

'Public Function gbCreateTableAliquotas() As Boolean
'  Dim td As TableDef
'  Dim fd As Field
'  Dim ix As Index
'  Dim nI As Integer
'
'  On Error GoTo ErrCreate
'
'  Set td = db.CreateTableDef("FISAliquotas")
'
'  Set fd = td.CreateField("Código", dbByte)
'  td.Fields.Append fd
'  Set fd = td.CreateField("Qtde", dbInteger)
'  td.Fields.Append fd
'  For nI = 1 To 16
'    Set fd = td.CreateField("Aliq" & CStr(nI), dbSingle)
'    td.Fields.Append fd
'    Set fd = td.CreateField("ISS" & CStr(nI), dbBoolean)
'    td.Fields.Append fd
'  Next nI
'
'  Set ix = td.CreateIndex("Código")
'  ix.Fields.Append ix.CreateField("Código")
'  ix.Primary = True
'  ix.Unique = True
'  td.Indexes.Append ix
'
'  db.TableDefs.Append td
'
'  Set ix = Nothing
'  Set td = Nothing
'  Set td = Nothing
'
'  gbCreateTableAliquotas = True
'  Exit Function
'
'ErrCreate:
'  gbCreateTableAliquotas = False
'
'End Function

Private Function gbNewStringDB() As Boolean
  Dim nRet As Integer
  Dim nMajor As Integer
  Dim nMinor As Integer
  Dim nRevision As Integer
  Dim bToChange As Boolean
  Dim sVersion As String
  Dim nP1 As Integer
  Dim nP2 As Integer
  Dim rsZZZ As Recordset
  
  On Error GoTo ErrHandle
    
  'Existe a versão antiga do .MDB?
  If Dir(gsOldDBFileName) <> "" Then
    gsTitle = LoadResString(201)
    gsMsg = "Encontrado uma base de dados que pode ser convertida: " & gsOldDBFileName
    gsMsg = gsMsg & vbCrLf & "Se a conversão ocorrer, a base atual QuickStore.MDB será trocada pela base antiga convertida."
    gsMsg = gsMsg & vbCrLf & "No entanto, essa conversão necessitará de acesso exclusivo a base de dados."
    gsMsg = gsMsg & vbCrLf & "Caso não queira realizar a conversão, a base atual QuickStore.MDB será utilizada."
    gsMsg = gsMsg & vbCrLf & "LEMBRE-SE: faça uma cópia das bases de dados em questão antes dessa conversão."
    gsMsg = gsMsg & vbCrLf & "Você deseja prosseguir com a conversão? "
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbYes Then
      'Sim--Abra-o em modo exclusivo e promova as alterações necessárias.
      'Elimine o MDB novo vazio e renomeie o antigo para o novo nome.
      Call gbSetNewStringDB(gsOldDBFileName)
      gbNewStringDB = True
      Exit Function
    End If
  End If
  '
  '20250327
  'Verifique se o banco de dados atual está na versão atual de alterações
  nRet = gnOpenDB(gsQuickDBFileName, False, False)
  If nRet = -1 Then
    End
  End If
  
  Set rsZZZ = db.OpenRecordset("ZZZ", dbOpenDynaset)
  
  On Error Resume Next
  bToChange = False
  sVersion = rsZZZ("DBVersion")
  If Err.Number <> 0 Then
    bToChange = True
  Else
    nP1 = InStr(sVersion, ".")
    If nP1 > 0 Then
      nMajor = Mid(sVersion, 1, nP1 - 1)
    End If
    sVersion = Mid(sVersion, nP1 + 1, Len(sVersion))
    nP2 = InStr(sVersion, ".")
    If nP2 > 0 Then
      nMinor = Mid(sVersion, 1, nP2 - 1)
      nRevision = Mid(sVersion, nP2 + 1, Len(sVersion))
    End If
    
    If nMajor < App.Major Then
      bToChange = True
    Else
      If nMinor < App.Minor Then
        bToChange = True
      Else
        If nRevision < App.Revision Then
          bToChange = True
        End If
      End If
    End If
  End If
  
  If nMajor = 6 And nMinor = 0 And (nRevision = 21 Or nRevision = 22) Then
    Call LimpaEstoque
  End If
  
  If bToChange Or gbCheckDB Then
    gsTitle = LoadResString(201)
    gsMsg = "A estrutura da Base de dados pode estar desatualizada em relação a versão atual do software."
    gsMsg = gsMsg & vbCrLf & "Uma verificação da estrutura da base irá acontecer em seguida."
    gsMsg = gsMsg & vbCrLf & "Caso a rotina necessite implementar alterações na base de dados,"
    gsMsg = gsMsg & vbCrLf & "um acesso momentâneo em modo exclusivo a base de dados será necessário."
    gnStyle = vbOKOnly + vbExclamation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    Call EraseAllConfigTB
    Call gbSetNewStringDB(gsQuickDBFileName)
    If App.Major = 6 And App.Minor >= 0 And App.Revision >= 22 Then
      If Dir(gsConfigFileName) <> "" Then
        Call GetConfigValues
      End If
    End If
  End If
  
  gbNewStringDB = True
  Exit Function
  
ErrHandle:
  gsTitle = LoadResString(201)
  gsMsg = "Erro na Rotina de Inicialização da Base de Dados: revise a instalação do produto, caminho da pasta de arquivos em rede e/ou a base de dados."
  gsMsg = gsMsg & vbCrLf & "Erro: " & Err.Number & " - " & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  gbNewStringDB = False
  Exit Function
End Function

Private Sub EraseAllConfigTB()
  If Dir(gsConfigPath & "*.TB") <> "" Then
    Kill gsConfigPath & "*.TB"
  End If
End Sub

Public Function CriptografaSenha(ByVal sz As String) As Double
  Dim Senha As Double
  Dim n As Integer
  Senha = 0
  If sz <> "" Then
     For n = 0 To Len(sz) - 1
        Senha = Senha + Asc(Mid(sz, n + 1, 1)) * (10 ^ n)
     Next n
  End If
  CriptografaSenha = Senha
End Function

Public Sub EditCopy(F As Form)
  On Error Resume Next
  Clipboard.SetText F.ActiveForm.ActiveControl.SelText
End Sub

Public Sub EditCut(F As Form)
  On Error Resume Next
  Clipboard.SetText F.ActiveForm.ActiveControl.SelText
  F.ActiveForm.ActiveControl.SelText = ""
End Sub

Public Sub EditPaste(F As Form)
  On Error Resume Next
  F.ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End Sub

Public Sub GetNewCode(ByRef F As Form, ByRef rs As Recordset, ByVal nMaxCod As Long)
  Dim rsTemp As Recordset
  Dim nCod As Long
  
  Call F.ClearScreen
  
  nCod = 1
  Set rsTemp = rs.Clone
  With rsTemp
    If Not .EOF Then
      .MoveLast
      nCod = .Fields("Código") + 1
      If nCod > nMaxCod Then
        For nCod = 1 To nMaxCod
          DoEvents
          .FindFirst "Código = " & nCod
          If .NoMatch Then
            F.Código.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
            Exit Sub
          End If
        Next nCod
        gsTitle = LoadResString(201)
        gsMsg = "Nenhum próximo Código livre disponível para o intervalo."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Else
        F.Código.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
      End If
    Else
      F.Código.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
    End If
    .Close
  End With
  
  Set rsTemp = Nothing
  
End Sub

Public Sub GetNewCode2(ByVal F As Form, ByVal ctlControl As SSDBCombo, ByRef rs As Recordset, ByVal nMaxCod As Long)
  Dim rsTemp As Recordset
  Dim nCod As Long
  
  On Error Resume Next
  
  Call F.ClearScreen
  
  nCod = 1
  Set rsTemp = rs.Clone
  With rsTemp
    If Not .EOF Then
      .MoveLast
      nCod = .Fields("Código") + 1
      If nCod > nMaxCod Then
        For nCod = 1 To nMaxCod
          DoEvents
          .FindFirst "Código = " & nCod
          If .NoMatch Then
            ctlControl.Código.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
            Exit Sub
          End If
        Next nCod
        gsTitle = LoadResString(201)
        gsMsg = "Nenhum próximo Código livre disponível para o intervalo."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Else
        ctlControl.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
      End If
    Else
      ctlControl.Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
    End If
    .Close
  End With
  
  Set rsTemp = Nothing
  
  On Error GoTo 0
  
End Sub

Public Sub DisplayMsg(ByVal sMsg As String)
  'Limpa a Barra de Status
  Call StatusMsg("")
  'Exibe a mensagem
  gsTitle = LoadResString(201)
  gsMsg = sMsg
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
End Sub

Public Function gsSupressSpecialChars(ByVal sCaption As String) As String
  Dim nI As Integer
  Dim sCh As String
  Dim sText As String
  sText = ""
  For nI = 1 To Len(sCaption)
    sCh = Mid(sCaption, nI, 1)
    If InStr(".&", sCh) = 0 Then
      sText = sText & UCase(sCh)
    End If
  Next nI
  gsSupressSpecialChars = sText
End Function

Public Function bGridBeforeDelete() As Boolean
  gsTitle = LoadResString(201)
  gsMsg = LoadResString(253)
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    bGridBeforeDelete = False
  Else
    bGridBeforeDelete = True
  End If
End Function

Public Function IsGoodNumber(ByRef txtValor As TextBox) As Boolean
  Dim nPos As Integer
'  Dim sInt As String
'  Dim sFrac As String
  IsGoodNumber = True
  txtValor.Text = gsFormatCurrency(txtValor.Text, gnCurrencyDecimals)
  nPos = InStr(txtValor.Text, gsCurrencyDecimal)
  If nPos > 0 Then
'    sInt = CStr(CDbl((gsHandleNull(Left(txtValor.Text, nPos - 1)))))
'    sFrac = CStr(CDbl((gsHandleNull(Mid(txtValor.Text, nPos + 1, Len(txtValor.Text) - nPos)))))
    nPos = InStr(Mid(txtValor.Text, nPos + 1, Len(txtValor.Text)), gsCurrencyDecimal)
    If nPos > 0 Then
      IsGoodNumber = False
      Exit Function
    End If
'    txtValor.Text = Format((CDbl(sInt) + CDbl(sFrac) / 10 ^ (Len(sFrac))), "#########0.00")
  End If
  If Not IsGoodNumber Then
    DisplayMsg "Valor inválido. Verifique."
    txtValor.SetFocus
  End If
End Function

Public Function IsVarGoodNumber(ByVal sVar As String) As Boolean
  Dim nLen As Integer
  Dim nI As Integer
  Dim sCh As String
  Dim nCtDec As Integer
  nLen = Len(Trim(sVar))
  nCtDec = 0
  For nI = 1 To nLen
    sCh = Mid(sVar, nI, 1)
    If InStr("+-,.0123456789", sCh) = 0 Then
      IsVarGoodNumber = False
      Exit Function
    Else
      If gsCurrencyDecimal = sCh Then
        nCtDec = nCtDec + 1
      End If
    End If
  Next nI
  If nCtDec > 1 Then
    IsVarGoodNumber = False
  Else
    IsVarGoodNumber = True
  End If
End Function

Public Sub UpdateArqConfig(ByVal sSession As String, ByVal sKeyName As String, ByVal vVar As Variant)
  Dim sValor As String
  If VarType(vVar) = vbBoolean Then
    If vVar = True Then
      sValor = "True"
    Else
      sValor = "False"
    End If
  Else
    sValor = CStr(vVar)
  End If
  Call SaveSetting("QuickStore", sSession, sKeyName, sValor)
End Sub

Public Function gsGetDateField(ByVal dtDate As Date) As String
  If IsDate(dtDate) Then
    gsGetDateField = CDate(dtDate)
  Else
    gsGetDateField = ""
  End If
End Function

Public Function gsOpenFile(ByRef F As Form, ByVal sTitle As String, ByVal sFilter As String) As String
  On Error Resume Next
  With F.cdgFileOpen
    .DialogTitle = sTitle
    .CancelError = True
    .InitDir = App.Path & "\Config"
    .Filter = sFilter
    .Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly + cdlOFNLongNames
    .Filename = ""
    .ShowOpen
  End With
  If Err.Number = 0 Then
    gsOpenFile = F.cdgFileOpen.Filename
  Else
    gsOpenFile = ""
  End If
  On Error GoTo 0
  
End Function

Public Sub GetNumberOfUsers()
  Dim nI As Integer
'  gsCurrentUsers = gsGetMDBUsers(gsQuickTMPFileName)
  gsCurrentUsers = gsGetMDBUsers(gsQuickDBFileName)
  gnCtCurrentUsers = 0
  For nI = LBound(gsCurrentUsers) To UBound(gsCurrentUsers)
    DoEvents
    If Len(gsCurrentUsers(nI)) = 0 Then
      Exit For
    End If
    gnCtCurrentUsers = gnCtCurrentUsers + 1
  Next nI
End Sub

Public Function gnAtualizaSaldoBancario(ByVal nContaCorrente As Integer) As Long
  Dim nCount As Long
  Dim nSaldo As Double
  Dim sSql As String
  Dim rsLançamentos As Recordset
  
  nCount = 0
  nSaldo = 0
  
  sSql = "SELECT * FROM [Lançamentos Bancários] "
  sSql = sSql & "WHERE Conta = " & nContaCorrente & " ORDER BY Data, Ordem"
  Set rsLançamentos = db.OpenRecordset(sSql, dbOpenDynaset)

  Call ws.BeginTrans
    
  Do While Not rsLançamentos.EOF
    rsLançamentos.Edit
    rsLançamentos("Saldo Anterior") = nSaldo
    nSaldo = nSaldo + rsLançamentos("Crédito") - rsLançamentos("Débito")
    rsLançamentos("Saldo Atual") = Format(nSaldo, "############0.00")
    rsLançamentos.Update
    Call StatusMsg("Atualizando dia " & Format(rsLançamentos("Data").Value, "dd/mm/yyyy"))
    nCount = nCount + 1
    rsLançamentos.MoveNext
  Loop

  Call ws.CommitTrans

  rsLançamentos.Close
  Set rsLançamentos = Nothing
  
  gnAtualizaSaldoBancario = nCount
  
End Function

Public Function GetCodigoCombos(strCodigo As String) As Integer
  Dim intCont         As Integer
  Dim strTemp         As String
  Dim strAcumulativa  As String
  
  For intCont = 1 To Len(strCodigo)
    strTemp = Mid(strCodigo, intCont, 1)
    
    If strTemp <> "-" Then
      strAcumulativa = strAcumulativa & strTemp
    Else
      GetCodigoCombos = CInt(Trim(strAcumulativa))
      Exit Function
    End If
  Next intCont
End Function

Public Sub g_GravaLog(ByVal datData As Date, ByVal strTexto As String, ByVal strTipo As String)
  '---[ Gera Log do usuário ]---'
    Dim rstZZZLog As Recordset
    Dim blnInTransaction As Boolean
    
    On Error GoTo Erro
    
    ws.BeginTrans
    blnInTransaction = True
    
    Set rstZZZLog = db.OpenRecordset("SELECT * FROM ZZZLog", dbOpenDynaset)
    With rstZZZLog
      .AddNew
      .Fields("Data") = datData
      .Fields("Texto") = Left(strTexto, 80)
      .Fields("Tipo") = Left(strTipo, 20)
      .Update
      .Close
    End With
    Set rstZZZLog = Nothing
    
    ws.CommitTrans
    blnInTransaction = False
  '---[ Gera Log do usuário ]---'
  
  Exit Sub
  
Erro:
  If blnInTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

'13/07/2007 - Anderson
'Função criada para verificar os acessos permitidos no programa
'Retorna uma string com 0 ou 1, sendo:
'Primeira posição = Se a função está cadastrada no usuário, permitindo VISUALIZAR.
'Segunda Posição  = Se o programa permite GRAVAR
'Terceira Posição = Se o programa permite APAGAR
Public Function gbGetUserPermition(ByVal CodigoUsuario As Integer, ByVal CodigoPrograma As Integer) As String

  Dim rsAcessos As Recordset
 
  gbGetUserPermition = "000"
  
  Set rsAcessos = db.OpenRecordset("SELECT * FROM Acessos WHERE Usuário = " & CStr(CodigoUsuario) & " AND Numero = " & CStr(CodigoPrograma), dbOpenDynaset)

  If Not rsAcessos.EOF Then
    
    gbGetUserPermition = "1"
    
    If rsAcessos("Gravar") = -1 Then
      gbGetUserPermition = gbGetUserPermition & "1"
    Else
      gbGetUserPermition = gbGetUserPermition & "0"
    End If
    
    If rsAcessos("Apagar") = -1 Then
      gbGetUserPermition = gbGetUserPermition & "1"
    Else
      gbGetUserPermition = gbGetUserPermition & "0"
    End If
    
  End If
  
  rsAcessos.Close
  Set rsAcessos = Nothing
  
End Function

'19/07/2007 - Anderson
'Função criada para verificar os acessos permitidos as funções do usuário
'Retorna verdadeiro ou falso, de acordo com o que está registrado
Public Function gbGetUserFunction(ByVal CodigoUsuario As Integer, ByVal CampoTabelaFuncionario As String) As Boolean

  Dim rsFuncionarios As Recordset
  
  Set rsFuncionarios = db.OpenRecordset("SELECT * FROM Funcionários WHERE Código = " & CStr(CodigoUsuario), dbOpenDynaset)

  If Not rsFuncionarios.EOF Then
    
    gbGetUserFunction = rsFuncionarios(CampoTabelaFuncionario)
    
  End If
  
  rsFuncionarios.Close
  Set rsFuncionarios = Nothing
  
End Function

Private Sub LimpaEstoque()
  Dim rs As Recordset
  Set rs = db.OpenRecordset("Estoque", dbOpenDynaset)
  If Not rs.EOF Then
    rs.FindFirst "Filial = 1 And Produto = '1' And Data = #01/18/2000#"
    If Not rs.NoMatch Then
      rs.Delete
    End If
  End If
  rs.Close
  Set rs = db.OpenRecordset("Estoque Final", dbOpenDynaset)
  If Not rs.EOF Then
    rs.FindFirst "Filial = 1 And Produto = '1' And [Estoque Atual] = -1 And [Última Data] = #01/18/2000#"
    If Not rs.NoMatch Then
      rs.Delete
    End If
  End If
  rs.Close
  Set rs = Nothing
End Sub

Public Sub ConsultaRpt(sqlRdp As String)
  If Not (connRpt Is Nothing) Then connRpt.Close
  Set connRpt = Nothing
  Set connRpt = New ADODB.Connection

  connRpt.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;User ID=Admin;Data Source=" & App.Path & "\QuickStore.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password=' & gsGetPValue() & ';Jet OLEDB:Global Partial Bulk Ops=2"

  'If Not (rsRdp Is Nothing) Then rsRdp.Close
  Set rsRdp = Nothing
  Set rsRdp = New ADODB.Recordset
  rsRdp.CursorLocation = adUseClient
  rsRdp.Open sqlRdp, connRpt, adOpenForwardOnly, adLockReadOnly
End Sub

Public Sub EncerraRpt()
  If Not (rsRdp Is Nothing) Then rsRdp.Close
  Set rsRdp = Nothing
  If Not (connRpt Is Nothing) Then connRpt.Close
  Set connRpt = Nothing
End Sub
