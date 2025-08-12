VERSION 5.00
Begin VB.Form frmAbout2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFA324&
   BorderStyle     =   0  'None
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   Icon            =   "frmAbout2.frx":0000
   LinkTopic       =   "Form1"
   Palette         =   "frmAbout2.frx":4E95A
   Picture         =   "frmAbout2.frx":87DD8
   ScaleHeight     =   6795
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   7245
      Top             =   4995
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   360
      Picture         =   "frmAbout2.frx":91339
      ScaleHeight     =   4395
      ScaleWidth      =   7905
      TabIndex        =   0
      Top             =   2070
      Visible         =   0   'False
      Width           =   7935
      Begin VB.Label lblFileDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "file description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6930
         TabIndex        =   8
         Top             =   3960
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "application title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6030
         TabIndex        =   7
         Top             =   4005
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblUserInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "user information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6210
         TabIndex        =   6
         Top             =   3735
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblUserName 
         BackStyle       =   0  'Transparent
         Caption         =   "user name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   4725
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "warning message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6435
         TabIndex        =   4
         Top             =   4230
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPathEXE 
         BackStyle       =   0  'Transparent
         Caption         =   "path and exe information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4500
         TabIndex        =   3
         Top             =   3600
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "copyright information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   2
         Top             =   5085
         Visible         =   0   'False
         Width           =   4530
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "version information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   3870
         TabIndex        =   1
         Top             =   3195
         Visible         =   0   'False
         Width           =   2715
      End
   End
End
Attribute VB_Name = "frmAbout2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'28/01/2009 - mpdea
'Adaptado para versão 7

' API declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
() '        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
() '        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
() '        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
() '        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

' API Constants
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CAPTION_NOT As Long = &HFFFFFFFF - WS_CAPTION

Private Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGKEYSYSINFO As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"

Private Const gREGVALSYSINFOLOC As String = "MSINFO"
Private Const gREGVALSYSINFO As String = "PATH"

' NT location of user name and company
Private Const gNTREGKEYINFO As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Private Const gNTREGVALUSER As String = "RegisteredOwner"
Private Const gNTREGVALCOMPANY As String = "RegisteredOrganization"

' Win95 locataion of user name and company
Private Const g95REGKEYINFO As String = "Software\Microsoft\MS Setup (ACME)\User Info"
Private Const g95REGVALUSER As String = "DefName"
Private Const g95REGVALCOMPANY As String = "DefCompany"

' Change these to what you want the default name and user info to be
Private Const DEFAULT_USER_NAME As String = vbNullString
Private Const DEFAULT_USER_INFO As String = vbNullString

' Information for warning information at bottom of form
Private Const gWarningInfo As String = ""

Public bMyProgramIsRegistered As Boolean

Private mStyle As StyleType
Private mTitleBarHidden As Boolean

' Type declarations
Private Type StyleType
    OldStyle As Long
    NewStyle As Long
End Type 'StyleType

Public sRETORNO_WEBAPI_LICENCA As String

Private Sub Form_Load()

  Screen.MousePointer = vbHourglass
  
  bSoapClient_MSSoapInit = False

  Call CenterForm(Me)

  lblWarning.Caption = LoadResString(14)

' Fill in all of the information that comes from the App object
  With App
    Caption = "Infopar " & .ProductName
    lblTitle.Caption = .ProductName

    lblVersion.Caption = "Versão " & .Major & "." & .Minor & "." & .Revision
    sVersaoDoSistema = lblVersion.Caption
    lblCopyright.Caption = .LegalCopyright
    ' lblTrademark.Caption = .LegalTrademarks
    lblPathEXE.Caption = .Path & "\" & .EXEName & ".exe"
    lblFileDescription.Caption = LoadResString(1)
  End With 'App
  
  ' Pilatti Abril/2108
  leArquivoERP_APP_QUICKTORE

  'lblMisc.Caption = LoadResString(104)

  Screen.MousePointer = vbDefault

End Sub

'Public Sub About(frmParent As Form, Optional lUserName As String, _
'                 Optional lUserInfo As String)
'    ' imgIcon.Picture = frmParent.Icon
'    cmdOK.Enabled = True
'    cmdSysInfo.Enabled = True
'
''   Add user information to form
'    If lUserName <> "" Then
'        lblUserName.Caption = lUserName
'        lblUserInfo.Caption = lUserInfo
'    Else
'        lblUserName.Caption = LoadResString(60)
'        lblUserInfo.Caption = ""
'    End If
'
'    Show vbModal
'End Sub

'Public Sub SplashOn(frmParent As Form, Optional MinDisplay As Long, _
'                    Optional lUserName As String, Optional lUserInfo As String)
Public Sub SplashOn(Optional MinDisplay As Long, _
                    Optional lUserName As String, Optional lUserInfo As String)
'''On Error GoTo ERRO_Geral
''''  ' ***************************************
''''  ' Tratamento de ACESSO PERIODO VIA WEBAPI
''''  ' PILATTI DEZ/17
'''  Dim dtCorrente As String
'''  Dim dtValida As Date
'''  Dim rsValida As Recordset
'''  Dim v1 As String
'''  Dim v2 As String
'''  Dim vD As String
'''  Dim dtCorrenteBlindada As String
'''  Dim dtCorrenteDesblindada As String
'''  Dim sSQLLicenca As String
'''
'''
''''  Dim sCGCCPFLicenca As String
''''  Dim rsParamFilialLicenca As Recordset
''''  Set rsParamFilialLicenca = db.OpenRecordset("Select CGC FROM [Parâmetros Filial] where Filial=1 ", dbOpenDynaset)
''''  sCGCCPFLicenca = rsParamFilialLicenca.Fields("CGC").Value
''''  rsParamFilialLicenca.Close
''''  Set rsParamFilialLicenca = Nothing
'''
'''  ' ********************************************
'''  ' FLAG GERAL PARA VERIFICAR SE ESTA VERSÃO DO QUICKSTORE TRATA OU NÃO O CONTROLE DE LICENÇA DE USO DO QUICK VIA WEBAPI
'''  ' 1-TRATA;  0-NÃO TRATA
'''  gTrataControleDeLicencaWebApi = 1
'''
'''  ' ********************************************
'''  ' FLAG GERAL SE USUARIO TEM OU NÃO ACESSO AO 'ESTRATÉGICO RELATORIOS'... 1-Tem acesso;  0-Não tem;
'''  gESTRATEGICO_Relatorios = 0
'''
'''  ' FLAG GERAL PARA ABRIR OU NÃO AS TELAS DO XML... 1-ABRE; 0-NÃO ABRE;
'''  gAbreModuloXML = 0
'''  ' ********************************************
'''
'''
'''  ' Abrir arquivo .txt para recuperar endereço alvo do Componente WebApi (Homolog ou Produção)
'''  Dim ff As Integer
'''  Dim sAux As String
'''  Dim sEndereco As String
'''  Dim sCODIGO_ERRO As String
'''  Dim sTesteApenasSourceSafe As String
'''  Dim sNrSerieNFTratado As String
'''  Dim iNrSerieNFTratado As Integer
'''  Dim xConta As Integer
'''  Dim sLendoLinha As String
'''
'''  ' Controle de mensagem de erro...
'''  sCODIGO_ERRO = "Acesso ao arquivo de endereço"
'''
'''  ff = FreeFile
'''  Open App.Path + "\ERP_WEBAPI_Acesso_Config.txt" For Input As #ff
'''    Line Input #ff, sEndereco
'''    Line Input #ff, sSoapClient_MSSoapInit
'''    Line Input #ff, sSoapClient_ConnectorProperty_EndPointURL
'''    Line Input #ff, sCaminhoDanfe_Benefix
'''    Line Input #ff, sNrSerieNFTratado           'Numero de CNPJ tratados neste arquivo (para o campo 'Serie' da NF)
'''
'''    iNrSerieNFTratado = CInt(sNrSerieNFTratado)
'''    gNrSerieNF = iNrSerieNFTratado
'''    If iNrSerieNFTratado > 0 Then
'''        For xConta = 0 To gNrSerieNF - 1
'''            'NrCnpj1 , SerieNFe1, SerieNFCe1
'''            'NrCnpj2 , SerieNFe2, SerieNFCe2
'''
'''            Line Input #ff, sLendoLinha         'CNPJ
'''            gArrayNrSerieNF(xConta, 0) = sLendoLinha
'''            Line Input #ff, sLendoLinha         'NrSerieNFe
'''            gArrayNrSerieNF(xConta, 1) = sLendoLinha
'''            Line Input #ff, sLendoLinha         'NrSerieNFCe
'''            gArrayNrSerieNF(xConta, 2) = sLendoLinha
'''        Next
'''    End If
'''  Close #ff
'''
'''  sCODIGO_ERRO = "Após leitura do arq. de endereço"
'''
'''  If gTrataControleDeLicencaWebApi = 1 Then
'''
'''      Dim bLICENCA As Boolean
'''      bLICENCA = True
'''
'''      dtCorrente = Format(Now, "yyyy/MM/dd")
'''      dtValida = Format(dtCorrente, "yyyy/MM/dd") 'define uma nova data
'''
'''      sCODIGO_ERRO = "Acesso ao DB Busca Chave"
'''
'''      'Date = mdate  ´altera a data do sistema
'''      'Set rsValida = db.OpenRecordset("SELECT dadoValida, dadoConteudo, dadoV2, managerNFEXML, estrategicoREL FROM ParametroVLC", dbOpenDynaset)
'''      Set rsValida = db.OpenRecordset("SELECT * FROM ParametroVLC", dbOpenDynaset)
'''
'''      sCODIGO_ERRO = "Após o Acesso ao DB Busca Chave"
'''
'''      If Not rsValida.EOF Then
'''        v1 = rsValida.Fields("dadoValida").Value
'''
'''        If IsNull(rsValida.Fields("dadoConteudo").Value) Then
'''          vD = ""
'''        Else
'''          vD = rsValida.Fields("dadoConteudo").Value
'''        End If
'''
'''        v2 = rsValida.Fields("dadoV2").Value
'''
'''        ' FLAG SE USUARIO TEM OU NÃO ACESSO AO 'ESTRATÉGICO RELATORIOS'... 1-Tem acesso;  0-Não tem;
'''        If rsValida.Fields("estrategicoREL").Value <> "1" Then
'''            gESTRATEGICO_Relatorios = 0   'Não tem acesso
'''        Else
'''            gESTRATEGICO_Relatorios = 1   'Tem acesso
'''        End If
'''
'''        ' FLAG  ABRIR OU NÃO AS TELAS DO XML... 1-ABRE; 0-NÃO ABRE;
'''        If rsValida.Fields("managerNFEXML").Value <> "1" Then
'''            gAbreModuloXML = 0   'Não tem acesso
'''        Else
'''            gAbreModuloXML = 1   'Tem acesso
'''        End If
'''
'''        ' Regra geral:
'''        ' Tabela ParametroVLC
'''        ' Campos dadoValida = 'X2' e dadoConteudo = 'AF'   ... Cliente deve ser controlado mensalmente o seu ACESSO via WEBAPI A3
'''        ' Campos dadoValida = 'AB' e dadoConteudo = 'X3'   ... Cliente que não precisa de verificação de acesso
'''        ' Se não existe registro algum nesta tabela (como uma das condições acima)    ...Acesso bloqueado
'''
'''        ' Condição
'''        If v1 = "X2" And v2 = "AF" Then
'''          If vD = "" Then
'''            ' Primeiro acesso do usuario...sem data de vencimento...então gravo neste acesso data de (hoje -90 dias) ou seja, já expirada
'''            ' Gravar com data de hoje
'''            'dtCorrenteBlindada = dtCorrente
'''            dtCorrenteBlindada = Format(Now - 90, "yyyy/MM/dd")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "1", "A")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "2", "M")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "3", "C")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "4", "P")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "5", "E")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "6", "H")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "7", "I")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "8", "K")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "9", "L")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "0", "S")
'''            dtCorrenteBlindada = Replace(dtCorrenteBlindada, "/", "R")
'''            dtCorrenteBlindada = "XABJEDEG" + dtCorrenteBlindada + "FIPBWEDA"
'''
'''            Call db.Execute("UPDATE ParametroVLC SET dadoConteudo= " & _
'''            " '" & dtCorrenteBlindada & "' WHERE dadoValida='X2' and dadoV2='AF' ", dbFailOnError)
'''
'''            ' Encerrar a aplicação para que o usuário agora possa chamar o WebApi...
'''            rsValida.Close
'''            Set rsValida = Nothing
'''            MsgBox "A configuração de licença de uso esta sendo validada. Por favor, faça o acesso novamente! (Módulo: Controle de Acesso Cód. Branco B1.)", , "Acesso"
'''            End
'''
'''
'''          Else
'''            ' Existe data para controle de acesso
'''            If Len(Trim(vD)) <> 26 Then
'''                ' Se a chave com a data blindada não tiver tamanho de 26 caracteres...bloquear acesso
'''                sCODIGO_ERRO = "Chave blindada mal formatada"
'''                bLICENCA = False
'''                ' Encerrar a aplicação !
'''            ElseIf Trim(sEndereco) = "" Then
'''                ' URL Inexistente ao WebApi...bloquear acesso
'''                sCODIGO_ERRO = "URL Inexistente ao WebApi"
'''                bLICENCA = False
'''                ' Encerrar a aplicação !
'''            Else
'''                dtCorrenteDesblindada = vD
'''                dtCorrenteDesblindada = Mid(dtCorrenteDesblindada, 9, 10)
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "A", "1")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "M", "2")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "C", "3")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "P", "4")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "E", "5")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "H", "6")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "I", "7")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "K", "8")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "L", "9")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "S", "0")
'''                dtCorrenteDesblindada = Replace(dtCorrenteDesblindada, "R", "/")
'''
'''                ' Validar se passou dos  30 dias!
'''                Dim dtBase As Date
'''                dtBase = Format(dtCorrenteDesblindada, "yyyy/MM/dd")
'''
'''                dtValida = dtValida - Day(30)
'''
'''                If dtValida > dtBase Then
'''                  ' Chamar webApi para verificar se este CNPJ esta liberado para acesso/uso
'''                  Dim s As String
'''                  s = "Chamar"
'''
'''                  sAux = lUserInfo
'''                  sAux = Replace(sAux, "/", "")
'''                  sAux = Replace(sAux, ".", "")
'''                  sAux = Replace(sAux, "-", "")
'''                  sAux = Replace(sAux, ",", "")
'''                  sAux = Replace(sAux, "\", "")
'''                  sAux = Replace(sAux, ";", "")
'''                  sAux = Replace(sAux, " ", "")
'''
'''                  If sAux = "" Then
'''                      Dim sCGCCPFLicenca As String
'''                      Dim rsParamFilialLicenca As Recordset
'''                      Set rsParamFilialLicenca = db.OpenRecordset("Select CGC FROM [Parâmetros Filial] where Filial=1 ", dbOpenDynaset)
'''                      sAux = rsParamFilialLicenca.Fields("CGC").Value
'''                      rsParamFilialLicenca.Close
'''                      Set rsParamFilialLicenca = Nothing
'''                      sAux = Replace(sAux, "/", "")
'''                      sAux = Replace(sAux, ".", "")
'''                      sAux = Replace(sAux, "-", "")
'''                      sAux = Replace(sAux, ",", "")
'''                      sAux = Replace(sAux, "\", "")
'''                      sAux = Replace(sAux, ";", "")
'''                      sAux = Replace(sAux, " ", "")
'''                  End If
'''
'''                  WebRequest sEndereco + "?_cnpj=" + sAux
'''
'''                  ' Ex dos dados que retornar do WebApi dentro da variavel sRETORNO_WEBAPI_LICENCA
'''                  ' LICENCA_ATIVA__XML1_EstratRel1
'''                  ' LICENCA_ATIVA__XML0_EstratRel0
'''                  ' LICENCA_ATIVA__XML0_EstratRel1
'''                  ' LICENCA_ATIVA__XML1_EstratRel0
'''
'''                  If InStr(1, sRETORNO_WEBAPI_LICENCA, "LICENCA_ATIVA") > 0 Then
'''                    ' Se esta liberado (pegar o retorno do webApi), então atualiza a base com data de hoje
'''                    ' Gravar com data de hoje
'''                    dtCorrenteBlindada = dtCorrente
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "1", "A")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "2", "M")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "3", "C")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "4", "P")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "5", "E")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "6", "H")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "7", "I")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "8", "K")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "9", "L")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "0", "S")
'''                    dtCorrenteBlindada = Replace(dtCorrenteBlindada, "/", "R")
'''                    dtCorrenteBlindada = "XABJEDEG" + dtCorrenteBlindada + "FIPBWEDA"
'''
'''                    sSQLLicenca = "UPDATE ParametroVLC SET dadoConteudo='" & dtCorrenteBlindada & "', "
'''                    If InStr(1, sRETORNO_WEBAPI_LICENCA, "XML1") > 0 Then
'''                      sSQLLicenca = sSQLLicenca & " estrategicoREL='1', "
'''                    Else
'''                      sSQLLicenca = sSQLLicenca & " estrategicoREL='0', "
'''                    End If
'''
'''                    If InStr(1, sRETORNO_WEBAPI_LICENCA, "EstratRel1") > 0 Then
'''                      sSQLLicenca = sSQLLicenca & " managerNFEXML='1' "
'''                    Else
'''                      sSQLLicenca = sSQLLicenca & " managerNFEXML='0' "
'''                    End If
'''
'''                    sSQLLicenca = sSQLLicenca & " WHERE dadoValida='X2' and dadoV2='AF' "
'''
'''                    Call db.Execute(sSQLLicenca, dbFailOnError)
'''
'''                  ElseIf InStr(1, sRETORNO_WEBAPI_LICENCA, "LICENCA_INATIVA") > 0 Then
'''                    ' Caso não esta liberado...Encerrar a aplicação !!
'''                    sCODIGO_ERRO = "Ret. LIÇENCA INATIVA"
'''                    MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E1. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
'''                    bLICENCA = False
'''                    ' Encerrar a aplicação !!
'''                  Else
'''                    sCODIGO_ERRO = "Ret. LIÇENCA ??"
'''                    ' Erro na chamada ao WebApi...Encerrar a aplicação !!
'''                    MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E2. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
'''                    bLICENCA = False
'''                    ' Encerrar a aplicação !!
'''                  End If
'''
'''                End If
'''
'''            End If
'''
'''          End If
'''        ElseIf v1 = "AB" And v2 = "X3" Then
'''          ' Cliente que não precisa de verificação de acesso liberado
'''          Dim ss As String
'''          ss = "ok"
'''        Else
'''          ' Inconsistente
'''          sCODIGO_ERRO = "Registro com Minichaves inconsistentes"
'''          MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E3. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
'''          bLICENCA = False
'''          ' Encerrar a aplicação !!
'''        End If
'''
'''      Else
'''        ' Acesso bloqueado !!
'''        sCODIGO_ERRO = "Nenhum Registro com Minichaves"
'''        MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E4. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
'''        bLICENCA = False
'''        ' Encerrar a aplicação !!
'''      End If
'''
'''      ' ***************************************
'''
'''      If bLICENCA = False Then
'''        rsValida.Close
'''        Set rsValida = Nothing
'''        MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E5. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
'''        End
'''      End If
'''  End If


  If Not Visible Then
      Dim lHeight As Integer

'       Add user information to form
      If lUserName <> "" Then
          lblUserName.Caption = lUserName
          lblUserInfo.Caption = lUserInfo
      Else
          lblUserName.Caption = GetUserName
          lblUserInfo.Caption = GetUserCompany
      End If
      
      'Height = linDivide(1).Y1 + 15 '(Height - ScaleHeight)

      Show vbModeless

'       For some reason, need a Refresh to make sure Splash Screen gets painted
      Refresh
  End If
  
'''  rsValida.Close
'''  Set rsValida = Nothing
'''
'''  Exit Sub
'''
'''ERRO_Geral:
'''    ' Acesso bloqueado !!
'''    MsgBox "Acesso bloqueado à Aplicação! (Módulo: Controle de Acesso Cód. Vermelho E6. " + sCODIGO_ERRO + ". Por favor, entre em contato conosco.)", , "Acesso"
'''    'MsgBox "Erro leitura do arquivo WS", , "Acesso"
'''    End

End Sub

Sub WebRequest(url As String)
On Error GoTo trata_WebApiErro

    ' Função que chama o componente WebApi
    Dim retWebApiAux As String

    Dim http As MSXML2.XMLHTTP
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    http.Open "GET", url, False
    http.Send

    retWebApiAux = http.statusText
    sRETORNO_WEBAPI_LICENCA = http.responseText

    Exit Sub
    
trata_WebApiErro:
'    Dim iInStrRet As Integer
'    iInStrRet = InStr(1, Err.Description, "timed out")
'    If iInStrRet > 0 And tentativasWEBAPI < 4 Then
'        tentativasWEBAPI = tentativasWEBAPI + 1
'        WebRequest url
'    End If
'    sRETORNO_ERRO_WEBAPI = Err.Description
End Sub

Public Sub SplashOff()
  If Visible Then
    'Wait until any minimum display time elapses
    If gnDeltaTime < 10 Then
      Timer1.Interval = 5000
      Timer1.Enabled = True
      Do While Timer1.Enabled
        DoEvents
      Loop
    End If
    Unload Me
  End If
End Sub



Private Sub Timer1_Timer()
  Timer1.Enabled = False
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Temporary Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...

    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size

    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format("&h" + KeyVal)                     ' Convert Double Word To String
    End Select

    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit

GetKeyError:      ' Cleanup After An Error Has Occurred...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function GetUserName() As String
    Dim KeyVal As String

'   For WindowsNT
    If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALUSER, KeyVal)) Then
        GetUserName = KeyVal
'   For Windows95
    ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALUSER, KeyVal)) Then
        GetUserName = KeyVal
'   None of the above
    Else
        GetUserName = DEFAULT_USER_NAME
    End If
End Function

Private Function GetUserCompany() As String
    Dim KeyVal As String

'   For WindowsNT
    If (GetKeyValue(HKEY_LOCAL_MACHINE, gNTREGKEYINFO, gNTREGVALCOMPANY, KeyVal)) Then
        GetUserCompany = KeyVal
'   For Windows95
    ElseIf (GetKeyValue(HKEY_CURRENT_USER, g95REGKEYINFO, g95REGVALCOMPANY, KeyVal)) Then
        GetUserCompany = KeyVal
'   None of the above
    Else
        GetUserCompany = DEFAULT_USER_INFO
    End If
End Function



