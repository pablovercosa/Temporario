Attribute VB_Name = "modWindowPos"
Option Explicit

' API declarations
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

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
Public Const GWL_STYLE As Long = (-16)
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_CAPTION_NOT As Long = &HFFFFFFFF - WS_CAPTION

Public Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
Public Const gREGKEYSYSINFO As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"

Public Const gREGVALSYSINFOLOC As String = "MSINFO"
Public Const gREGVALSYSINFO As String = "PATH"

' NT location of user name and company
Public Const gNTREGKEYINFO As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Public Const gNTREGVALUSER As String = "RegisteredOwner"
Public Const gNTREGVALCOMPANY As String = "RegisteredOrganization"

' Win95 locataion of user name and company
Public Const g95REGKEYINFO As String = "Software\Microsoft\MS Setup (ACME)\User Info"
Public Const g95REGVALUSER As String = "DefName"
Public Const g95REGVALCOMPANY As String = "DefCompany"

' Change these to what you want the default name and user info to be
Public Const DEFAULT_USER_NAME As String = "USER INFORMATION NOT AVAILABLE"
Public Const DEFAULT_USER_INFO As String = vbNullString

' Information for warning information at bottom of form
Public Const gWarningInfo As String = ""

Public bMyProgramIsRegistered As Boolean


Public mBoxHeight As Integer
Public mStyle As StyleType
Public mTitleBarHidden As Boolean

' Type declarations
Public Type StyleType
    OldStyle As Long
    NewStyle As Long
End Type 'StyleType





'Constants for topmost.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Declare Function SetWindowPos _
    Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
   
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As Any, _
      ByVal lpWindowName As Any) As Long

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal nCmdShow As Long) As Long

Private Const SW_SHOW = 5
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3


Public CTRL_OFFSET As Single
Public SPLT_COLOUR As Single

' Windows 32-bit API Declarations
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'
' 32-bit return variable declaration
Public apiRetVal As Long

'07/08/2003 - mpdea
'Implementado verificação com opção em Parâmetros da Filial
Public Sub InstanceControl(ByRef F As Form)
'''  Dim lngResult As Long
'''  Dim strCaption As String
'''
'''  Dim rstParametros As Recordset
'''  Dim strSQL As String
'''  Dim blnCheckInstance As Boolean
'''
'''
'''  strSQL = "SELECT CheckInstance FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial
'''  Set rstParametros = db.OpenRecordset(strSQL, dbOpenSnapshot)
'''  With rstParametros
'''    If Not (.BOF And .EOF) Then
'''      blnCheckInstance = .Fields("CheckInstance").Value
'''    End If
'''    .Close
'''  End With
'''  Set rstParametros = Nothing
'''
'''  If Not blnCheckInstance Then Exit Sub
'''
'''  'Procura janela atual
'''  strCaption = F.Caption
'''  F.Caption = F.Caption & " - Self"
'''  lngResult = FindWindow(vbNullString, strCaption)
'''  F.Caption = strCaption
'''
'''
'''  If lngResult <> 0 Then
'''
'''    '07/08/2003 - mpdea
'''    'Fecha conexão com a base de dados
'''    db.Close
'''    ws.Close
'''    Set db = Nothing
'''    Set ws = Nothing
'''
'''    strCaption = "Aplicativo Quick Store configurado para que não seja executado "
'''    strCaption = strCaption & "mais de uma vez na estação ao mesmo tempo."
'''    strCaption = strCaption & vbCrLf
'''    strCaption = strCaption & "Clique em [OK] para finalizar a aplicação."
'''
'''    MsgBox strCaption, vbInformation, "Aplicativo já está sendo executado"
'''
'''    End
'''  End If
  
  
End Sub

Public Function FindOpennedWindow(ByRef F As Form) As Boolean
  Dim lResult As Long
  Dim hWndOther As Integer
  Dim sCaption As String
  Dim bWhat As Boolean
  
  sCaption = F.Caption
  F.Caption = F.Caption & " - Self"
  hWndOther = FindWindow(0&, sCaption)
  bWhat = IIf(hWndOther = 0, False, True)
  F.Caption = sCaption
  If bWhat = True Then
    FindOpennedWindow = True
    lResult = ShowWindow(hWndOther, 1)
  Else
    FindOpennedWindow = False
  End If
  
End Function

Public Sub HideTitleBar(ByVal F As Form)
'   Change the style of the form to not show a title bar
    If mTitleBarHidden Then Exit Sub
    
    mTitleBarHidden = True
    
    With mStyle
        .OldStyle = GetWindowLong(F.hwnd, GWL_STYLE)
        .NewStyle = .OldStyle And WS_CAPTION_NOT
        SetWindowLong F.hwnd, GWL_STYLE, .NewStyle
    End With 'mStyle
End Sub

Public Sub ShowTitleBar(ByVal F As Form)
'   Change the style of the form to show a title bar
    If Not mTitleBarHidden Then Exit Sub
    mTitleBarHidden = False
    SetWindowLong F.hwnd, GWL_STYLE, mStyle.OldStyle
End Sub
