VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6135
   ClientLeft      =   930
   ClientTop       =   1305
   ClientWidth     =   8055
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFA324&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      ScaleHeight     =   510
      ScaleWidth      =   7785
      TabIndex        =   9
      Top             =   3960
      Width           =   7815
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
         Left            =   135
         TabIndex        =   11
         Top             =   15
         Width           =   5895
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6120
      Top             =   5520
   End
   Begin VB.CommandButton cmdSysInfo 
      BackColor       =   &H00FFA324&
      Caption         =   "&Sistema..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFA324&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Image imgLogo 
      Height          =   2730
      Left            =   1320
      Picture         =   "About.frx":4E95A
      Top             =   240
      Width           =   5415
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   8040
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   8040
      X2              =   8040
      Y1              =   0
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   0
      Y2              =   6120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   8040
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line linDivide 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   8040
      Y1              =   5040
      Y2              =   5040
   End
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Esta cópia está licenciada para"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   7785
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
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   7785
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
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4530
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "warning message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   6255
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
      Height          =   270
      Left            =   5160
      TabIndex        =   1
      Top             =   3360
      Width           =   2715
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'28/01/2009 - mpdea
'Adaptado para versão 7

' API declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
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

Private Sub Form_Load()
    
  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)

  lblWarning.Caption = "www.infopar.com.br"

' Fill in all of the information that comes from the App object
  With App
    Caption = "Infopar " & .ProductName
    lblTitle.Caption = .ProductName
    
    lblVersion.Caption = "Versão " & .Major & "." & .Minor & "." & .Revision
    lblCopyright.Caption = .LegalCopyright
    lblPathEXE.Caption = .Path & "\" & .EXEName & ".exe"
    lblFileDescription.Caption = LoadResString(1)
  End With 'App
  
  lblMisc.Caption = LoadResString(104)
      
  Screen.MousePointer = vbDefault

End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Public Sub About(frmParent As Form, Optional lUserName As String, _
                 Optional lUserInfo As String)
    cmdOK.Enabled = True
    cmdSysInfo.Enabled = True
    
'   Add user information to form
    If lUserName <> "" Then
        lblUserName.Caption = lUserName
        lblUserInfo.Caption = lUserInfo
    Else
        lblUserName.Caption = LoadResString(60)
        lblUserInfo.Caption = ""
    End If
        
    Show vbModal
End Sub

Public Sub SplashOn(Optional MinDisplay As Long, _
                    Optional lUserName As String, Optional lUserInfo As String)
    If Not Visible Then
        Dim lHeight As Integer
        
        cmdOK.Enabled = False
        cmdSysInfo.Enabled = False
    
'       Add user information to form
        If lUserName <> "" Then
            lblUserName.Caption = lUserName
            lblUserInfo.Caption = lUserInfo
        Else
            lblUserName.Caption = GetUserName
            lblUserInfo.Caption = GetUserCompany
        End If
        
        Height = linDivide(1).Y1 + 15 '(Height - ScaleHeight)
        
'       Show the form
        Show vbModeless

'       For some reason, need a Refresh to make sure Splash Screen gets painted
        Refresh
    End If
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

Private Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existence Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        Else
            GoTo SysInfoErr
        End If
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
    
SysInfoErr:
  MsgBox "Utilitário de informações gerais sobre o sistema não localizado em: " & SysInfoPath, vbOKOnly
    
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
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError               ' Handle Error...
    
    tmpVal = String$(1024, 0)                                    ' Allocate Variable Space
    KeyValSize = 1024                                            ' Mark Variable Size
    
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
        KeyVal = Format("&h" + KeyVal)                      ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:                                                ' Cleanup After An Error Has Occurred...
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

