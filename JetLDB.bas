Attribute VB_Name = "modJetLDB"
Option Explicit

Private Declare Function LDBUser_GetUsers Lib "MSLDBUSR.DLL" (lpszUserBuffer() As String, ByVal lpszFilename As String, ByVal nOptions As Long) As Integer
Private Declare Function LDBUser_GetError Lib "MSLDBUSR.DLL" (ByVal nErrorNo As Long) As String

Private Const OptAllLDBUsers = &H1
Private Const OptLDBLoggedUsers = &H2
Private Const OptLDBCorruptUsers = &H4
Private Const OptLDBUserCount = &H8
Private Const OptLDBUserAuthor = &HB0B

Public Function gsGetMDBUsers(ByVal sFileName As String) As String()
  ReDim lpszUserBuffer(1) As String
  Dim sError As String
  Dim nUsers As Long

  On Error GoTo ErrHandler

  nUsers = LDBUser_GetUsers(lpszUserBuffer(), sFileName, OptLDBLoggedUsers)

  If (nUsers = 0) Then
    lpszUserBuffer(0) = ""
    gsGetMDBUsers = lpszUserBuffer()
    Exit Function
  End If

  If (nUsers < 0) Then
    lpszUserBuffer(0) = ""
    gsGetMDBUsers = lpszUserBuffer()
    Exit Function
  End If

  gsGetMDBUsers = lpszUserBuffer()
  Exit Function

ErrHandler:
  If Err.Number = 48 Then
    gsTitle = "Verificação de Base de Dados"
    gsMsg = "Componente Vital da Aplicação (MSLDBUSR.DLL) não foi encontrado na localização padrão."
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Else
    gsTitle = "Verificação de Base de Dados"
    gsMsg = "Erro: " & Err.Number & "-" & Err.Description
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
  Call DBEngine.Idle(dbRefreshCache)
  End

End Function
