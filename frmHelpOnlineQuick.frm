VERSION 5.00
Begin VB.Form frmHelpOnlineQuick 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Acessando o help on-line..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   3525
   End
End
Attribute VB_Name = "frmHelpOnlineQuick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Private Const conSwNormal = 1

Private Sub Form_Activate()
On Error GoTo Erro:

  Dim iIndice As Integer
  Dim sEnderecoWebHelp As String
  
  '''MsgBox "Você pode acessar o Help do QuickStore pelo seu Navegador de Internet." & vbCrLf & vbCrLf & "Digite:" & vbCrLf & "     https://www.infopar.com.br/help", vbInformation, "Help pelo seu Navegador"
  
  ShellExecute hwnd, "open", "https://www.infopar.com.br/help", vbNullString, vbNullString, 1 ' conSwNo
  
  MsgBox "Você pode acessar o Help do QuickStore pelo seu Navegador de Internet." & vbCrLf & vbCrLf & "Digite:" & vbCrLf & "     https://www.infopar.com.br/help", vbInformation, "Help pelo seu Navegador"
  
  
'''  If giQuick_viaRDP = 1 Then
'''    '''Shell "\\tsclient\c\windows\explorer.exe " & """http://www.infopar.com.br/index.html"""
'''    Call ExecCmd("\\tsclient\explorer.exe " & """https://www.infopar.com.br/help""")
'''  Else
'''    '''iIndice = InStr(1, sEnderecoValidaLicencaQuickStoreWEBAPI, "/geralapi/api/acao/validarLicenca")
'''    '''sEnderecoWebHelp = Mid(sEnderecoValidaLicencaQuickStoreWEBAPI, 1, iIndice)
'''
'''    Shell "explorer.exe " & """https://www.infopar.com.br/help"""
'''  End If

  Unload Me
  Exit Sub
Erro:
  MsgBox "Acesso não disponível do Help On-line. Detalhe: " & Err.Description, "Atenção"
  Unload Me
End Sub

