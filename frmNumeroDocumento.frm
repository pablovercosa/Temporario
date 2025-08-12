VERSION 5.00
Begin VB.Form frmNumeroDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe o número do documento (CPF/CNPJ)"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "frmNumeroDocumento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   -120
      TabIndex        =   3
      Top             =   -120
      Width           =   5535
      Begin VB.Image Image1 
         Height          =   930
         Left            =   360
         Picture         =   "frmNumeroDocumento.frx":058A
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNumeroDocumento.frx":0BCE
         ForeColor       =   &H00808080&
         Height          =   975
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtNumeroDocumento 
      Height          =   375
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
End
Attribute VB_Name = "frmNumeroDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'29/04/2008 - mpdea
'Tela para informar o número do documento (CPF/CNPJ)

Private m_bln_press_ok As Boolean
Private m_str_nr_documento As String

Public Function Start(ByVal lngCodigoCliente As Long, ByVal strNumeroDocumentoDefault As String) As String
  Dim str_nr_documento As String
  
  'Valor default
  m_bln_press_ok = False
  
  'Obtém o número do documento
  If strNumeroDocumentoDefault <> "" Then
    str_nr_documento = strNumeroDocumentoDefault
  Else
    str_nr_documento = gstrGetCliForNumeroDocumento(lngCodigoCliente)
  End If
  txtNumeroDocumento.Text = FormatDocumento(str_nr_documento)
  
  'Exibe tela
  Me.Show vbModal
  
  'Retorno da função
  If m_bln_press_ok Then
    Start = m_str_nr_documento
  Else
    Start = ""
  End If
  
End Function

Private Sub cmdCancel_Click()
  m_bln_press_ok = False
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim str_nr_doc As String
  
  str_nr_doc = g_str_SomenteNumero(Trim(txtNumeroDocumento.Text))
  
  If str_nr_doc = "" Then
    If MsgBox("O número do documento não foi informado. Deseja continuar?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  
  If Len(str_nr_doc) <> 11 And Len(str_nr_doc) <> 14 Then
    DisplayMsg "O número do documento não foi informado corretamente."
    Exit Sub
  End If
  
  If Not (Valid_CPF(str_nr_doc) Or Valid_CGC(str_nr_doc)) Then
    DisplayMsg "Número do documento incorreto."
    Exit Sub
  End If

  m_str_nr_documento = str_nr_doc
  m_bln_press_ok = True
  Unload Me
End Sub

Private Function FormatDocumento(ByVal strNumeroDocumento As String) As String
  Dim str_ret As String
  
  Select Case Len(strNumeroDocumento)
    Case 11 'CPF
      str_ret = Format(strNumeroDocumento, "000.000.000-00")
    Case 14 'CNPJ
      str_ret = Format(strNumeroDocumento, "00.000.000/0000-00")
    Case Else
      str_ret = strNumeroDocumento
  End Select
  
  FormatDocumento = str_ret
End Function
