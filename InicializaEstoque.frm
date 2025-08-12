VERSION 5.00
Begin VB.Form frmInicializaEstoque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Inicialização de Estoque do Produto"
   ClientHeight    =   2805
   ClientLeft      =   2385
   ClientTop       =   2130
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InicializaEstoque.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2805
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEstoque 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1380
      TabIndex        =   0
      Top             =   1650
      Width           =   1320
   End
   Begin VB.CommandButton B_OK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   430
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2130
      Width           =   6825
   End
   Begin VB.Label Label1 
      Caption         =   "Atenção:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Estoque Inicial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label L_Atenção 
      Caption         =   $"InicializaEstoque.frx":4E95A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   180
      TabIndex        =   2
      Top             =   450
      Width           =   6885
   End
End
Attribute VB_Name = "frmInicializaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gbPressOK As Boolean
Private gnValEstoque As Single
Private gsCodigo As String

Private Sub B_OK_Click()
  Dim bln_cancel As Boolean
  
  '16/12/2005 - mpdea
  'Validação de estoque
  Call txtEstoque_Validate(bln_cancel)
  If bln_cancel Then Exit Sub
  
  gnValEstoque = CSng(txtEstoque.Text)
  gbPressOK = True
  Unload Me
End Sub

Private Sub Form_Load()
  Call StatusMsg("")
  Call CenterForm(Me)
End Sub

Public Function gnInitializeEstoque(ByVal sCodigo As String) As Single
  gsCodigo = sCodigo
  Me.Show vbModal
  If gbPressOK Then
    gnInitializeEstoque = gnValEstoque
  Else
    gnInitializeEstoque = -1
  End If
  Set frmInicializaEstoque = Nothing
End Function


Private Sub txtEstoque_KeyPress(KeyAscii As Integer)
  KeyAscii = gnSomenteValor(KeyAscii)
End Sub

'16/12/2005 - mpdea
'Corrigido RT-6
Private Sub txtEstoque_Validate(Cancel As Boolean)
  Dim nQtdeCasaDec As Integer
  Dim sng_ret As Single
    
  If Not IsDataType(dtSingle, txtEstoque.Text, sng_ret) Then
    DisplayMsg "Valor de Estoque inválido."
    Call SelectAllText(txtEstoque)
    Cancel = True
    Exit Sub
  End If
  
  If sng_ret < 0 Then
    DisplayMsg "Valor de Estoque deve ser maior que 0."
    Call SelectAllText(txtEstoque)
    Cancel = True
    Exit Sub
  End If
    
  If sng_ret > 0 Then
    If gbIsFrac(gsCodigo, nQtdeCasaDec) Then
      txtEstoque.Text = Round(CSng(txtEstoque.Text), nQtdeCasaDec)
    Else
      txtEstoque.Text = Format(CSng(txtEstoque.Text), "#0")
    End If
  End If
End Sub
