VERSION 5.00
Begin VB.Form frmImpressaoComprovante 
   BackColor       =   &H00FFA324&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Comprovante"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpressaoComprovante.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraQtdeCopias 
      BackColor       =   &H00FFA324&
      Caption         =   "Deseja Imprimir quantas cópias:"
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
      Height          =   710
      Left            =   120
      TabIndex        =   8
      Top             =   1872
      Width           =   3525
      Begin VB.TextBox txtCopias 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Entre com a Quantidade de cópias a serem impressas."
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFA324&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3525
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdImprimirTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Imprimir"
         Default         =   -1  'True
         Height          =   375
         Left            =   150
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFA324&
      Caption         =   "Tipo "
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3525
      Begin VB.OptionButton optFatura 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Fatura"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optRecibo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Recibo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   810
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optBoleto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFA324&
         Caption         =   "Boleto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmImpressaoComprovante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public g_strTipoEmiss As String
Public g_blnFlag      As Boolean
'11/03/2004 - Daniel
'Var para gerenciamento da qtde de cópias
'a serem impressas
Public g_intCopias    As Integer

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdImprimirTipo_Click()

  If optBoleto.Value = True Then
    g_strTipoEmiss = "B"
  ElseIf optRecibo.Value = True Then
    g_strTipoEmiss = "R"
  ElseIf optFatura.Value = True Then
    g_strTipoEmiss = "F"
  End If

  g_blnFlag = True
  
  '11/03/2004 - Daniel
  If Len(txtCopias.Text) <= 0 Then
    g_intCopias = 1
  End If
  
  If Not IsNumeric(txtCopias.Text) Then
    txtCopias.Text = ""
    txtCopias.Text = "1"
  End If
  
  g_intCopias = CInt(txtCopias.Text)
  '------------------------------------
  
  frmManContasReceber.cmdEmiss_Click
  
  Unload Me

End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  '11/03/2004 - Daniel
  'Caso seja F. Linhares, irá sempre forçar o
  'usuário a imprimir o ticket
  If CheckSerialCaseMod("QS37818-990") Then
    cmdCancelar.Visible = False
  Else
    cmdCancelar.Visible = True
  End If
  
End Sub

