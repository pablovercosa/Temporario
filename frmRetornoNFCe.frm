VERSION 5.00
Begin VB.Form frmRetornoNFCe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retorno da NFCe"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   LinkTopic       =   "frmRetornoNFCe"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Status_Cancelamento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      HideSelection   =   0   'False
      Index           =   0
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6000
      Width           =   4935
   End
   Begin VB.TextBox Status_Autorizacao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      HideSelection   =   0   'False
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6000
      Width           =   4935
   End
   Begin VB.TextBox Numero_Protocolo_Autorizacao 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      HideSelection   =   0   'False
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   10095
   End
   Begin VB.TextBox Ex_Message 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      HideSelection   =   0   'False
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   3960
      Width           =   10095
   End
   Begin VB.TextBox Detalhe_Cancelamento 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      HideSelection   =   0   'False
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3120
      Width           =   10095
   End
   Begin VB.TextBox Detalhe_Autorizacao 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      HideSelection   =   0   'False
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   10095
   End
   Begin VB.TextBox Chave 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      HideSelection   =   0   'False
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   10095
   End
   Begin VB.Label lblStatusCancelamento 
      AutoSize        =   -1  'True
      Caption         =   "Status da Cancelamento"
      Height          =   195
      Left            =   5400
      TabIndex        =   13
      Top             =   5760
      Width           =   1740
   End
   Begin VB.Label lblStatusAutorizacao 
      AutoSize        =   -1  'True
      Caption         =   "Status da Autorização"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   1560
   End
   Begin VB.Label URL_QRCode 
      AutoSize        =   -1  'True
      Caption         =   "Visualizar a NFC-e no site da SEFAZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   7080
      TabIndex        =   10
      Top             =   360
      Width           =   3270
   End
   Begin VB.Label lblProtocolo 
      AutoSize        =   -1  'True
      Caption         =   "Protocolo"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label lblMensagem 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   780
   End
   Begin VB.Label lblDetalheCancelamento 
      AutoSize        =   -1  'True
      Caption         =   "Detalhe Cancelamento"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1620
   End
   Begin VB.Label lblDetalheAutorização 
      AutoSize        =   -1  'True
      Caption         =   "Detalhe Autorização"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1440
   End
   Begin VB.Label lblChave 
      AutoSize        =   -1  'True
      Caption         =   "Chave"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   465
   End
End
Attribute VB_Name = "frmRetornoNFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Sub URL_QRCode_Click()
    If CStr(URL_QRCode.Tag) = "" Then Exit Sub
    ShellExecute ByVal 0&, "open", CStr(URL_QRCode.Tag), vbNullString, vbNullString, 3
End Sub

Private Sub URL_QRCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, 32649&)
End Sub

Public Sub CarregaValores()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = CStr(ctrl.Tag)
        End If
    Next
    URL_QRCode.Visible = (URL_QRCode.Tag <> "")
End Sub
