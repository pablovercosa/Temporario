VERSION 5.00
Begin VB.Form frmObsDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Carnês"
   ClientHeight    =   2730
   ClientLeft      =   3405
   ClientTop       =   2295
   ClientWidth     =   5355
   Icon            =   "ObservacaoCarne.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2730
   ScaleWidth      =   5355
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Limpar"
      Height          =   400
      Left            =   90
      TabIndex        =   11
      Top             =   2190
      Width           =   1335
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   400
      Left            =   3870
      TabIndex        =   5
      Top             =   2205
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2280
      TabIndex        =   4
      Top             =   2220
      Width           =   1335
   End
   Begin VB.TextBox Obs 
      Height          =   285
      Index           =   1
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Obs 
      Height          =   285
      Index           =   2
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1695
      Width           =   3615
   End
   Begin VB.TextBox Obs 
      Height          =   285
      Index           =   0
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Retorno 
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Observação 2 :"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Observação 3 :"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Observação 1 :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo de Configuração:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   300
      Width           =   1935
   End
End
Attribute VB_Name = "frmObsDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gsFileExt As String

Private Sub cmdClear_Click()
  Dim nI As Integer
  For nI = 0 To 2
    Obs(nI).Text = ""
  Next nI
End Sub

Private Sub Command1_Click()
  If Len(Trim(Combo.Text)) = 0 Then
    DisplayMsg "Boleto não informado. Verifique."
    Combo.SetFocus
    Exit Sub
  End If
  gsRetornoDoc = "OK"
  gsObsDoc(0) = Obs(0).Text
  gsObsDoc(1) = Obs(1).Text
  gsObsDoc(2) = Obs(2).Text
  gsDocFileName = Combo.Text
  Unload Me
End Sub


Private Sub Command2_Click()
  gsRetornoDoc = "NÃO"
  Unload Me
End Sub


Private Sub Form_Activate()
  Dim Nome As String
  Dim Aux As String
  Dim Aux2 As String
  Dim Tamanho As String
  Dim Fim_Loop As Integer
  
  Combo.Clear
  
  Aux = gsConfigPath & "*" & gsFileExt  '"*.CCA"
  Nome = Dir(Aux)
  If Nome = "" Then Exit Sub
  
  Tamanho = Len(Nome)
  Aux2 = Left$(Nome, (Tamanho - 4))
  
  Combo.AddItem Aux2
  
  Fim_Loop = False
  Do
   Nome = Dir
   If Nome = "" Then Fim_Loop = True
   If Nome <> "" Then
     Tamanho = Len(Nome)
     Aux2 = Left$(Nome, (Tamanho - 4))
     Combo.AddItem Aux2
   End If
  Loop Until Fim_Loop = True
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Call GetSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim nI As Integer
  Dim sObs As String
  For nI = 0 To 2
    sObs = "Obs" & CInt(nI)
    Call SaveSetting("QuickStore", "ObsDoc", sObs, Obs(nI))
  Next nI
End Sub

Private Sub GetSettings()
  Dim nI As Integer
  Dim sObs As String
  For nI = 0 To 2
    sObs = "Obs" & CInt(nI)
    Obs(nI).Text = GetSetting("QuickStore", "ObsDoc", sObs, "")
  Next nI
End Sub


