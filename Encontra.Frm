VERSION 5.00
Begin VB.Form frmEncontra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encontra cliente / fornecedor"
   ClientHeight    =   2895
   ClientLeft      =   6615
   ClientTop       =   4605
   ClientWidth     =   4665
   HelpContextID   =   40
   Icon            =   "Encontra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2895
   ScaleWidth      =   4665
   Begin VB.CommandButton B_Recomeça 
      Caption         =   "Recomeçar"
      Height          =   400
      Left            =   1680
      TabIndex        =   11
      Top             =   2325
      Width           =   1335
   End
   Begin VB.CommandButton B_Procura 
      Caption         =   "Próximo"
      Default         =   -1  'True
      Height          =   400
      Left            =   3225
      TabIndex        =   10
      Top             =   2325
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Campo"
      Height          =   1095
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   1575
      Begin VB.OptionButton O_Ambos 
         Caption         =   "Ambos"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton O_Fantasia 
         Caption         =   "Fantasia"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton O_Razão 
         Caption         =   "Razão Social"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      Begin VB.OptionButton O_Todos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton O_Forn 
         Caption         =   "Fornecedores"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton O_Cli 
         Caption         =   "Clientes"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Nome_Encontrar 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Digite abaixo o nome ou parte do nome que deseja encontrar. "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmEncontra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCliFor As Recordset
Dim Último As Long



Private Sub B_Procura_Click()
 Dim Pos, Pos2 As Integer
 Dim Str_Achar As String

 If IsNull(Nome_Encontrar.Text) Then Exit Sub
 If Nome_Encontrar.Text = "" Then Exit Sub

 Nome_Encontrar.Enabled = False

 rsCliFor.Index = "Código"
 
Lp1:
 rsCliFor.Seek ">", Último
 If rsCliFor.NoMatch Then
   Beep
   MsgBox ("Final da Pesquisa")
   Exit Sub
 End If
 Último = rsCliFor("Código")
 
 If O_Cli = True Then
  If rsCliFor("Tipo") <> "C" Then GoTo Lp1
 End If
 
 If O_forn = True Then
  If rsCliFor("Tipo") <> "F" Then GoTo Lp1
 End If
 
 
 Pos2 = 0
 
 If O_Razão.Value = True Then
  Str_Achar = ""
  If Not IsNull(rsCliFor("Nome")) Then Str_Achar = rsCliFor("Nome")
  Pos = InStr(1, Str_Achar, Nome_Encontrar.Text, 1)
 End If
 
 If O_Fantasia.Value = True Then
  Str_Achar = ""
  If Not IsNull(rsCliFor("Fantasia")) Then Str_Achar = rsCliFor("Fantasia")
  Pos = InStr(1, Str_Achar, Nome_Encontrar.Text, 1)
 End If
 
 If O_Ambos.Value = True Then
  Str_Achar = ""
  If Not IsNull(rsCliFor("Nome")) Then Str_Achar = rsCliFor("Nome")
  Pos = InStr(1, Str_Achar, Nome_Encontrar.Text, 1)
  
  Str_Achar = ""
  If Not IsNull(rsCliFor("Fantasia")) Then Str_Achar = rsCliFor("Fantasia")
  Pos2 = InStr(1, Str_Achar, Nome_Encontrar.Text, 1)
 End If
 
 If Pos = 0 And Pos2 = 0 Then GoTo Lp1
 
 frmCliFor.cboCodigo.Text = rsCliFor("Código")
 frmCliFor.cboCodigo.DoClick
 
End Sub


Private Sub B_Recomeça_Click()
 Call frmCliFor.ClearScreen(1)
 Nome_Encontrar.Enabled = True
 frmEncontra.Nome_Encontrar.SetFocus
End Sub

Private Sub Form_Load()
 Set rsCliFor = db.OpenRecordset("Cli_For", , dbReadOnly)
 Último = 0
End Sub


Private Sub SSCommand1_Click()
SendKeys "{F1}"
End Sub


