VERSION 5.00
Begin VB.Form frmEncontraProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encontra produto"
   ClientHeight    =   1515
   ClientLeft      =   6615
   ClientTop       =   4605
   ClientWidth     =   4590
   HelpContextID   =   40
   Icon            =   "EncontraProduto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1515
   ScaleWidth      =   4590
   Begin VB.CommandButton B_Recomeça 
      Caption         =   "Recomeça"
      Height          =   375
      Left            =   1785
      TabIndex        =   3
      Top             =   1050
      Width           =   1140
   End
   Begin VB.CommandButton B_Procura 
      Caption         =   "Procurar Próximo"
      Height          =   375
      Left            =   3045
      TabIndex        =   2
      Top             =   1050
      Width           =   1455
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
Attribute VB_Name = "frmEncontraProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsProdutos As Recordset
Dim Último As String

Private Sub B_Procura_Click()
 Dim Pos, Pos2 As Integer
 Dim Str_Achar As String

 If IsNull(Nome_Encontrar.Text) Then Exit Sub
 If Nome_Encontrar.Text = "" Then Exit Sub

 Nome_Encontrar.Enabled = False

 rsProdutos.Index = "Código"
 
Lp1:
 rsProdutos.Seek ">", Último
 If rsProdutos.NoMatch Then
   Beep
   MsgBox ("Final da Pesquisa")
   Exit Sub
 End If
 Último = rsProdutos("Código")
 
 
 Pos = 0
 
 Str_Achar = ""
 Str_Achar = rsProdutos("Nome")
 Pos = InStr(1, Str_Achar, Nome_Encontrar.Text, 1)
 
 If Pos = 0 Then GoTo Lp1
 
 frmProdutos.cboCodigo.Text = rsProdutos("Código")
 frmProdutos.cboCodigo.DoClick
 'Call frmProdutos.Código_LostFocus
 
End Sub

Private Sub cmdSearchProd_Click()

End Sub

Private Sub B_Recomeça_Click()
 Call frmProdutos.ClearScreen(1)
 Nome_Encontrar.Enabled = True
 Nome_Encontrar.SetFocus
 Nome_Encontrar.Text = ""
 Último = "0"
End Sub

Private Sub Form_Load()
 Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
 Último = "0"
End Sub





