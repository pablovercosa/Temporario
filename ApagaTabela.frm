VERSION 5.00
Begin VB.Form frmPrecosResetTab 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apaga Tabela de Preços"
   ClientHeight    =   2295
   ClientLeft      =   1815
   ClientTop       =   2250
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   1040
   Icon            =   "ApagaTabela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2295
   ScaleWidth      =   5385
   Begin VB.CommandButton B_Apaga 
      Caption         =   "Apagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3945
      TabIndex        =   3
      Top             =   1785
      Width           =   1335
   End
   Begin VB.ComboBox Lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1230
      Width           =   1830
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ApagaTabela.frx":058A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Tabela 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tabela a apagar :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1260
      Width           =   1380
   End
End
Attribute VB_Name = "frmPrecosResetTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPreços As Recordset
Dim rsTabelas As Recordset

'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado o suporte a transação com tratamento a erro
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Apaga_Click()
  Dim Produto As Double
  
  Dim blnOnTransaction As Boolean
  
  On Error GoTo ErrHandler

  Call StatusMsg("")
  
  If IsNull(Lista.Text) Or Lista.Text = "" Then
    '09/07/2002 - mpdea
    'Mensagem e foco
    DisplayMsg "Selecione a tabela a ser apagada."
    Lista.SetFocus
    Exit Sub
  End If
  
  gsTitle = LoadResString(201)
  gsMsg = "Deseja apagar esta tabela de preços?"
  gnStyle = vbYesNo + vbQuestion
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    Exit Sub
  End If
  
  On Error GoTo ErrHandler
  
  Screen.MousePointer = vbHourglass
  Call ws.BeginTrans
  blnOnTransaction = True
  
  Produto = 0
  rsPreços.Index = "Tabela"
Lp1:
  rsPreços.Seek ">", Lista.Text, Produto
  If rsPreços.NoMatch Then GoTo Fim
  If rsPreços("Tabela") <> Lista.Text Then GoTo Fim
  
  Call StatusMsg("Apagando preço do produto " & rsPreços("Produto"))
  
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(rsPreços("Produto"))
  
  rsPreços.Delete
  GoTo Lp1
  
Fim:
  rsTabelas.Index = "Tabela"
  rsTabelas.Seek "=", Lista.Text
  If Not rsTabelas.NoMatch Then
    rsTabelas.Delete
  End If
  
  Call ws.CommitTrans
  Screen.MousePointer = vbDefault
  blnOnTransaction = False
  
  '09/07/2002 - mpdea
  'Remove o item da lista
  Lista.RemoveItem Lista.ListIndex
  
  Call StatusMsg("")
  DisplayMsg "Tabela de Preços apagada."
  
  Exit Sub
  
ErrHandler:
  Screen.MousePointer = vbDefault
  If blnOnTransaction Then ws.Rollback
  gsTitle = LoadResString(201)
  gsMsg = "Erro ao tentar apagar tabela de preços."
  gsMsg = gsMsg & vbCrLf & Err.Number & "-" & Err.Description
  gnStyle = vbOKOnly + vbExclamation
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
  End If
End Sub

Private Sub Form_Load()
  Dim Últ_Tabela As String
  Dim Lugar As Integer

  Call CenterForm(Me)
  
  Set rsPreços = db.OpenRecordset("Preços")
  Set rsTabelas = db.OpenRecordset("Tabela de Preços")

  Rem Pega as tabela usada e joga na lista
  rsPreços.Index = "Só Tabela"
  Lugar = 0
  Últ_Tabela = ""

  Do
    rsPreços.Seek ">", Últ_Tabela
    If Not rsPreços.NoMatch Then
       Últ_Tabela = rsPreços("Tabela")
       Lista.AddItem Últ_Tabela, Lugar
       Lugar = Lugar + 1
    End If
  Loop Until (rsPreços.NoMatch)

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call StatusMsg("")
End Sub
