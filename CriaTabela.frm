VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosCriaTab 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cria��o de Tabela de Pre�os"
   ClientHeight    =   1980
   ClientLeft      =   1620
   ClientTop       =   2670
   ClientWidth     =   4470
   HelpContextID   =   1140
   Icon            =   "CriaTabela.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   4470
   Begin VB.Data datPrecos 
      Caption         =   "Preco"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   165
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1005
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton B_Cria 
      Caption         =   "Criar"
      Height          =   400
      Left            =   2925
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "CriaTabela.frx":058A
      Height          =   315
      Left            =   2340
      TabIndex        =   0
      Top             =   180
      Width           =   1935
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   3413
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Tabela"
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Nome da tabela a ser criada :"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   255
      Width           =   2295
   End
End
Attribute VB_Name = "frmPrecosCriaTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsProdutos As Recordset
Dim rsPre�os As Recordset
Dim rsTabelas As Recordset

'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado o suporte a transa��o com tratamento a erro
'Implementado a atualiza��o de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub B_Cria_Click()
  Dim Cod As String
  
  Dim blnOnTransaction As Boolean
  
  On Error GoTo ErrHandler

  Call StatusMsg("")

  If IsNull(cboLista.Text) Or cboLista.Text = "" Then
    DisplayMsg "Digite o nome da tabela a ser criada."
    cboLista.SetFocus
    Exit Sub
  End If

  If Len(cboLista.Text) > 15 Then
    DisplayMsg "O nome da Tabela de Pre�os deve ter at� 15 caracteres."
    cboLista.SetFocus
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass
  ws.BeginTrans
  blnOnTransaction = True
  
  cboLista.Text = UCase(cboLista.Text)
  
  rsProdutos.MoveFirst
  
  Cod = 0
  rsProdutos.Index = "C�digo"
  rsPre�os.Index = "Tabela"
  Do Until rsProdutos.NoMatch
   rsProdutos.Seek ">", Cod
   If Not rsProdutos.NoMatch Then
      Call StatusMsg("Criando tabela para produto " & Cod)
      Cod = rsProdutos("C�digo")
      If rsProdutos("Desativado") = False Then
        rsPre�os.Seek "=", cboLista.Text, Cod
        If rsPre�os.NoMatch Then
          rsPre�os.AddNew
          rsPre�os("Tabela") = cboLista.Text
          rsPre�os("Produto") = Cod
          rsPre�os("Pre�o") = 0
          rsPre�os("Data Altera��o") = Format(Date, "dd/mm/yyyy")
          rsPre�os.Update
          'Atualiza o sincronismo para o produto WEB alterado
          Call WEB_SynchronizeProduct(Cod)
        End If
      End If
   End If
  Loop

  Call CheckConfigTablePrice(cboLista.Text)
  
  ws.CommitTrans
  Screen.MousePointer = vbDefault
  blnOnTransaction = False
  
  datPrecos.Refresh
  cboLista.Refresh
  
  DisplayMsg "Tabela criada."
    
  Call StatusMsg("")
  Exit Sub

ErrHandler:
  Screen.MousePointer = vbDefault
  If blnOnTransaction Then ws.Rollback
  MsgBox "Erro [" & Err.Number & "] - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cboLista_KeyPress(KeyAscii As Integer)
  KeyAscii = gnLimitKeyPress(cboLista, 15, KeyAscii)
  If KeyAscii <> 0 Then
    KeyAscii = gnTypeValidKey(KeyAscii)
  End If
End Sub

Private Sub cboLista_LostFocus()
  If Len(Trim(cboLista.Text)) > 0 Then
    cboLista.Text = UCase(cboLista.Text)
  End If
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  Screen.MousePointer = vbHourglass
  Set rsPre�os = db.OpenRecordset("Pre�os")
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsTabelas = db.OpenRecordset("Tabela de Pre�os")
  datPrecos.DatabaseName = gsQuickDBFileName
  Set datPrecos.Recordset = db.OpenRecordset(SQL_CONS_TAB_PRECO_T1, dbOpenSnapshot)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsPre�os.Close
  rsProdutos.Close
  rsTabelas.Close
  Set rsPre�os = Nothing
  Set rsProdutos = Nothing
  Set rsTabelas = Nothing
End Sub
