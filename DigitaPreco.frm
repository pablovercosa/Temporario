VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosDigita 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Preços"
   ClientHeight    =   2820
   ClientLeft      =   2955
   ClientTop       =   3450
   ClientWidth     =   7245
   Icon            =   "DigitaPreco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "DigitaPreco.frx":058A
   ScaleHeight     =   2820
   ScaleWidth      =   7245
   Begin VB.CheckBox chkContaClientes 
      Caption         =   "&Refletir alteração também na Conta de Clientes"
      Height          =   345
      Left            =   165
      TabIndex        =   11
      Top             =   1845
      Width           =   3780
   End
   Begin VB.Data datPrecos 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   225
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela"
      Top             =   2970
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   750
      Left            =   5145
      TabIndex        =   7
      Top             =   1050
      Width           =   1905
      Begin VB.OptionButton O_Produto 
         Caption         =   "Produto + Tabela"
         Height          =   225
         Left            =   105
         TabIndex        =   9
         Top             =   420
         Width           =   1695
      End
      Begin VB.OptionButton O_Tabela 
         Caption         =   "Tabela + Produto"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Produto"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSMask.MaskEdBox Preço 
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Top             =   1155
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   "###,###,##0.00"
      PromptChar      =   "_"
   End
   Begin SSDataWidgets_B.SSDBCombo Combo_Produto 
      Bindings        =   "DigitaPreco.frx":1254
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   735
      Width           =   2010
      DataFieldList   =   "Nome"
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   8043
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3307
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   3545
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "DigitaPreco.frx":1268
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   285
      Width           =   2010
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
      _Version        =   196617
      Columns(0).Width=   3200
      _ExtentX        =   3545
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Tabela"
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   90
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "DigitaPreco.frx":1280
   End
   Begin VB.Label Label1 
      Caption         =   "Tabela :"
      Height          =   210
      Left            =   225
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label4 
      Height          =   345
      Left            =   30
      TabIndex        =   10
      Top             =   195
      Width           =   915
   End
   Begin VB.Label Nome_Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3150
      TabIndex        =   6
      Top             =   735
      Width           =   4005
   End
   Begin VB.Label Label3 
      Caption         =   "Preço :"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   1155
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Produto :"
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   735
      Width           =   750
   End
End
Attribute VB_Name = "frmPrecosDigita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPreços As Recordset
Dim rsProdutos As Recordset
Private rsConta_Cli As Recordset
Dim Num_Registro As Variant

Sub ShowRecord()

  cboLista.Text = rsPreços("Tabela")
  Combo_Produto.Text = rsPreços("Produto")
  Combo_Produto_LostFocus
  Preço.Text = rsPreços("Preço")

End Sub

Private Sub DeleteRecord()
  Call StatusMsg("")
  
  If IsNull(Num_Registro) Then
    DisplayMsg "Não existe registro a apagar."
    Exit Sub
  End If
  
  If MsgBox("Deseja realmente apagar este preço ?", vbQuestion + vbYesNo) = vbYes Then
    rsPreços.Delete
    Call ClearScreen
  End If

End Sub

'-----------------------------------------------------------------------------------
'08/07/2002 - mpdea
'Implementado o suporte a transação com tratamento a erro
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub UpdateRecord()
  Dim Novo_Preço As Double
  
  Dim blnOnTransaction As Boolean
  
  On Error GoTo ErrHandler
  
  Call StatusMsg("")
 
  If cboLista.Text = "" Then
    DisplayMsg "Tabela de preço inválida."
    cboLista.SetFocus
    Exit Sub
  End If
  
  If Nome_Produto.Caption = "" Then
    DisplayMsg "Produto inválido."
    Combo_Produto.SetFocus
    Exit Sub
  End If
  
  Call StatusMsg("Gravando...")
  Screen.MousePointer = vbHourglass
  ws.BeginTrans
  blnOnTransaction = True
  
  rsPreços.FindFirst "Tabela = '" & cboLista.Text & "' And Produto = '" & Combo_Produto.Text & "'"
  If rsPreços.NoMatch Then
    rsPreços.AddNew
    rsPreços("Tabela") = cboLista.Text
    rsPreços("Produto") = Combo_Produto.Text
  Else
    rsPreços.Edit
  End If
  rsPreços("Preço") = CDbl(Preço.Text)
  rsPreços("Data Alteração") = Format(Date, "dd/mm/yyyy")
  rsPreços.Update
   
  If chkContaClientes.Value = vbChecked Then
    Novo_Preço = Preço.Text
    Call UpdateContaClientes(cboLista.Text, Combo_Produto.Text, Novo_Preço)
  End If
          
  'Atualiza o sincronismo para o produto WEB alterado
  Call WEB_SynchronizeProduct(Combo_Produto.Text)
  
  ws.CommitTrans
  Screen.MousePointer = vbDefault
  blnOnTransaction = False
  
  Call StatusMsg("")
  Exit Sub

ErrHandler:
  Screen.MousePointer = vbDefault
  If blnOnTransaction Then ws.Rollback
  MsgBox "Erro [" & Err.Number & "] - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub ClearScreen()
  cboLista.Text = ""
  Combo_Produto.Text = ""
  Nome_Produto.Caption = ""
  Preço.Text = ""
  
  Combo_Produto.SetFocus
  
  Num_Registro = Null
  
End Sub

Private Sub cboLista_LostFocus()
 
  If IsNull(cboLista.Text) Then Exit Sub
  If cboLista.Text = "" Then Exit Sub
 
  If Nome_Produto.Caption = "" Then Exit Sub
  
  rsPreços.FindFirst "Tabela = '" & cboLista.Text & "' And Produto = '" & Combo_Produto.Text & "'"
  
  Num_Registro = Null

  If rsPreços.NoMatch Then Exit Sub

  Preço.Text = rsPreços("Preço")
  Num_Registro = rsPreços.Bookmark
  
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  With rsPreços
    .MoveFirst
    If .BOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  On Error GoTo 0
End Sub

Private Sub MoveLast()
  On Error Resume Next
  With rsPreços
    .MoveLast
    If .EOF Then
      Beep
    Else
      Call ShowRecord
    End If
  End With
  On Error GoTo 0
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  With rsPreços
    .MovePrevious
    If Not .BOF Then
      Call ShowRecord
    Else
      Beep
      .MoveNext
    End If
  End With
  On Error GoTo 0
End Sub

Private Sub MoveNext()
  On Error Resume Next
  With rsPreços
    .MoveNext
    If Not .EOF Then
      Call ShowRecord
    Else
      Beep
      .MovePrevious
    End If
  End With
  On Error GoTo 0
End Sub

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Select Case Tool.Name
    Case "miOpFirst"
      Call MoveFirst
    Case "miOpPrevious"
      Call MovePrevious
    Case "miOpNext"
      Call MoveNext
    Case "miOpLast"
      Call MoveLast
    Case "miOpClear"
      Call ClearScreen
    Case "miOpUpdate"
      Call UpdateRecord
    Case "miOpDelete"
      Call DeleteRecord
    Case "miOpSearch"
      'Call SearchRecord
  End Select
End Sub

Private Sub chkContaClientes_Click()
  If chkContaClientes.Value = vbChecked Then
    If Not frmGerente.gbSenhaGerente Then
      chkContaClientes.Value = vbUnchecked
      Exit Sub
    End If
'    frmGerente.Show vbModal
'    If gsRetornoDoc <> "OK" Then
'      chkContaClientes.Value = vbUnchecked
'      Exit Sub
'    End If
  End If
End Sub

Private Sub Combo_Produto_CloseUp()
  Combo_Produto.Text = Combo_Produto.Columns(1).Text
  Combo_Produto_LostFocus
End Sub

Private Sub Combo_Produto_LostFocus()
  Nome_Produto.Caption = ""
  If IsNull(Combo_Produto.Text) Then Exit Sub
  If Combo_Produto.Text = "" Then Exit Sub
  
  rsProdutos.Index = "Código"
  rsProdutos.Seek "=", Combo_Produto.Text
  If rsProdutos.NoMatch Then Exit Sub
  
  Nome_Produto.Caption = rsProdutos("Nome") & ""
  
  If IsNull(cboLista.Text) Then Exit Sub
  If cboLista.Text = "" Then Exit Sub
  
  If Nome_Produto.Caption = "" Then Exit Sub
  
  rsPreços.FindFirst "Tabela = '" & cboLista.Text & "' And Produto = '" & Combo_Produto.Text & "'"
  
  Num_Registro = Null
  
  If rsPreços.NoMatch Then Exit Sub
  
  Preço.Text = rsPreços("Preço")
  Num_Registro = rsPreços.Bookmark
  
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
  
  KeyPreview = True
  
  Call CenterForm(Me)
  
  Screen.MousePointer = vbHourglass
  
  Data1.DatabaseName = gsQuickDBFileName
  datPrecos.DatabaseName = gsQuickDBFileName
  
  Set rsPreços = db.OpenRecordset("SELECT * FROM Preços ORDER BY Tabela, Produto", dbOpenDynaset)
  Set rsProdutos = db.OpenRecordset("Produtos")
  Set rsConta_Cli = db.OpenRecordset("SELECT * FROM [Conta Cliente]", dbOpenDynaset)
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsPreços.Close
  rsProdutos.Close
  rsConta_Cli.Close
  Set rsPreços = Nothing
  Set rsProdutos = Nothing
  Set rsConta_Cli = Nothing
End Sub

Private Sub O_Produto_Click()
  Set rsPreços = db.OpenRecordset("SELECT * FROM Preços ORDER BY Produto, Tabela", dbOpenDynaset)
End Sub

Private Sub O_Tabela_Click()
  Set rsPreços = db.OpenRecordset("SELECT * FROM Preços ORDER BY Tabela, Produto", dbOpenDynaset)
End Sub

Private Sub Preço_GotFocus()
  Preço.SelStart = 0
  Preço.SelLength = Len(Preço.Text)
End Sub
