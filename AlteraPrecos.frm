VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPrecosAltera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Preços"
   ClientHeight    =   6255
   ClientLeft      =   2475
   ClientTop       =   1230
   ClientWidth     =   7710
   HelpContextID   =   1030
   Icon            =   "AlteraPrecos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Caption         =   "Opções"
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7515
      Begin VB.TextBox txtFabricante 
         Height          =   315
         Left            =   5280
         MaxLength       =   15
         TabIndex        =   1
         Top             =   315
         Width           =   2055
      End
      Begin VB.CheckBox chkPrecoZero 
         Caption         =   "&Mostrar apenas produtos com preço 0 (zero)"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   3435
      End
      Begin VB.CheckBox chkContaClientes 
         Caption         =   "&Refletir alteração também na Conta de Clientes"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   3660
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "AlteraPrecos.frx":058A
         DataSource      =   "datSubclasse"
         Height          =   315
         Left            =   1695
         TabIndex        =   3
         Top             =   1095
         Width           =   735
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
         Columns(0).Width=   5398
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1402
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "AlteraPrecos.frx":05A5
         DataSource      =   "datClasse"
         Height          =   315
         Left            =   1695
         TabIndex        =   2
         Top             =   735
         Width           =   735
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
         Columns(0).Width=   6244
         Columns(0).Caption=   "Nome"
         Columns(0).Name =   "Nome"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Nome"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1455
         Columns(1).Caption=   "Código"
         Columns(1).Name =   "Código"
         Columns(1).Alignment=   1
         Columns(1).CaptionAlignment=   1
         Columns(1).DataField=   "Código"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboLista 
         Bindings        =   "AlteraPrecos.frx":05BD
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   315
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
      Begin VB.Label Label1 
         Caption         =   "Tabela de Preços :"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fabricante :"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   12
         Top             =   315
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Classe :"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   735
         Width           =   855
      End
      Begin VB.Label lblClasse 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2535
         TabIndex        =   10
         Top             =   735
         Width           =   4800
      End
      Begin VB.Label Label1 
         Caption         =   "Sub Classe :"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblSubClasse 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2535
         TabIndex        =   8
         Top             =   1095
         Width           =   4800
      End
   End
   Begin VB.Data datPrecos 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Tabela FROM [Tabela de Preços] ORDER BY Tabela"
      Top             =   6600
      Width           =   1710
   End
   Begin SSDataWidgets_B.SSDBGrid grdPrecos 
      Height          =   3360
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Utilize as teclas Ctrl e/ou Shift para selecionar as linhas de registros"
      Top             =   2400
      Width           =   7515
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   3
      AllowDelete     =   -1  'True
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3572
      Columns(0).Caption=   "Código "
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3043
      Columns(1).Caption=   "Preço"
      Columns(1).Name =   "Preco"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   5636
      Columns(2).Caption=   "Nome"
      Columns(2).Name =   "Nome"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      _ExtentX        =   13256
      _ExtentY        =   5927
      _StockProps     =   79
   End
   Begin VB.Data datSubclasse 
      Caption         =   "Data3"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6570
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data datClasse 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1860
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6555
      Visible         =   0   'False
      Width           =   1695
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   120
      Top             =   5760
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
      Bands           =   "AlteraPrecos.frx":05D5
   End
End
Attribute VB_Name = "frmPrecosAltera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsClasses As Recordset
Dim rsSubclasses As Recordset
Private rsContaCliente As Recordset
Private gsWhere As String
Private gsOrder As String
Private gbChanged As Boolean

Private Sub SearchPrecos()
  Call StatusMsg("")
  
  grdPrecos_LostFocus
  cboClasse_LostFocus
  cboSubClasse_LostFocus
  
  If cboLista.Text = "" Then
    DisplayMsg "Escolha uma tabela de preços antes."
    cboLista.SetFocus
    Exit Sub
  End If
  
  If gbChanged = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja pesquisar sem gravar alterações?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
  End If
    
'  grdPrecos.RemoveAll
  
  gsWhere = ""
  If txtFabricante.Text <> "" Then
    gsWhere = gsWhere & " And Produtos.Fabricante = '" & txtFabricante.Text & "'"
  End If
  If lblClasse.Caption <> "" Then
    gsWhere = gsWhere & " And Produtos.Classe = " & Val(cboClasse.Text)
  End If
  If lblSubClasse.Caption <> "" Then
    gsWhere = gsWhere & " And Produtos.[Sub Classe] = " & Val(cboSubClasse.Text)
  End If
  If chkPrecoZero.Value = vbChecked Then
    gsWhere = gsWhere & " And Preços.Preço = 0"
  End If
  gsWhere = gsWhere & " "
  Call LoadGridPrecos
  
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
      Call WriteGridPrecos
    Case "miOpDelete"
      Call DeleteRecord
    Case "miOpDeleteAll"
      Call WriteGridPrecos(True)
    Case "miOpSearch"
      Call SearchPrecos
  End Select
End Sub

Private Sub ActiveBar1_ComboSelChange(ByVal Tool As ActiveBarLibraryCtl.Tool)
  gsOrder = ""
  Select Case Tool.Name
    Case "miOpOrdem"
      Select Case Tool.CBListIndex
        Case 0 '"Por Código"
          gsOrder = "ORDER BY Produtos.[Código Ordenação]"
        Case 1 '"Por Nome"
          gsOrder = "ORDER BY Produtos.Nome"
      End Select
  End Select
End Sub
  
Private Sub chkContaClientes_Click()
  If chkContaClientes.Value = vbChecked Then
    If Not frmGerente.gbSenhaGerente Then
      chkContaClientes.Value = vbUnchecked
      Exit Sub
    End If
  End If
End Sub

Private Sub cboClasse_CloseUp()
  cboClasse.Text = cboClasse.Columns(1).Text
  cboClasse_LostFocus
End Sub

Private Sub cboClasse_LostFocus()
  Dim Aux As Variant
  
  lblClasse.Caption = ""
  Aux = cboClasse.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsClasses.Index = "Código"
  rsClasses.Seek "=", Val(Aux)
  If rsClasses.NoMatch Then Exit Sub
  
  lblClasse.Caption = rsClasses("Nome")
  
End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.Text = cboSubClasse.Columns(1).Text
  cboSubClasse_LostFocus
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim Aux As Variant
  
  lblSubClasse.Caption = ""
  Aux = cboSubClasse.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsSubclasses.Index = "Código"
  rsSubclasses.Seek "=", Val(Aux)
  If rsSubclasses.NoMatch Then Exit Sub
  
  lblSubClasse.Caption = rsSubclasses("Nome")

End Sub

Private Sub Command1_Click()

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
  
  Screen.MousePointer = vbHourglass
  
  Call CenterForm(Me)
  
  gsOrder = "ORDER BY Produtos.[Código Ordenação]"
  
  With ActiveBar1
    .Tools("miOpOrdem").CBList.Clear
    .Tools("miOpOrdem").CBList.InsertItem 0, "Por Código"
    .Tools("miOpOrdem").CBList.InsertItem 1, "Por Nome"
    .Tools("miOpOrdem").Text = ActiveBar1.Tools("miOpOrdem").CBList(0)
    .RecalcLayout
    Call ActiveBarLoadToolTips(frmPrecosAltera)
    .Tools("miOpDeleteAll").Enabled = gbPodeApagar
  End With
  
  grdPrecos.AllowDelete = gbPodeApagar
  
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsContaCliente = db.OpenRecordset("SELECT * FROM [Conta Cliente]", dbOpenDynaset)
  
  datClasse.DatabaseName = gsQuickDBFileName
  Set datClasse.Recordset = rsClasses
  
  datSubClasse.DatabaseName = gsQuickDBFileName
  Set datSubClasse.Recordset = rsSubclasses
  
  datPrecos.DatabaseName = gsQuickDBFileName
  'Set datPrecos.Recordset = db.OpenRecordset("SELECT Tabela FROM [Tabela de Preços] ORDER BY Tabela", dbOpenDynaset)
  Set datPrecos.Recordset = db.OpenRecordset(SQL_CONS_TAB_PRECO_T2, dbOpenSnapshot)
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
  grdPrecos_LostFocus
  If gbChanged = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja gravar as modificações de preços?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton1
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbYes Then
      Call WriteGridPrecos
    End If
  End If
  rsClasses.Close
  rsSubclasses.Close
  rsContaCliente.Close
  Set rsClasses = Nothing
  Set rsSubclasses = Nothing
  Set rsContaCliente = Nothing
  Set frmPrecosAltera = Nothing
End Sub

Private Sub grdPrecos_AfterDelete(RtnDispErrMsg As Integer)
  grdPrecos.Scroll 0, -32767
  grdPrecos.Scroll 0, 32767
End Sub

Private Sub grdPrecos_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  Call StatusMsg("")
  If Not bGridBeforeDelete() Then
    Cancel = True
  Else
    gbChanged = True
  End If
End Sub

Private Sub grdPrecos_BeforeUpdate(Cancel As Integer)
  Dim cNewValue As Currency
  Dim bErro As Boolean
  
  On Error GoTo ErrHandler
  
  cNewValue = gsHandleNull(grdPrecos.Columns("Preco").Text)
  
  If IsNull(cNewValue) Then bErro = True
  If bErro = False Then If Not IsNumeric(cNewValue) Then bErro = True
  If bErro = False Then If CDbl(cNewValue) < 0 Then bErro = True
  If bErro = True Then
    DisplayMsg "Digite um valor maior ou igual a 0 (zero)."
    Cancel = True
    Exit Sub
  End If
  
  grdPrecos.Columns("Preco").Text = Format(cNewValue, FORMAT_VALUE)
 
  gbChanged = True
  
  Exit Sub
  
ErrHandler:
  Cancel = True
  
End Sub

Private Sub grdPrecos_LostFocus()
 If grdPrecos.RowChanged = True Then
   grdPrecos.Update
 End If
End Sub

Private Sub LoadGridPrecos()
  Dim rsPrecos As Recordset
  Dim sRecord As String
  Dim bAllow As Boolean
  Dim sCodProd As String
  Dim nRow As Long
  Dim sSql As String
  
  Screen.MousePointer = vbHourglass
  
  sSql = "SELECT Preços.Produto as Produto, Preços.Preço as Preco, Produtos.Nome as Nome FROM Preços"
  sSql = sSql & " INNER JOIN Produtos ON [Preços].Produto = Produtos.Código"
  sSql = sSql & " WHERE Tabela = '" & cboLista.Text & "'"
  sSql = sSql & " AND Preços.Produto <> '0'"
  
  Set rsPrecos = db.OpenRecordset(sSql & gsWhere & gsOrder, dbOpenDynaset)

  If rsPrecos.EOF And rsPrecos.BOF Then
    DisplayMsg "Nenhum registro cadastrado com estas opções."
    rsPrecos.Close
    Set rsPrecos = Nothing
    Call StatusMsg("")
    Screen.MousePointer = vbDefault
    If cboLista.Enabled Then
      cboLista.SetFocus
    End If
    Exit Sub
  End If
  
  gbChanged = False
  
  grdPrecos.Redraw = False
  
'  bAllow = grdPrecos.AllowAddNew
'  grdPrecos.AllowAddNew = True
'  grdPrecos.AllowUpdate = True
  grdPrecos.RemoveAll
    
  cboLista.Enabled = False
  txtFabricante.Enabled = False
  cboClasse.Enabled = False
  cboSubClasse.Enabled = False
  chkPrecoZero.Enabled = False
  
  Call StatusMsg("Aguarde...")
  
  With rsPrecos
    Do While Not .EOF
      sRecord = .Fields("Produto").Value & vbTab & _
        Format(.Fields("Preco").Value, FORMAT_VALUE) & vbTab & _
        .Fields("Nome").Value
      grdPrecos.AddItem sRecord
      .MoveNext
    Loop
    .MoveFirst
  End With
  grdPrecos.Scroll -99, -99

  Call StatusMsg("")

'  grdPrecos.AllowAddNew = bAllow
'  grdPrecos.AllowUpdate = bAllow

  grdPrecos.Redraw = True
  
  rsPrecos.Close
  Set rsPrecos = Nothing
  
  Screen.MousePointer = vbDefault

End Sub

'-----------------------------------------------------------------------------------
'05/07/2002 - mpdea
'Implementado a atualização de sincronismo a produtos do tipo WEB com a Loja Virtual
'-----------------------------------------------------------------------------------
Private Sub WriteGridPrecos(Optional ByVal bOnlyDelete As Boolean = False)
  Dim rsPrecos As Recordset
  Dim rsPrecos2 As Recordset
  Dim nRow As Long
  Dim sSql As String
  Dim bm As Variant
  Dim sCodProd As String
  Dim bUpdateClientAccount As Boolean
  Dim sTabelaPrecos As String
  
  grdPrecos_LostFocus
  If Not gbChanged And Not bOnlyDelete Then
    Exit Sub
  End If
  
  If bOnlyDelete And grdPrecos.Rows > 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja realmente apagar TODOS os preços da pesquisa atual?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
    Call StatusMsg("Apagando tabela de preços...")
  ElseIf bOnlyDelete And grdPrecos.Rows = 0 Then
    Exit Sub
  Else
    Call StatusMsg("Gravando tabela de preços...")
  End If
  
  Screen.MousePointer = vbHourglass
  
  If chkContaClientes.Value = vbChecked Then
    bUpdateClientAccount = True
  End If
  sTabelaPrecos = cboLista.Text
  
  sSql = "SELECT Preços.Produto as Produto, Preços.Preço as Preco, Produtos.Nome as Nome FROM Preços"
  sSql = sSql + " INNER JOIN Produtos ON [Preços].Produto = Produtos.Código"
  sSql = sSql + " WHERE Tabela = '" & sTabelaPrecos & "'"
  '16/03/2007 - Anderson
  'Retirada essa condições porque estava causando erro ao atualizar preços com valores igual a zero
  'sSql = sSql + " AND Preços.Produto <> '0'"
  
  Set rsPrecos = db.OpenRecordset(sSql & gsWhere, dbOpenDynaset)

  sSql = "SELECT * FROM Preços"
  sSql = sSql + " WHERE Tabela = '" & sTabelaPrecos & "'"
  '16/03/2007 - Anderson
  'Retirada essa condições porque estava causando erro ao atualizar preços com valores igual a zero
  'sSql = sSql + " AND Preços.Produto <> '0'"
  
  Set rsPrecos2 = db.OpenRecordset(sSql, dbOpenDynaset)
  
    
  On Error GoTo ErrTrans
  
  Call ws.BeginTrans
  
  With rsPrecos
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        rsPrecos2.FindFirst "Produto = '" & .Fields("Produto") & "'"
        If Not rsPrecos2.NoMatch Then
          rsPrecos2.Delete
        End If
        .MoveNext
      Loop
    End If
    If Not bOnlyDelete Then
      For nRow = 0 To grdPrecos.Rows - 1
        bm = grdPrecos.AddItemBookmark(nRow)
        sCodProd = grdPrecos.Columns("Codigo").CellText(bm)
        If Len(sCodProd) > 0 Then
          With rsPrecos2
            .AddNew
            .Fields("Tabela").Value = sTabelaPrecos
            .Fields("Produto").Value = sCodProd
            .Fields("Preço").Value = CSng(gsHandleNull(grdPrecos.Columns("Preco").CellText(bm) & ""))
            .Fields("Data Alteração").Value = Format(Date, "dd/mm/yyyy")
            'Atualiza a Conta de Clientes
            If bUpdateClientAccount Then
              Call UpdateContaClientes(sTabelaPrecos, sCodProd, .Fields("Preço").Value)
            End If
            'Atualiza o sincronismo para o produto WEB alterado
            Call WEB_SynchronizeProduct(sCodProd)
            .Update
          End With
          
        End If
      Next nRow
    End If
  End With

  Call ws.CommitTrans
  
  gbChanged = False
  
  If bOnlyDelete Or grdPrecos.Rows = 0 Then
    grdPrecos.RemoveAll
    Call ClearScreen
  End If
  
  datPrecos.Refresh
  cboLista.Refresh
  
  Screen.MousePointer = vbDefault
  
  Call StatusMsg("")
  
  rsPrecos.Close
  Set rsPrecos = Nothing
  rsPrecos2.Close
  Set rsPrecos2 = Nothing
  
  Exit Sub
  
ErrTrans:
  ws.Rollback
  Screen.MousePointer = vbDefault
  gsTitle = LoadResString(221)
  If bOnlyDelete Then
    gsMsg = "Erro ao apagar Item de Preços."
  Else
    gsMsg = "Erro ao gravar Item de Preços."
  End If
  gsMsg = gsMsg & vbCrLf & Err.Number & " - " & Err.Description
  gnStyle = vbOKOnly + vbCritical
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)

End Sub

Private Sub ClearScreen()
  If gbChanged = True Then
    gsTitle = LoadResString(201)
    gsMsg = "Deseja inicializar a tela sem gravar atualizações?"
    gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    If gnResponse = vbNo Then
      Exit Sub
    End If
  End If
  cboLista.Enabled = True
  cboLista.Text = ""
  txtFabricante.Enabled = True
  txtFabricante.Text = ""
  cboClasse.Enabled = True
  cboClasse.Text = ""
  lblClasse.Caption = ""
  cboSubClasse.Enabled = True
  cboSubClasse.Text = ""
  lblSubClasse.Caption = ""
  chkPrecoZero.Enabled = True
  chkPrecoZero.Value = vbUnchecked
  chkContaClientes.Value = vbUnchecked
  grdPrecos.RemoveAll
  cboLista.SetFocus
  gbChanged = False
End Sub

Private Sub MoveFirst()
  On Error Resume Next
  grdPrecos.MoveFirst
  On Error GoTo 0
End Sub

Private Sub MoveLast()
  On Error Resume Next
  grdPrecos.MoveLast
  On Error GoTo 0
End Sub

Private Sub MovePrevious()
  On Error Resume Next
  grdPrecos.MovePrevious
  On Error GoTo 0
End Sub

Private Sub MoveNext()
  On Error Resume Next
  grdPrecos.MoveNext
  On Error GoTo 0
End Sub

Private Sub DeleteRecord()
  Dim Resposta As Integer
  Dim Num_Registro2 As Variant
  If grdPrecos.SelBookmarks.Count = 0 Then
    DisplayMsg "Selecione linhas para apagar."
    Exit Sub
  End If
  grdPrecos.DeleteSelected
End Sub
