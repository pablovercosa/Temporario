VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCliForCaract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de Características para Clientes/Fornecedores"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   Icon            =   "CliForCaract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9195
   Begin SSDataWidgets_B.SSDBGrid grdCodDesc 
      Bindings        =   "CliForCaract.frx":058A
      Height          =   4380
      Left            =   75
      TabIndex        =   0
      Top             =   675
      Width           =   9060
      _Version        =   196617
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BackColorOdd    =   12648447
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   1164
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "CodCaract"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   13758
      Columns(1).Caption=   "Nome ou Descrição"
      Columns(1).Name =   "Nome ou Descrição"
      Columns(1).CaptionAlignment=   0
      Columns(1).AllowSizing=   0   'False
      Columns(1).DataField=   "DescCaract"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "TipoCliCaract"
      Columns(2).Name =   "TipoCliCaract"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "TipoCliCaract"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   15981
      _ExtentY        =   7726
      _StockProps     =   79
      Caption         =   "Tabela de Características para Clientes/Fornecedores"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   5535
      Begin VB.OptionButton optTipo 
         Caption         =   "&Cliente"
         Height          =   255
         Index           =   0
         Left            =   285
         TabIndex        =   9
         Tag             =   "C"
         Top             =   195
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Fornecedor"
         Height          =   255
         Index           =   1
         Left            =   1365
         TabIndex        =   8
         Tag             =   "F"
         Top             =   195
         Width           =   1215
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Revendedor"
         Height          =   255
         Index           =   2
         Left            =   2805
         TabIndex        =   7
         Tag             =   "R"
         Top             =   195
         Width           =   1335
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Outros"
         Height          =   255
         Index           =   3
         Left            =   4365
         TabIndex        =   6
         Tag             =   "O"
         Top             =   195
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkMultiSelect 
      Caption         =   "&Seleção múltipla de itens para Apagar"
      Height          =   270
      Left            =   6075
      TabIndex        =   2
      Top             =   5115
      Width           =   2970
   End
   Begin VB.CheckBox chkUCase 
      Caption         =   "&Maiúsculas habilitadas em novos textos"
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   5115
      Value           =   1  'Checked
      Width           =   3330
   End
   Begin VB.Data datMaster 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM TabCaractCliFor ORDER By TipoCliCaract, CodCaract"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label3 
      Caption         =   "Pressione ESC duas vezes para abandonar edição da linha atual"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   5670
      Width           =   4890
   End
   Begin VB.Label Label1 
      Caption         =   "Clique no cabeçalho da coluna para reclassificar"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   5445
      Width           =   4050
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   6330
      Top             =   5340
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
      Bands           =   "CliForCaract.frx":05A2
   End
End
Attribute VB_Name = "frmCliForCaract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bNew As Boolean
Private nGridError As Integer
Private nColIndex As Integer
Private gsCriteria As String
Private gvBM As Variant

Public gnMaxCod As Long

Private gsLayOutFileName As String

'01/11/2004 - Daniel
'Validação caso a tabela esteja vazia
Dim m_blnTableVazia As Boolean

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Dim nRow As Long
  Dim nCol As Integer
  
  Select Case Tool.Name
    Case "miOpFirst"
      On Error Resume Next
      Screen.ActiveForm.datMaster.Recordset.MoveFirst
    Case "miOpPrevious"
      On Error Resume Next
      Screen.ActiveForm.datMaster.Recordset.MovePrevious
      If Screen.ActiveForm.datMaster.Recordset.BOF Then
        Screen.ActiveForm.datMaster.Recordset.MoveNext
      End If
    Case "miOpNext"
      On Error Resume Next
      Screen.ActiveForm.datMaster.Recordset.MoveNext
      If Screen.ActiveForm.datMaster.Recordset.EOF Then
        Screen.ActiveForm.datMaster.Recordset.MovePrevious
      End If
    Case "miOpLast"
      On Error Resume Next
      Screen.ActiveForm.datMaster.Recordset.MoveLast
    
    Case "miOpClear"
      Screen.ActiveForm.datMaster.Recordset.AddNew
      Screen.ActiveForm.bNew = True
      SendKeys "{Tab}"
      SendKeys "{F5}"
      
    Case "miOpUpdate"
'      If Screen.ActiveForm.datMaster.Recordset.EditMode <> dbEditNone Then
'        Screen.ActiveForm.datMaster.Recordset.Update
'        Screen.ActiveForm.datMaster.Recordset.Bookmark = Screen.ActiveForm.datMaster.Recordset.LastModified
'      End If
      grdCodDesc.Update
      If Screen.ActiveForm.bNew Then
        MsgBox "Registro Inserido."
      Else
        MsgBox "Registro Atualizado."
      End If
      Screen.ActiveForm.bNew = False
    Case "miOpDelete"
      If grdCodDesc.SelBookmarks.Count > 0 Then
        grdCodDesc.DeleteSelected
      Else
        gsTitle = LoadResString(201)
        gsMsg = "Selecione pelo menos uma linha para a operação."
        gnStyle = vbOKOnly + vbInformation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      End If
    Case "miOpClear"
     '
    Case "miOpSearch"
      If m_blnTableVazia Then
        MsgBox "Nenhum ítem encontrado.", vbExclamation, "Quick Store"
        Exit Sub
      End If
      grdCodDesc.SelBookmarks.RemoveAll
      gvBM = datMaster.Recordset.Bookmark
      gsCriteria = "[" & grdCodDesc.Columns(1).DataField & "] Like '*" & ActiveBar1.Tools("miOpText").Text & "*'"
      datMaster.Recordset.FindFirst gsCriteria
      If datMaster.Recordset.NoMatch Then
        gsTitle = "Pesquisa de Texto"
        gsMsg = "Texto não encontrado."
        gnStyle = vbOKOnly + vbInformation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        Exit Sub
      End If
    '
    Case "miOpSearchNext"
      If Len(gsCriteria) > 0 Then
        datMaster.Recordset.FindNext gsCriteria
        If datMaster.Recordset.NoMatch Then
          gsTitle = "Pesquisa de Texto"
          gsMsg = "Nenhuma outra ocorrência encontrada."
          gnStyle = vbOKOnly + vbInformation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          If Not IsNull(gvBM) Then
            datMaster.Recordset.Bookmark = gvBM
          End If
          Exit Sub
        End If
      End If
    
  End Select

  On Error GoTo 0
  
  Exit Sub
  
ErrExit:
  Select Case Err.Number
    Case 3022  ' Duplicate Index
      MsgBox "Registro já existe no cadastro."
    Case 3314  ' Campo Obrigatório faltando
      Dim nControl As Integer
      MsgBox "Campo obrigatório no registro não foi entrado."
    Case 3315  ' Campo Nulo Invalido
      Dim sDesc As String
      sDesc = "3315 - " & Err.Description
      MsgBox "Campo não pode ficar sem valor."
    Case Else
      MsgBox Err.Number & " - " & Err.Description
  End Select
  On Error GoTo 0
  Exit Sub
  
End Sub

Private Sub chkMultiSelect_Click()
  If chkMultiSelect.Value = 1 Then
    grdCodDesc.SelectTypeRow = ssSelectionTypeMultiSelect
  Else
    grdCodDesc.SelectTypeRow = ssSelectionTypeSingleSelect
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call HandleKeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  ActiveBar1.Tools("miOpClear").Enabled = gbPodeGravar
  Call SetDBNametoDC
  Call GridInitStyle(grdCodDesc)
  grdCodDesc.AllowAddNew = gbPodeGravar
  grdCodDesc.AllowUpdate = gbPodeGravar
  grdCodDesc.AllowDelete = gbPodeApagar
  chkUCase.Enabled = gbPodeGravar
  chkMultiSelect.Enabled = gbPodeApagar
  Call ActiveBarLoadToolTips(Me)
  optTipo(0).Value = True
  Call optTipo_Click(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  If Screen.ActiveForm.datMaster.Recordset.EditMode <> dbEditNone Then
    Screen.ActiveForm.datMaster.Recordset.CancelUpdate
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub grdCodDesc_AfterUpdate(RtnDispErrMsg As Integer)
  If RtnDispErrMsg = True Then
    RtnDispErrMsg = False      'Turn off SSDBGrid default message
    Call GridAfterUpdate
    grdCodDesc.Columns(nColIndex).CellStyleSet ("color"), grdCodDesc.Row
    Exit Sub
  End If
End Sub

Private Sub grdCodDesc_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
  If chkUCase.Value = 1 Then
    grdCodDesc.Columns(ColIndex).Text = UCase(grdCodDesc.Columns(ColIndex).Text)
  End If
End Sub

Private Sub grdCodDesc_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  Dim nK As Integer
  Dim sTipo As String
  Dim bm As Variant
  Dim sSql As String
  Dim nCod As Integer
  Dim nRow As Long
  DispPromptMsg = False
  If gbPodeApagar = False Then
    Beep
    Cancel = True
    Exit Sub
  End If
  If Len(Trim(grdCodDesc.ActiveCell.Text)) = 0 Then
    If bGridBeforeDelete(datMaster.Recordset) = True Then
      For nK = 0 To 3
        If optTipo(nK).Value = True Then
          sTipo = optTipo(nK).Tag
          Exit For
        End If
      Next nK
      Call ws.BeginTrans
      For nRow = 0 To grdCodDesc.SelBookmarks.Count - 1
        bm = grdCodDesc.SelBookmarks(nRow)
        nCod = grdCodDesc.Columns(0).CellText(bm)
        sSql = "DELETE * FROM CliForCaract WHERE TipoCliCaract = '" & sTipo & "' AND CodCaract = " & nCod
        Call db.Execute(sSql, dbFailOnError)
      Next nRow
      Call ws.CommitTrans
      gsTitle = LoadResString(201)
      gsMsg = "Seleção de itens apagada."
      gnStyle = vbOKOnly + vbInformation
      gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Cancel = False
    Else
      Cancel = True
    End If
  Else
    grdCodDesc.ActiveCell.Text = Mid(grdCodDesc.ActiveCell.Text, 1, grdCodDesc.ActiveCell.SelStart) & _
      Mid(grdCodDesc.ActiveCell.Text, grdCodDesc.ActiveCell.SelStart + grdCodDesc.ActiveCell.SelLength + 2, Len(grdCodDesc.ActiveCell.Text))
    Cancel = True
  End If
End Sub

Private Sub grdCodDesc_BeforeUpdate(Cancel As Integer)
  Dim sCod As String
  Dim nK As Integer
  
  If Not gbPodeGravar Then
    Beep
    Cancel = True
    Exit Sub
  End If
  sCod = gsHandleNull(grdCodDesc.Columns(0).Text)
  If Not IsNumeric(sCod) Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor de código não é numérico."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    grdCodDesc.Col = 0
    grdCodDesc.ActiveCell.SelStart = 1
    grdCodDesc.ActiveCell.SelLength = gnMaxCod
    'grdCodDesc.CancelUpdate
    Cancel = True
    Exit Sub
  End If
  If CLng(sCod) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor de código não deve ser zero."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    grdCodDesc.Col = 0
    grdCodDesc.ActiveCell.SelStart = 1
    grdCodDesc.ActiveCell.SelLength = gnMaxCod
    'grdCodDesc.CancelUpdate
    Cancel = True
    Exit Sub
  End If
  If Len(Trim(grdCodDesc.Columns(0).Text)) = 0 Or Len(Trim(grdCodDesc.Columns(1).Text)) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Entre com algum texto para a coluna."
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    grdCodDesc.Col = 0
    Cancel = True
  End If
  
  For nK = 0 To 3
    If optTipo(nK).Value = True Then
      grdCodDesc.Columns(2).Text = optTipo(nK).Tag
      Exit For
    End If
  Next nK
  
End Sub

Private Sub grdCodDesc_HeadClick(ByVal ColIndex As Integer)
  Dim nCol As Integer
  Dim sSQLOrdem As String
  Dim sSql As String
  Dim nPos As Integer
  sSQLOrdem = " ORDER BY [" + Trim(grdCodDesc.Columns(ColIndex).DataField) + "]"
  sSql = datMaster.RecordSource
  nPos = InStr(sSql, "ORDER BY")
  If nPos > 0 Then
    sSql = Left(sSql, nPos - 1)
  End If
  datMaster.RecordSource = sSql & sSQLOrdem
  datMaster.Refresh

End Sub

Private Sub grdCodDesc_InitColumnProps()
  grdCodDesc.Columns(0).FieldLen = 3
  grdCodDesc.Columns(1).FieldLen = datMaster.Recordset.Fields("DescCaract").Size
End Sub

Private Sub grdCodDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If grdCodDesc.Col = 0 Then
    If KeyCode = vbKeyF5 Then
      Call GetNewCode(Me, datMaster.Recordset, gnMaxCod)
    End If
  End If
End Sub

Private Sub grdCodDesc_KeyPress(KeyAscii As Integer)
  If gbSkipKey = True Then
    KeyAscii = 0
    gbSkipKey = False
    Exit Sub
  End If
  If KeyAscii = vbKeyReturn Then
    Beep
    KeyAscii = 0
  End If
End Sub

Private Sub grdCodDesc_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  Call GridRowColChange(grdCodDesc)
End Sub

Private Sub grdCodDesc_UpdateError(ByVal ColIndex As Integer, Text As String, ErrCode As Integer, ErrString As String, Cancel As Integer)
  nColIndex = ColIndex
End Sub

Private Sub SetDBNametoDC()
  '01/11/2004 - Daniel
  'Adicionado rotina para verificar se a tabela está vazia
  Dim rstCarac As Recordset
  
  datMaster.DatabaseName = gsQuickDBFileName
  
  '-------------------[Início]-------------------
  Set rstCarac = db.OpenRecordset("CliForCaract", dbOpenDynaset)
  
  If rstCarac.RecordCount = 0 Then m_blnTableVazia = True
  
  rstCarac.Close
  
  Set rstCarac = Nothing
  '-------------------[Fim]----------------------
  
End Sub

Private Function GridRowColChange(ByRef grdGrid As SSDBGrid)
  Dim bm As Variant
  On Error Resume Next
  If chkMultiSelect.Value = 0 Then
    grdGrid.SelBookmarks.Remove 0
    grdGrid.SelBookmarks.Add grdGrid.Bookmark
  End If
End Function

Public Sub GridInitStyle(ByRef grdGrid As SSDBGrid)
  grdGrid.StyleSets("color").BackColor = RGB(0, 0, 255)
  grdGrid.StyleSets("color").ForeColor = RGB(255, 255, 255)
End Sub

Public Sub GridAfterUpdate()
  If DBEngine.Errors.Count Then
    gsTitle = LoadResString(201)
    Select Case DBEngine.Errors(0).Number
      Case 3022  ' Tentativa de inserir o mesmo item duas vezes...
        gsMsg = "Tentativa de inserir um item de código já existente."
      Case 3163  ' Field Overflow
        gsMsg = "Overflow"
      Case 3314, 3058  ' Invalid Use of Null to Required Field
        gsMsg = "Campo vazio inválido."
      Case 3421
        gsMsg = "Erro de conversão de dados."
      Case Else
        gsMsg = "Erro:  " & DBEngine.Errors(0).Number & " - " & "Description: " & _
            DBEngine.Errors(0).Description
    End Select
    gnStyle = vbOKOnly + vbCritical
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  End If
End Sub

Public Function bGridBeforeDelete(ByRef rsData As Recordset) As Boolean
  gsTitle = LoadResString(201)
  gsMsg = "Atenção: Valores inseridos em ""Outros Dados"" no Cadastro de Clientes/Fornecedores "
  gsMsg = gsMsg & vbCrLf & "também serão excluídos nesta ação. Deseja apagar realmente a seleção atual?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  If gnResponse = vbNo Then
    bGridBeforeDelete = False
  Else
    If Not rsData.EOF Then
      If rsData.EditMode <> dbEditNone Then
        rsData.CancelUpdate
      End If
    End If
    bGridBeforeDelete = True
  End If
End Function

Private Sub GetNewCode(ByRef F As Form, ByRef rs As Recordset, ByVal nMaxCod As Long)
  Dim rsTemp As Recordset
  Dim nCod As Long
  
  nCod = 1
  Set rsTemp = rs.Clone
  With rsTemp
    If Not .EOF Then
      .MoveLast
      nCod = .Fields(0) + 1
      If nCod > nMaxCod Then
        For nCod = 1 To nMaxCod
          DoEvents
          .FindFirst .Fields(0).Name & " = " & nCod
          If .NoMatch Then
            F.grdCodDesc.Columns(0).Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
            Exit Sub
          End If
        Next nCod
        gsTitle = LoadResString(201)
        gsMsg = "Nenhum Código livre disponível para o intervalo."
        gnStyle = vbOKOnly + vbExclamation
        gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
      Else
        F.grdCodDesc.Columns(0).Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
      End If
    Else
      F.grdCodDesc.Columns(0).Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
    End If
    .Close
  End With
  
  Set rsTemp = Nothing
  
End Sub

Private Sub optTipo_Click(Index As Integer)
  Dim sTipo As String
  Dim sSql As String
  
  sTipo = optTipo(Index).Tag
  sSql = "SELECT CodCaract, TipoCliCaract, DescCaract FROM TabCaractCliFor "
  sSql = sSql & "WHERE TipoCliCaract = '" & sTipo & "' "
  sSql = sSql & "ORDER BY TipoCliCaract, CodCaract"
  
  Set datMaster.Recordset = db.OpenRecordset(sSql, dbOpenDynaset)
  If Not datMaster.Recordset.EOF Then
    datMaster.Recordset.MoveLast
    datMaster.Recordset.MoveFirst
  End If
  
  datMaster.Refresh
  Set grdCodDesc.DataSource = datMaster

End Sub
