VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmTabela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Tabela"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tabela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9360
   Begin SSDataWidgets_B.SSDBGrid grdCodDesc 
      Bindings        =   "Tabela.frx":4E95A
      Height          =   4920
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9090
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
      ExtraHeight     =   53
      Columns.Count   =   3
      Columns(0).Width=   3704
      Columns(0).Caption=   "C�digo"
      Columns(0).Name =   "C�digo"
      Columns(0).Alignment=   1
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "C�digo"
      Columns(0).DataType=   3
      Columns(0).FieldLen=   256
      Columns(1).Width=   7752
      Columns(1).Caption=   "Nome ou Descri��o"
      Columns(1).Name =   "Nome ou Descri��o"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Nome"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Lucro M�nimo Permitido"
      Columns(2).Name =   "Lucro M�nimo Permitido"
      Columns(2).DataField=   "LucroMinimoPermitido"
      Columns(2).FieldLen=   256
      _ExtentX        =   16034
      _ExtentY        =   8678
      _StockProps     =   79
      Caption         =   "Tabela"
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
   Begin VB.CheckBox chkMultiSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Sele��o m�ltipla de �tens"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6960
      TabIndex        =   2
      Top             =   5160
      Width           =   2250
   End
   Begin VB.CheckBox chkUCase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Mai�sculas habilitadas em novos textos"
      ForeColor       =   &H80000008&
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
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT * FROM Cores ORDER By C�digo"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pressione ESC duas vezes para abandonar edi��o da linha atual"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   5760
      Width           =   4590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dicas: Clique no cabe�alho da coluna para reclassificar"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   5475
      Width           =   3900
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   8520
      Top             =   5520
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
      Bands           =   "Tabela.frx":4E972
   End
End
Attribute VB_Name = "frmTabela"
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
Private rsProdClasse As Recordset

Public Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  Dim nRow As Long
  Dim nCol As Integer
  Dim sSql As String
  Dim nI As Integer
  Dim vBook As Variant
  
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
        Call StatusMsg("Registro Inserido.")
      Else
        Call StatusMsg("Registro Atualizado.")
      End If
      Screen.ActiveForm.bNew = False
    Case "miOpDelete"
      '05/05/2005 - Daniel
      'Tratamento para n�o excluir Centro de Custo [Altera��o na beta 6.52.0.67 e inferiores]
      '
      '19/07/2005 - Daniel
      '[Altera��o da beta 6.52.0.68 em diante]
      'Foi solicitado pelo Carlos (OSM) para abrir novamente a possibilidade de exclus�o
      'do Centro de Custo mas damos a sugest�o atrav�s da mensagem para desabilitar
      If Me.Caption = "CENTROS DE CUSTO" Then
        Dim strTexto As String
      
        strTexto = ""
        strTexto = "Sobre a exclus�o de um ou mais Centro de Custo:" & vbCrLf & vbCrLf
        strTexto = strTexto & "O Quick Store recomenda que o Centro de Custo seja 'Desativado' atrav�s das rotinas" & vbCrLf
        strTexto = strTexto & "disponibilizadas para isso ao inv�s de 'excluir', evitando assim que algumas informa��es" & vbCrLf
        strTexto = strTexto & "fiquem incompletas no Banco de Dados." & vbCrLf
        strTexto = strTexto & "Exclua o Centro somente se ele n�o foi utilizado para nada." & vbCrLf & vbCrLf
        strTexto = strTexto & "Deseja mesmo excluir ?"
        
        If MsgBox(strTexto, vbQuestion + vbYesNo + vbDefaultButton2, "Aten��o") = vbYes Then
          If grdCodDesc.SelBookmarks.Count > 0 Then
             grdCodDesc.DeleteSelected
          Else
             gsTitle = LoadResString(201)
             gsMsg = "Selecione pelo menos uma linha para a opera��o."
             gnStyle = vbOKOnly + vbInformation
             gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          End If
        Else
          Exit Sub
        End If
      
        'MsgBox "Imposs�vel excluir o(s) Centro(s) de Custo, utilize o recurso de desativa��o.", vbExclamation, "Centros de Custo"
        'Exit Sub
      
      Else '<< Quando n�o for Centro de Custo >>
        If grdCodDesc.SelBookmarks.Count > 0 Then
           grdCodDesc.DeleteSelected
        Else
           gsTitle = LoadResString(201)
           gsMsg = "Selecione pelo menos uma linha para a opera��o."
           gnStyle = vbOKOnly + vbInformation
           gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
        End If
      End If
      '--------------------------------------------

    Case "miOpClear"
     '
    Case "miOpSearch"
      If Len(Trim(ActiveBar1.Tools("miOpText").Text)) > 0 Then
        grdCodDesc.SelBookmarks.RemoveAll
        gvBM = datMaster.Recordset.Bookmark
        gsCriteria = "[" & grdCodDesc.Columns(1).DataField & "] Like '*" & ActiveBar1.Tools("miOpText").Text & "*'"
        datMaster.Recordset.FindFirst gsCriteria
        If datMaster.Recordset.NoMatch Then
          gsTitle = "Pesquisa de Texto"
          gsMsg = "Texto n�o encontrado."
          gnStyle = vbOKOnly + vbInformation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          Exit Sub
        End If
      End If
    '
    Case "miOpSearchNext"
      If Len(gsCriteria) > 0 And Len(Trim(ActiveBar1.Tools("miOpText").Text)) > 0 Then
        datMaster.Recordset.FindNext gsCriteria
        If datMaster.Recordset.NoMatch Then
          gsTitle = "Pesquisa de Texto"
          gsMsg = "Nenhuma outra ocorr�ncia encontrada."
          gnStyle = vbOKOnly + vbInformation
          gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          If Not IsNull(gvBM) Then
            datMaster.Recordset.Bookmark = gvBM
          End If
          Exit Sub
        End If
      End If
    '
    '05/05/2005 - Daniel
    'Projeto: Melhorias para o Centro de Custo
    '
    'Adicionado para o Centro de Custo rotinas de
    'ativa��o e desativa��o de centro(s)
    Case "miOpDesativarCentro"
      Call DesativaCentroCusto
    Case "miOpAtivarCentros"
      Call AtivarCentroInativo
    '19/07/2005 - Daniel
    'Adicionado rotinas: Reativa��o de Centro de Custo
    'Individualmente e RefreshLinhasDaGrid
    Case "miOpReativarCentroIndividualmente"
      frmReativacaoCentroIndividualmente.Show
    Case "miOpRefresh"
      Call RefreshLinhasDaGrid
    '---------------------------------------------
    
  End Select

  On Error GoTo 0
  
  Exit Sub
  
ErrExit:
  Select Case Err.Number
    Case 3022  ' Duplicate Index
      MsgBox "Registro j� existe no cadastro.", vbExclamation
    Case 3314  ' Campo Obrigat�rio faltando
'      Dim nControl As Integer
      MsgBox "Campo obrigat�rio no registro n�o foi entrado.", vbExclamation
    Case 3315  ' Campo Nulo Invalido
'      Dim sDesc As String
'      sDesc = "3315 - " & Err.Description
      MsgBox "Campo n�o pode ficar sem valor.", vbExclamation
    Case Else
      MsgBox Err.Number & " - " & Err.Description, vbExclamation
  End Select
  On Error GoTo 0
  Exit Sub
  
End Sub

Private Sub chkMultiSelect_Click()
  If chkMultiSelect.Value = vbChecked Then
    grdCodDesc.SelectTypeRow = ssSelectionTypeMultiSelect
  Else
    grdCodDesc.SelectTypeRow = ssSelectionTypeSingleSelect
  End If
End Sub

Private Sub datMaster_Validate(Action As Integer, Save As Integer)
  If Action = vbDataActionAddNew Or Action = vbDataActionUpdate Then
    ActiveBar1.Tools("miOpSearch").Enabled = False
    ActiveBar1.Tools("miOpSearchNext").Enabled = False
    ActiveBar1.Tools("miOpText").Enabled = False
  Else
    ActiveBar1.Tools("miOpSearch").Enabled = True
    ActiveBar1.Tools("miOpSearchNext").Enabled = True
    ActiveBar1.Tools("miOpText").Enabled = True
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
  
  '----------------------------------------------------------------------------
  '21/07/2004 - mpdea
  'Case: Livraria Resultado (QS40590-987)
  'Bloqueia a edi��o da coluna c�digo
  grdCodDesc.Columns(0).Locked = CheckSerialCaseMod("QS40590-987")
  '----------------------------------------------------------------------------
  
  chkUCase.Enabled = gbPodeGravar
  chkMultiSelect.Enabled = gbPodeApagar
  Call ActiveBarLoadToolTips(Me)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If datMaster.Recordset.EditMode <> dbEditNone Then
    Screen.ActiveForm.datMaster.Recordset.CancelUpdate
  End If
  Screen.MousePointer = vbDefault
  Call StatusMsg("")
End Sub

Private Sub grdCodDesc_AfterUpdate(RtnDispErrMsg As Integer)
On Error GoTo Erro
  If RtnDispErrMsg = True Then
    RtnDispErrMsg = False      'Turn off SSDBGrid default message
    Call GridAfterUpdate
    If nColIndex = -1 Then Exit Sub
    grdCodDesc.Columns(nColIndex).CellStyleSet ("color"), grdCodDesc.Row
    Exit Sub
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro no m�todo AfterUpdate " & Err.Number & " " & Err.Description, vbInformation, "Aten��o"
End Sub

Private Sub grdCodDesc_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Variant, Cancel As Integer)
On Error GoTo Erro

  If chkUCase.Value = vbChecked Then
    grdCodDesc.Columns(ColIndex).Text = UCase(grdCodDesc.Columns(ColIndex).Text)
  End If
  
  '02/06/2008 - mpdea
  'Verifica c�digo m�ximo
  Dim str_code As String
  str_code = gsHandleNull(grdCodDesc.Columns(0).Text)
  If CLng(str_code) > gnMaxCod Then
    DisplayMsg "Valor de c�digo n�o deve ser maior do que " & gnMaxCod & "."
    grdCodDesc.Col = 0
    grdCodDesc.ActiveCell.SelStart = 1
    grdCodDesc.ActiveCell.SelLength = Len(CStr(gnMaxCod))
    Cancel = True
    Exit Sub
  End If
  
  'Verifica��es para Centros de Custos
  '
  '04/04/2005 - Daniel
  'Permitiremos alterar a descri��o do centro de custo 01 a partir da
  'vers�o 6.52.0.33
  '
'  If Me.Caption = "CENTROS DE CUSTO" Then
'    If (ColIndex = 0 And CStr(OldValue) = "1" And _
'      UCase(grdCodDesc.Columns("Nome ou Descri��o").Text) = "FORNECEDORES" And _
'      grdCodDesc.Columns("C�digo").Text <> "1") Or _
'      (ColIndex = 1 And UCase(CStr(OldValue)) = "FORNECEDORES" And _
'      grdCodDesc.Columns("C�digo").Text = "1" And _
'      UCase(grdCodDesc.Columns("Nome ou Descri��o").Text) <> "FORNECEDORES") Then
'
'      MsgBox "N�o � poss�vel alterar o Centro de Custo Fornecedores.", vbExclamation
'      grdCodDesc.Columns(ColIndex).Text = OldValue
'      grdCodDesc.Update
'      Cancel = True
'      Exit Sub
'    End If
'  End If
  
  Dim NF As Integer
  For NF = datMaster.Recordset.Fields.Count - 1 To 0 Step -1
    If datMaster.Recordset.Fields(NF).Name = "Data Altera��o" Then
      datMaster.Recordset.Fields("Data Altera��o") = Format(Date, "dd/mm/yyyy")
      '09/05/2005 - Daniel
      'Atualiza��o do campo Ativo da table Centros de Custo
      If Me.Caption = "CENTROS DE CUSTO" Then datMaster.Recordset.Fields("Ativo").Value = True
      '----------------------------------------------------
      Exit For
    End If
    
  Next NF
  
  Exit Sub
Erro:
  If Not IsNumeric(grdCodDesc.Columns(ColIndex).Text) And ColIndex = 0 Then
      MsgBox "Sempre digite n�mero na coluna C�digo", vbInformation, "Aten��o"
  ElseIf Not IsNumeric(grdCodDesc.Columns(0).Text) And ColIndex = 1 Then
      MsgBox "Sempre digite n�mero na coluna C�digo", vbInformation, "Aten��o"
  Else
      MsgBox "Aconteceu algo inexperado neste cadastro" & Err.Number & " " & Err.Description, vbInformation, "Aten��o"
  End If
End Sub

Private Sub grdCodDesc_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  DispPromptMsg = False
  If gbPodeApagar = False Then
    Beep
    Cancel = True
    Exit Sub
  End If
  'Verifica��es para Centros de Custos
  If Me.Caption = "CENTROS DE CUSTO" Then
'    C�digo comentado em 04/04/2005
'    If grdCodDesc.Columns("C�digo").Text = "1" And _
'      UCase(grdCodDesc.Columns("Nome ou Descri��o").Text) = "FORNECEDORES" Then
'      MsgBox "N�o � poss�vel excluir o Centro de Custo Fornecedores.", vbExclamation
'      Cancel = True
'      Exit Sub
'    End If

    '04/04/2005 - Daniel
    'Permitimos alterar a descri��o do centro de custo '1' mas n�o exclu�-la
    If grdCodDesc.Columns("C�digo").Text = "1" Then
      MsgBox "N�o � poss�vel excluir o Centro de Custo '1'.", vbExclamation
      Cancel = True
      Exit Sub
    End If

  End If
  If Len(Trim(grdCodDesc.ActiveCell.Text)) = 0 Then
    If bGridBeforeDelete(datMaster.Recordset) = True Then
      Call StatusMsg("Sele��o de itens apagada.")
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
On Error GoTo Erro

  Dim sCod As String
  If Not gbPodeGravar Then
    Beep
    Cancel = True
    Exit Sub
  End If
  sCod = gsHandleNull(grdCodDesc.Columns(0).Text)
  If Not IsNumeric(sCod) Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor de c�digo n�o � num�rico."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    grdCodDesc.Col = 0
    grdCodDesc.ActiveCell.SelStart = 1
    grdCodDesc.ActiveCell.SelLength = Len(CStr(gnMaxCod))
    'grdCodDesc.CancelUpdate
    Cancel = True
    Exit Sub
  End If
  
  '08/11/07 - Celso
  'Verifica se existe duplicidade e informa ao usu�rio
  
  Dim sNome As String
  Dim sNomeGrid As String
  Dim nRow As Integer
  Dim bm As Variant
  
  sNome = grdCodDesc.Columns(1).Text
    
  For nRow = 0 To grdCodDesc.Rows - 1
     bm = grdCodDesc.RowBookmark(nRow)
     sNomeGrid = grdCodDesc.Columns(1).CellText(bm)
     If UCase(sNome) = UCase(sNomeGrid) Then
        If MsgBox("Nome j� existe! Deseja gravar mesmo assim?", vbQuestion + vbYesNo, "Aten��o") = vbNo Then
           grdCodDesc.CancelUpdate
           grdCodDesc.AddNew
           grdCodDesc.Col = 1
           Cancel = True
           Exit Sub
        End If
        Exit For
     End If
  Next

  '----------------------------------------------------------------------------
  '21/07/2004 - mpdea
  'Obt�m o pr�ximo c�digo dispon�vel caso n�o esteja preenchido
  '----------------------------------------------------------------------------
  If CLng(sCod) = 0 Then
    Call GetNewCode(Me, datMaster.Recordset, gnMaxCod)
  End If
  'Verifica se o c�digo � v�lido
  sCod = gsHandleNull(grdCodDesc.Columns(0).Text)
  '----------------------------------------------------------------------------
  
  If CLng(sCod) = 0 Then
    gsTitle = LoadResString(201)
    gsMsg = "Valor de c�digo n�o deve ser zero."
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    grdCodDesc.Col = 0
    grdCodDesc.ActiveCell.SelStart = 1
    grdCodDesc.ActiveCell.SelLength = Len(CStr(gnMaxCod))
    'grdCodDesc.CancelUpdate
    Cancel = True
    Exit Sub
  End If
  If Len(Trim(grdCodDesc.Columns(0).Text)) = 0 Or Len(Trim(grdCodDesc.Columns(1).Text)) = 0 Then
    gsTitle = LoadResString(201)
    
    Dim sNomeTela As String

    If Me.Caption = "CLASSES" Then
        sNomeTela = "classe"
    ElseIf Me.Caption = "SUBCLASSES" Then
        sNomeTela = "subclasse"
    ElseIf Me.Caption = "CORES" Then
        sNomeTela = "cor"
    ElseIf Me.Caption = "TAMANHOS" Then
        sNomeTela = "tamanho"
    End If
    gsMsg = "Digite o nome da " & sNomeTela & "." & vbCrLf & vbCrLf & "Por�m, caso n�o deseje seguir com o cadastro, feche a tela ou selecione a linha e aperte no �cone apagar."
    
    gnStyle = vbOKOnly + vbInformation
    gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
    grdCodDesc.Col = 0
    Cancel = True
  End If
  
  Exit Sub
Erro:
  MsgBox "Erro no m�todo BeforeUpdate " & Err.Number & " " & Err.Description, vbInformation, "Aten��o"
  
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
    sSql = Trim(Left(sSql, nPos - 1))
  End If
  datMaster.RecordSource = sSql & sSQLOrdem
  datMaster.Refresh
  Call StatusMsg("")
End Sub

'10/06/2008 - mpdea
'Corrigido exibi��o da coluna "Lucro M�nimo Permitido"
Private Sub grdCodDesc_InitColumnProps()
  grdCodDesc.Columns(0).FieldLen = Len(CStr(gnMaxCod))
  grdCodDesc.Columns(1).FieldLen = datMaster.Recordset.Fields(1).Size
  grdCodDesc.Columns(0).Width = 2099
  '19/10/2007 - Anderson
  'Implementa��o do campo Lucro M�nimo Permitido conforme solicita��o da Agrotama
  'grdCodDesc.Columns(1).Width = 6240
  If g_bolLucroMinimoClasse And Me.Caption = "CLASSES" Then
    grdCodDesc.Columns(1).Width = 4395
    grdCodDesc.Columns("Lucro M�nimo Permitido").Visible = True
  Else
    grdCodDesc.Columns(1).Width = 6240
    On Error Resume Next
    grdCodDesc.Columns("Lucro M�nimo Permitido").Visible = False
  End If
End Sub

Private Sub grdCodDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If grdCodDesc.Col = 0 Then
    If KeyCode = vbKeyF5 Then
'      Call GetNewCode(Me, datMaster.Recordset, gnMaxCod)
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
  datMaster.DatabaseName = gsQuickDBFileName
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
        gsMsg = "Tentativa de inserir um item de c�digo j� existente."
        
        '----------------------------------------------------------------------------
        '21/07/2004 - mpdea
        'Se estiver bloqueado a edi��o da coluna c�digo ent�o gera novo c�digo
        If grdCodDesc.Columns(0).Locked Then
          Call GetNewCode(Me, datMaster.Recordset, gnMaxCod)
        End If
        '----------------------------------------------------------------------------
        
      Case 3163  ' Field Overflow
        gsMsg = "Overflow"
      Case 3314, 3058  ' Invalid Use of Null to Required Field
        gsMsg = "Campo vazio inv�lido."
      Case 3421
        gsMsg = "Erro de convers�o de dados."
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
  gsMsg = "Apagar sele��o atual?"
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
  Dim nPos As Integer
  Dim sSql As String
  
  Dim lngTestNumber As Long
  
  '18/07/2005 - Daniel
  'Adicionado tratamento de erro
  On Error GoTo TratarErro
  
  Call StatusMsg("")
  
  nCod = 1
  
  sSql = rs.Name
  nPos = InStr(sSql, "ORDER BY")
  If nPos > 0 Then
    sSql = Trim(Left(sSql, nPos - 1))
  End If
  sSql = sSql & " ORDER BY " & rs.Fields(0).Name
  
  Set rsTemp = db.OpenRecordset(sSql, dbOpenSnapshot)
  With rsTemp
    If Not .EOF Then
      .MoveLast
      nCod = .Fields(0) + 1
      
      '19/07/2005 - Daniel
      'Se o maior c�digo do Centro estiver como inativo o sistema dever�
      'considerar tamb�m para n�o trazer valores errados na inser��o de
      'um novo c�digo para o Centro pois a sSql est� trazendo a seguinte
      'instru��o [SELECT C�digo, Nome, [Data Altera��o], Ativo FROM [Centros de Custo]
      '           WHERE Ativo = TRUE ORDER BY C�digo]
      If Me.Caption = "CENTROS DE CUSTO" Then Call ValidarUltimoCodigoCentro(nCod)
      
      '-----------------------------------------------------
      '18/07/2005 - Daniel
      'Adicionado tratamento para a empresa Irm�os Ambr�zio
      'Aberto permiss�o para cadastrar at� 254 c�digos para
      'Classifica��o Fiscal n�o limitando at� '99'
      If Me.Caption = "CLASSIFICA��O FISCAL" Then
        If CheckSerialCaseMod("QS35288-570", "QS36824-735") Then
          'N�o poder� ser superior a 254
          If nCod > 254 Then
            gsTitle = LoadResString(201)
            gsMsg = "Nenhum C�digo livre dispon�vel para o intervalo."
            gnStyle = vbOKOnly + vbExclamation
            gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
          Else
            F.grdCodDesc.Columns(0).Text = nCod
          End If
          'Fechamos o recordset
          rsTemp.Close
          Set rsTemp = Nothing
          'Sa�mos da rotina
          Exit Sub
        End If
      End If
      '-----------------------------------------------------
      
      If nCod > nMaxCod Then
'        For nCod = 1 To nMaxCod
'          DoEvents
'          .FindFirst .Fields(0).Name & " = " & nCod
'          If .NoMatch Then
'            F.grdCodDesc.Columns(0).Text = Format(nCod, String(Len(CStr(nMaxCod)), "0"))
'            .Close
'            Set rsTemp = Nothing
'            Exit Sub
'          End If
'        Next nCod
        
        '----------------------------------------------------------------------------
        '21/07/2004 - mpdea
        'Otimizado (e muito he he) a busca do pr�ximo c�digo livre ;-)
        .MoveFirst
        Do Until .EOF
          If CLng(.Fields(0).Value) > lngTestNumber + 1 Then
            F.grdCodDesc.Columns(0).Text = Format(lngTestNumber + 1, String(Len(CStr(nMaxCod)), "0"))
            .Close
            Set rsTemp = Nothing
            Exit Sub
          Else
            lngTestNumber = CLng(.Fields(0).Value)
          End If
          .MoveNext
        Loop
        '----------------------------------------------------------------------------
        
        gsTitle = LoadResString(201)
        gsMsg = "Nenhum C�digo livre dispon�vel para o intervalo."
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
  
  Exit Sub

TratarErro:
  MsgBox "Erro na Private <GetNewCode>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Sub

Private Sub DesativaCentroCusto()
  '05/05/2005 - Daniel
  'Projeto: Melhorias para o Centro de Custo
  Dim rstCentro As Recordset
  Dim varBook   As Variant
  Dim intAuxi   As Integer
  Dim strSQL    As String
  
  If MsgBox("Deseja inativar o(s) Centro(s) selecionado(s)?", vbQuestion + vbYesNo, "Aten��o") = vbNo Then Exit Sub
  
  For intAuxi = 0 To (grdCodDesc.SelBookmarks.Count - 1)
    varBook = grdCodDesc.SelBookmarks(intAuxi)
    grdCodDesc.Bookmark = varBook

    strSQL = ""
    strSQL = "SELECT Ativo, [Data Altera��o] FROM [Centros de Custo] WHERE C�digo = " & CInt(grdCodDesc.Columns("C�digo").Text)

    Set rstCentro = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    With rstCentro
      If Not (.BOF And .EOF) Then
        .MoveFirst
        .Edit
        .Fields("Ativo").Value = False
        .Fields("Data Altera��o").Value = Format(Date, "dd/mm/yyyy")
        .Update
      End If
      .Close
    End With
    
    Set rstCentro = Nothing

  Next intAuxi

  Call RefreshLinhasDaGrid

'  '----------------------------------
'  'Refresh na grid
'  '----------------------------------
'  Set datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o], Ativo FROM [Centros de Custo] WHERE Ativo = TRUE ORDER BY C�digo", dbOpenDynaset)
'  If Not datMaster.Recordset.EOF Then
'    datMaster.Recordset.MoveLast
'    datMaster.Recordset.MoveFirst
'  End If
'  datMaster.Refresh
'  Set grdCodDesc.DataSource = datMaster
  

End Sub

Private Sub AtivarCentroInativo()
  '05/05/2005 - Daniel
  'Projeto: Melhorias para o Centro de Custo
  Dim rstCentro As Recordset
  Dim strSQL    As String
  
  If MsgBox("Deseja ativar todos Centros inativados?", vbQuestion + vbYesNo, "Aten��o") = vbNo Then Exit Sub
  
  strSQL = "SELECT Ativo, [Data Altera��o] FROM [Centros de Custo] WHERE NOT Ativo "
  
  Set rstCentro = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstCentro
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        .Edit
        .Fields("Ativo").Value = True
        .Fields("Data Altera��o").Value = Format(Date, "dd/mm/yyyy")
        .Update
        
       .MoveNext
      Loop
    End If
    .Close
  End With
  
  Set rstCentro = Nothing
  
  Call RefreshLinhasDaGrid
  
'  '----------------------------------
'  'Refresh na grid
'  '----------------------------------
'  Set datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o], Ativo FROM [Centros de Custo] WHERE Ativo = TRUE ORDER BY C�digo", dbOpenDynaset)
'  If Not datMaster.Recordset.EOF Then
'    datMaster.Recordset.MoveLast
'    datMaster.Recordset.MoveFirst
'  End If
'  datMaster.Refresh
'  Set grdCodDesc.DataSource = datMaster

End Sub

Private Sub RefreshLinhasDaGrid()
  '19/07/2005 - Daniel
  On Error GoTo Erro
  
  Set datMaster.Recordset = db.OpenRecordset("SELECT C�digo, Nome, [Data Altera��o], Ativo FROM [Centros de Custo] WHERE Ativo = TRUE ORDER BY C�digo", dbOpenDynaset)
  If Not datMaster.Recordset.EOF Then
    datMaster.Recordset.MoveLast
    datMaster.Recordset.MoveFirst
  End If
  datMaster.Refresh
  Set grdCodDesc.DataSource = datMaster
  
  Exit Sub

Erro:
  MsgBox "Erro em Private <RefreshLinhasDaGrid>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  
End Sub

Private Sub ValidarUltimoCodigoCentro(ByRef intUltimo As Long)
  '19/07/2005 - Daniel
  'Se o maior c�digo do Centro estiver como inativo o sistema dever�
  'considerar tamb�m para n�o trazer valores errados na inser��o de
  'um novo c�digo para o Centro
  Dim rstCentro As Recordset
  
  On Error GoTo TratarErro
  
  Set rstCentro = db.OpenRecordset("SELECT MAX(C�digo) AS Maior FROM [Centros de Custo]", dbOpenSnapshot)
  
  With rstCentro
    If Not (.BOF And .EOF) Then
      .MoveFirst
      If intUltimo <= .Fields("Maior").Value Then intUltimo = .Fields("Maior").Value + 1
    End If
    .Close
  End With
  
  Set rstCentro = Nothing
  
  Exit Sub

TratarErro:
  MsgBox "Erro na Private <ValidarUltimoCodigoCentro>" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
    
End Sub

