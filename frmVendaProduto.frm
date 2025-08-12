VERSION 5.00
Begin VB.Form frmVendaProduto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa Produto para Venda"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   Icon            =   "frmVendaProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPaintConsProduct 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   7515
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1920
      Width           =   7575
      Begin VB.TextBox txtConsProductName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   7215
      End
      Begin VB.ListBox lstConsProduct 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "frmVendaProduto.frx":000C
         Left            =   120
         List            =   "frmVendaProduto.frx":0013
         TabIndex        =   1
         Top             =   1440
         Width           =   7215
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   7215
      End
      Begin VB.Label lblConsProductValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   4
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label lblConsProductEstoque 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ESTOQUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   3495
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PREÇO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   8
         Left            =   3840
         TabIndex        =   13
         Top             =   3000
         Width           =   3495
      End
   End
   Begin VB.PictureBox picPaintConsProductGE 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   7515
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   7575
      Begin VB.ListBox lstConsProductGE 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         ItemData        =   "frmVendaProduto.frx":0052
         Left            =   120
         List            =   "frmVendaProduto.frx":0059
         TabIndex        =   2
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PREÇO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   9
         Left            =   3840
         TabIndex        =   11
         Top             =   3000
         Width           =   3495
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ESTOQUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   3495
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RELAÇÃO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   435
         Index           =   11
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   7215
      End
      Begin VB.Label lblConsProductValueGE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   8
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label lblConsProductEstoqueGE 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   3495
      End
   End
   Begin VB.Label lblDescricao 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12345678901234567890123456789 Descrição"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmVendaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_rstConsProduct   As Recordset
Private m_rstConsProductGE As Recordset
Private m_strTabelaPreco As String

Public Sub ShowMsgInDisplay(ByVal sMsg As String, _
  Optional ByVal bStyleDescription As Boolean = False, _
  Optional ByVal bVisible As Boolean = True)
  
  With lblDescricao
    If bStyleDescription Then
      .FontName = "Courier New"
      .Alignment = vbLeftJustify
    Else
      .FontName = "Verdana"
      .Alignment = vbCenter
    End If
    .Caption = sMsg
    .Visible = bVisible
    .Refresh
  End With
End Sub

Private Sub Form_Load()
  Call StartConsProduct
  m_strTabelaPreco = frmVendaRap2.Combo_Preço.Text
End Sub

Private Sub lstConsProduct_KeyDown(KeyCode As Integer, Shift As Integer)
  
End Sub

Private Sub txtConsProductName_GotFocus()
  Call SelectAllText(txtConsProductName)
End Sub

Private Sub txtConsProductName_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown
      lstConsProduct.SetFocus
      SendKeys "{Down}"
    Case vbKeyUp
      lstConsProduct.SetFocus
      SendKeys "{Up}"
  End Select
End Sub

Private Sub txtConsProductName_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And txtConsProductName.Text <> "" Then
    KeyAscii = 0
    Call ListConsProductName(txtConsProductName.Text)
  ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    Unload Me
  Else
    KeyAscii = gnTypeValidKey(KeyAscii)
  End If
End Sub

Private Sub ListConsProductName(ByVal strSearch As String)
  Dim strSQL As String
  Dim strAux As String
  Dim intLenCodigo As Integer
  Dim intLenNome As Integer
  
  Call ShowMsgInDisplay(vbCrLf & "AGUARDE...")
  lstConsProduct.Clear
  lblConsProductEstoque.Caption = "0"
  lblConsProductValue.Caption = Format(0, FORMAT_VALUE)
  DoEvents
  
  strSQL = "SELECT * FROM Produtos WHERE Código <> '0' AND Nome LIKE '" & _
    strSearch & "*' ORDER BY Nome;"
  
  Set m_rstConsProduct = db.OpenRecordset(strSQL)
  
  With m_rstConsProduct
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      'Carrega ListBox com o resultado -> !limitar quantidade de registros
      Do Until .EOF
        intLenCodigo = Len(.Fields("Código").Value & "")
        intLenNome = Len(.Fields("Nome").Value & "")
        If intLenCodigo + intLenNome + 1 > 57 Then
          strAux = Left(.Fields("Nome").Value & "", 57 - 4 - intLenCodigo) & "... " & .Fields("Código").Value & ""
        Else
          strAux = Left(.Fields("Nome").Value & "" & Space(57), 57 - 1 - intLenCodigo) & " " & .Fields("Código").Value & ""
        End If
        lstConsProduct.AddItem strAux
        lstConsProduct.ItemData(lstConsProduct.NewIndex) = .AbsolutePosition
        .MoveNext
      Loop
      .MoveFirst
      .MovePrevious
      Call ShowMsgInDisplay(vbCrLf & "CONSULTA PRODUTO")
    Else
      Call ShowMsgInDisplay("NENHUM PRODUTO ENCONTRADO COM ESTES CRITÉRIOS")
    End If
  End With
    
End Sub

Private Sub StartConsProduct()
  Call ShowMsgInDisplay(vbCrLf & "CONSULTA PRODUTO")
  lstConsProduct.Clear
  txtConsProductName = ""
  Call DeActiveCons(True)
  lblConsProductValue.Caption = Format(0, "#0.00")
  lblConsProductEstoque.Caption = "0"
  Call SelectAllText(txtConsProductName, True)
  DoEvents
  Call lstConsProduct_Click
End Sub

Private Sub StartConsProductGE()
  Dim strSQL As String
  Dim strAux As String
  Dim strTipoProduto As String
  Dim intLenEstoque As Integer
  Dim intLenDescricao As Integer
  Dim blnEstoque As Boolean
  Dim dblEstoque As Double
  Dim strEstoque As String
  Dim intTam As Integer
  Dim intCor As Integer
  
  On Error GoTo ErrHandler
  
  If m_rstConsProduct Is Nothing Then
    Call ShowMsgInDisplay(vbCrLf & "REALIZE UMA CONSULTA")
    Exit Sub
  ElseIf m_rstConsProduct.AbsolutePosition = -1 Then
    Call ShowMsgInDisplay(vbCrLf & "ESCOLHA O PRODUTO")
    Exit Sub
  End If
  
  blnEstoque = m_rstConsProduct.Fields("Estoque").Value
  
  strTipoProduto = UCase(m_rstConsProduct.Fields("Tipo").Value)
  Select Case strTipoProduto
    Case "N"
      Call ShowMsgInDisplay(vbCrLf & "PRODUTO NORMAL")
      Exit Sub
    Case "E"
      strSQL = "SELECT * FROM Produtos INNER JOIN Edições ON " & _
        "Produtos.Código = Edições.Produto WHERE Produtos.Código = '" & _
        m_rstConsProduct.Fields("Código").Value & "' ORDER BY Edições.Código;"
      lblTitle(11).Caption = "EDIÇÕES"
    Case "G"
      strSQL = "SELECT * FROM Produtos INNER JOIN [Códigos da Grade] ON " & _
        "Produtos.Código = [Códigos da Grade].[Código Original] WHERE Produtos.Código = '" & _
        m_rstConsProduct.Fields("Código").Value & "' ORDER BY [Códigos da Grade].Código;"
      lblTitle(11).Caption = "GRADES"
    Case Else
      Call ShowMsgInDisplay(vbCrLf & "TIPO DESCONHECIDO")
      Exit Sub
  End Select
  
  Call ShowMsgInDisplay(vbCrLf & "AGUARDE...")
  lstConsProductGE.Clear
    
  Set m_rstConsProductGE = db.OpenRecordset(strSQL)
  
  With m_rstConsProductGE
    If .RecordCount > 0 Then
      .MoveLast
      .MoveFirst
      'Carrega ListBox com o resultado -> !limitar quantidade de registros
      Do Until .EOF
        'Código
        If strTipoProduto = "E" Then 'Edição
          strAux = Format(.Fields("Edições.Código").Value, String(5, "0"))
        Else 'Grade
          strAux = Right(.Fields("Códigos da Grade.Código").Value, 6)
        End If
        'Estoque
        If blnEstoque Then
          If m_blnCheckStockProduct(gnCodFilial, m_rstConsProduct.Fields("Código").Value & strAux, dblEstoque) Then
            strEstoque = CStr(dblEstoque)
          Else
            strEstoque = "?"
          End If
        Else
          strEstoque = "NC"
        End If
        'Nome
        If strTipoProduto = "E" Then 'Edição
          strAux = strAux & " - " & .Fields("Edições.Nome").Value
        Else 'Grade
          strAux = .Fields("Códigos da Grade.Código").Value
          intTam = CInt(Mid(strAux, Len(strAux) - 5, 3))
          intCor = CInt(Right(strAux, 3))
          strAux = Format(intTam, String(3, "0")) & "-" & gsGetNameTamanho(intTam) & " " & _
            Format(intCor, String(3, "0")) & "-" & gsGetNameCor(intCor)
        End If
        'Formatação
        intLenEstoque = Len(strEstoque)
        intLenDescricao = Len(strAux)
        If intLenEstoque + intLenDescricao + 1 > 57 Then
          strAux = Left(strAux, 57 - 4 - intLenEstoque) & "... " & strEstoque
        Else
          strAux = Left(strAux & Space(57), 57 - 1 - intLenEstoque) & " " & strEstoque
        End If
        'Adiciona a linha
        lstConsProductGE.AddItem strAux
        lstConsProductGE.ItemData(lstConsProductGE.NewIndex) = .AbsolutePosition
        .MoveNext
      Loop
      .MoveFirst
      .MovePrevious
      Call ShowMsgInDisplay(Trim(m_rstConsProduct.Fields("Nome").Value))
    Else
      If strTipoProduto = "E" Then 'Edição
        Call ShowMsgInDisplay("NENHUMA EDIÇÃO ENCONTRADA PARA ESTE PRODUTO")
      Else 'Grade
        Call ShowMsgInDisplay("NENHUMA GRADE ENCONTRADA PARA ESTE PRODUTO")
      End If
      Exit Sub
    End If
  End With
  
  lblConsProductEstoqueGE.Caption = lblConsProductEstoque.Caption
  lblConsProductValueGE.Caption = lblConsProductValue.Caption
  
  Call DeActiveCons(False)
  picPaintConsProductGE.Visible = True
  lstConsProductGE.SetFocus

  Exit Sub

ErrHandler:
  lblConsProductEstoque.Caption = "0"
  lblConsProductValue.Caption = Format(0, FORMAT_VALUE)
  Select Case Err.Number
    Case 3167
      Call ShowMsgInDisplay(vbCrLf & "PRODUTO FOI EXCLUÍDO")
    Case Else
      Call ShowMsgInDisplay(Err.Description)
  End Select
  
End Sub

Private Sub DeActiveCons(ByVal blnEnabled As Boolean)
  txtConsProductName.Enabled = blnEnabled
  lstConsProduct.Enabled = blnEnabled
  picPaintConsProduct.Visible = blnEnabled
  If blnEnabled Then
    Call SelectAllText(txtConsProductName, True)
  End If
End Sub

Private Sub ConcludeConsProduct(ByVal blnOk As Boolean)
  Dim intCancel As Integer
  
  On Error GoTo ErrHandler
  
  If blnOk Then
    'Validação
    If m_rstConsProduct Is Nothing Then
      Call ShowMsgInDisplay(vbCrLf & "REALIZE UMA CONSULTA")
      Exit Sub
    ElseIf m_rstConsProduct.AbsolutePosition = -1 Then
      Call ShowMsgInDisplay(vbCrLf & "ESCOLHA O PRODUTO")
      Exit Sub
    End If
    
    With frmVendaRap2
      'Insere o item
      .Grade1.Columns(0).Text = m_rstConsProduct.Fields("Código").Value
      .Grade1.Columns(1).Text = "1"
      'Atualiza grid
      .Grade1_BeforeColUpdate 0, "", intCancel
      If intCancel = -1 Then Exit Sub
      'Calcula totais
      .Calcula_Linha
      .Recalcula
      'Move para a próxima linha
      .Grade1.MoveNext
      .Grade1.DoClick
    End With
    
  End If
  
  Unload Me
  frmVendaRap2.Grade1.SetFocus
    
  Exit Sub
  
ErrHandler:
  Select Case Err.Number
    Case 3167
      Call ShowMsgInDisplay(vbCrLf & "PRODUTO FOI EXCLUÍDO")
    Case Else
      Call ShowMsgInDisplay(Err.Description)
  End Select

End Sub

Private Sub ConcludeConsProductGE(ByVal blnOk As Boolean)
  Dim strAux As String
  
  If blnOk Then
    'Validação
    If m_rstConsProductGE Is Nothing Then
      Call ShowMsgInDisplay(vbCrLf & "REALIZE UMA CONSULTA")
      Exit Sub
    ElseIf m_rstConsProductGE.AbsolutePosition = -1 Then
      Call ShowMsgInDisplay(vbCrLf & "ESCOLHA O PRODUTO")
      Exit Sub
    End If
    'Código
    If m_rstConsProduct("Tipo") = "E" Then 'Edição
      strAux = m_rstConsProductGE("Produtos.Código") & Format(m_rstConsProductGE("Edições.Código"), String(5, "0"))
    Else 'Grade
      strAux = m_rstConsProductGE("Códigos da Grade.Código")
    End If
    'Código completo
    frmVendaRap2.Grade1.AddItem strAux
    frmVendaRap2.Grade1.SetFocus
  Else
    Call ShowMsgInDisplay(vbCrLf & "CONSULTA PRODUTO")
    Call DeActiveCons(True)
    Call lstConsProduct_Click
    lstConsProduct.SetFocus
  End If
  picPaintConsProductGE.Visible = False
    
End Sub

Private Sub lstConsProduct_Click()
  Dim lngPosAtual   As Long
  Dim lngPosMover   As Long
  Dim intMoeda      As Integer
  Dim curAuxPreco   As Currency
  Dim curAuxCotacao As Currency
  Dim dblEstoque    As Double
  
  On Error GoTo ErrHandler
  
  If lstConsProduct.ListIndex = -1 Then Exit Sub
  
  'Posiciona o registro
  lngPosAtual = m_rstConsProduct.AbsolutePosition
  lngPosMover = lstConsProduct.ItemData(lstConsProduct.ListIndex)
  m_rstConsProduct.Move lngPosMover - lngPosAtual
  
  'Nome
  Call ShowMsgInDisplay(UCase(Trim(m_rstConsProduct.Fields("Nome").Value)))
    
  'Estoque
  If m_rstConsProduct.Fields("Estoque").Value Then
    Select Case UCase(m_rstConsProduct.Fields("Tipo").Value)
      Case "N" 'Normal
        If m_blnCheckStockProduct(gnCodFilial, m_rstConsProduct.Fields("Código").Value, dblEstoque) Then
          lblConsProductEstoque.Caption = dblEstoque
        Else
          lblConsProductEstoque.Caption = "?"
        End If
        lblConsProductEstoque.Tag = ""
      Case "E" 'Edição
        lblConsProductEstoque.Caption = "VER EDIÇÃO"
        lblConsProductEstoque.Tag = "EG"
      Case "G" 'Grade
        lblConsProductEstoque.Caption = "VER GRADE"
        lblConsProductEstoque.Tag = "EG"
      Case Else
        lblConsProductEstoque.Caption = "0"
        lblConsProductEstoque.Tag = ""
    End Select
  Else
    lblConsProductEstoque.Caption = "NC"
    Select Case UCase(m_rstConsProduct.Fields("Tipo").Value)
      Case "E", "G" 'Edição, Grade
        lblConsProductEstoque.Tag = "EG"
      Case Else
        lblConsProductEstoque.Tag = ""
    End Select
  End If
  
'----------------------------------------------------------------------------------
'31/12/2002 - mpdea
'Função gcGetPrecoProduto traz cotação agora
'
  'Preço
  intMoeda = rsProdutos.Fields("Moeda").Value
  curAuxCotacao = 1
  curAuxPreco = gcGetPrecoProduto(m_rstConsProduct.Fields("Código").Value, m_strTabelaPreco)
'
'  If intMoeda <> 1 Then
'    rsCotacao.Index = "Moeda"
'    rsCotacao.Seek "<", intMoeda, CDate("01/01/2099")
'    If Not rsCotacao.NoMatch Then
'      If rsCotacao("Moeda") = intMoeda Then
'        curAuxCotacao = rsCotacao("Cotação")
'      End If
'    End If
'  End If
'  curAuxPreco = curAuxPreco * curAuxCotacao
  lblConsProductValue.Caption = Format(curAuxPreco, FORMAT_VALUE)
'----------------------------------------------------------------------------------
  
  Exit Sub
  
ErrHandler:
  lblConsProductEstoque.Caption = "0"
  lblConsProductValue.Caption = Format(0, FORMAT_VALUE)
  Select Case Err.Number
    Case 3167
      Call ShowMsgInDisplay(vbCrLf & "PRODUTO FOI EXCLUÍDO")
    Case Else
      Call ShowMsgInDisplay(Err.Description)
  End Select
  
End Sub

Private Sub lstConsProduct_DblClick()
  Call lstConsProduct_KeyPress(vbKeyReturn)
End Sub

Private Sub lstConsProduct_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    'Verifica acesso a Grade/Edição...
    If lblConsProductEstoque.Tag = "EG" Then
      KeyAscii = 0
      Call StartConsProductGE
    Else
      KeyAscii = 0
      Call ConcludeConsProduct(True)
    End If
  ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    Call ConcludeConsProduct(False)
  End If
End Sub

Private Sub lstConsProductGE_Click()
  Dim nPosAtual   As Long
  Dim nPosMover   As Long
  
  If lstConsProductGE.ListIndex = -1 Then Exit Sub
  
  'Posiciona o registro
  nPosAtual = m_rstConsProductGE.AbsolutePosition
  nPosMover = lstConsProductGE.ItemData(lstConsProductGE.ListIndex)
  m_rstConsProductGE.Move nPosMover - nPosAtual
  
End Sub

Private Sub lstConsProductGE_DblClick()
  Call lstConsProductGE_KeyPress(vbKeyReturn)
End Sub

Private Sub lstConsProductGE_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    Call ConcludeConsProductGE(True)
  ElseIf KeyAscii = vbKeyEscape Then
    KeyAscii = 0
    Call ConcludeConsProductGE(False)
  End If
End Sub

'Verifica o estoque atual para o produto com o seu código completo
Private Function m_blnCheckStockProduct(ByVal nFilial As Integer, _
  ByVal sCodeComplete As String, ByRef nEstoque As Double) As Boolean
  
  Dim nErro As Integer
  Dim nTamanho As Integer
  Dim nCor As Integer
  Dim nEdicao As Long
  Dim sCodPrincipal As String
  
  If sCodeComplete <> "" Then
    Call Acha_Produto(sCodeComplete, sCodPrincipal, nTamanho, nCor, nEdicao, 0, nErro)
    If nErro = 0 Then
      nEstoque = Acha_Estoque(gnCodFilial, sCodPrincipal, nTamanho, nCor, nEdicao, nErro)
      m_blnCheckStockProduct = (nErro = 0)
    End If
  End If

End Function
