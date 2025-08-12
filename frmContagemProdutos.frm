VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmContagemProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de contagem dos produtos"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "frmContagemProdutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLeitor 
      Caption         =   "Utilizar leitor óptico"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar Cabeçalho"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Data datProdutos 
      Caption         =   "datProdutos"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' ORDER BY Nome"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox txtNomeProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox txtQtde 
      Height          =   285
      Left            =   6480
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Pressione ENTER para confirmar."
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdExecutar 
      Caption         =   "&Finalizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin SSDataWidgets_B.SSDBGrid grdLista 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7815
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   3
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3334
      Columns(0).Caption=   "Codigo"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6694
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2725
      Columns(2).Caption=   "Qtde"
      Columns(2).Name =   "Qtde"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   13785
      _ExtentY        =   6376
      _StockProps     =   79
      Caption         =   "Listagem"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBCombo cboProdutos 
      Bindings        =   "frmContagemProdutos.frx":058A
      DataSource      =   "datProdutos"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
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
      Columns(0).Width=   3916
      Columns(0).Caption=   "Codigo"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Código"
      Columns(0).FieldLen=   256
      Columns(1).Width=   7620
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      _ExtentX        =   3413
      _ExtentY        =   503
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label Label1 
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Qtde Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   4500
      Width           =   975
   End
End
Attribute VB_Name = "frmContagemProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private rsProdutos As Recordset
  Private rsCodGrade As Recordset
  Public lngCodigoCliente As Long    'Var criada em 18/01/2005

Private Sub cboProdutos_CloseUp()
  cboProdutos.Text = cboProdutos.Columns(0).Text
  cboProdutos_LostFocus
End Sub

'02/07/2004 - mpdea
'Incluído tratamento quando a combo está aberta
Private Sub cboProdutos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If cboProdutos.DroppedDown Then
      Call cboProdutos_CloseUp
    Else
      '04/03/2005 - Daniel
      'Otimizado para utilização do leitor óptico
      'Antes era chamada a PreencheGrid direto
      If chkLeitor.Value = vbUnchecked Then
        PreencheGrid
      Else
        cboProdutos_LostFocus
      End If
    End If
  End If
End Sub

Private Sub cboProdutos_LostFocus()
  If Len(Trim(cboProdutos.Text)) <= 0 Then Exit Sub
  
'  Set rsProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & cboProdutos.Text & "'", dbOpenSnapshot)
'
'  With rsProdutos
'    If (.BOF And .EOF) Then
'      Set rsCodGrade = db.OpenRecordset("SELECT * FROM [Códigos da Grade] WHERE Código = '" & cboProdutos.Text & "'", dbOpenSnapshot)
'
'      With rsCodGrade
'        If (.BOF And .EOF) Then
'          MsgBox "Produto inexistente !", vbCritical, "Quick Store"
'
'          '02/07/2004 - mpdea
'          'Seta o foco para a combo de produtos
'          cboProdutos.SetFocus
'
'          Exit Sub
'        End If
'      End With
'    End If
'
'    txtNomeProduto.Text = .Fields("Nome")
'  End With

  '17/01/2005 - Daniel
  '
  'Comentado antigo código acima para a partir de 01/2005 atender
  'a realidade de produtos com grade
  Dim Cód         As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor     As Integer
  Dim Aux_Edição  As Long
  Dim Aux_Produto As String
  Dim Aux_Tipo    As Integer
  Dim Aux_Erro    As Integer
  Dim Cancel      As Integer
  
  Cód = Trim(CStr(cboProdutos.Text))
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  
  Call Acha_Produto(Cód, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Tipo, Aux_Erro)
  If Aux_Erro <> 0 Then
    Cancel = True
    If Aux_Erro = 1 Then
      DisplayMsg "Produto não existe."
    ElseIf Aux_Erro = 2 Then
      DisplayMsg "Produto com grade, digite tamanho e cor."
    ElseIf Aux_Erro = 3 Then
      DisplayMsg "Produto com edição, digite a edição também."
    End If
    cboProdutos.SetFocus
    Exit Sub
  End If
  
  '-------------------------------------------------------------
  'Caso não ocorra um Exit Sub dá continuidade ao código abaixo
  '-------------------------------------------------------------
  If Aux_Tamanho = 0 And Aux_Cor = 0 Then 'Não usa Grade
  
    Set rsProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & Aux_Produto & "'", dbOpenSnapshot)
  
    With rsProdutos
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        txtNomeProduto.Text = .Fields("Nome").Value & ""
        
      End If
      '.Close
    End With
    
    'Set rsProdutos = Nothing
  
  Else 'Usa Grade
  
    Set rsCodGrade = db.OpenRecordset("SELECT * FROM [Códigos da Grade] WHERE Código = '" & cboProdutos.Text & "'", dbOpenSnapshot)
  
    With rsCodGrade
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Set rsProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & .Fields("Código Original").Value & "'", dbOpenSnapshot)
        
        rsProdutos.MoveFirst
        txtNomeProduto.Text = rsProdutos.Fields("Nome").Value & ""
      Else
        MsgBox "Produto inexistente !", vbCritical, "Quick Store"
      End If
        
    End With
    
    'rsCodGrade.Close
    'rsProdutos.Close
    'Set rsCodGrade = Nothing
    'Set rsProdutos = Nothing
  
  End If
  
  '04/03/2005 - Daniel
  'Otimizado para utilização do leitor óptico
  If chkLeitor.Value = vbChecked Then txtQtde.SetFocus
  
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdExecutar_Click()
  Dim rstProdutos As Recordset
  Dim rstGrade As Recordset
  Dim rstPreco As Recordset
  Dim rstParamFilial As Recordset
  
  Dim strCodigoProduto As String
  
  Dim intX As Integer
  Dim lngTotal As Long
  Dim dblTotalMonetario As Double
  
  grdLista.Update
  grdLista.MoveFirst
  
  With frmManutencaoConsignacao.grdMovimentacao
    .Redraw = False
    .RemoveAll
    
    Set rstParamFilial = db.OpenRecordset("SELECT Consignacao_TabelaPrecos FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenSnapshot)
    
    For intX = 0 To grdLista.Rows - 1
      strCodigoProduto = ""
      
      Set rstProdutos = db.OpenRecordset(" SELECT Código FROM Produtos " & _
                                         " WHERE Código = '" & grdLista.Columns("Codigo").Text & "'", dbOpenSnapshot)
      
        If (rstProdutos.BOF And rstProdutos.EOF) Then
          rstProdutos.Close
          Set rstProdutos = Nothing
          
          Set rstGrade = db.OpenRecordset(" SELECT * FROM [Códigos da Grade] " & _
                                          " WHERE Código = '" & grdLista.Columns("Codigo").Text & "'", dbOpenSnapshot)
                                          
          If rstGrade.BOF And rstGrade.EOF Then
            rstGrade.Close
            Set rstGrade = Nothing
          Else
            strCodigoProduto = rstGrade.Fields("Código Original")
          End If
        Else
          strCodigoProduto = rstProdutos.Fields("Código")
        End If
      
      If Len(Trim(strCodigoProduto)) > 0 Then
        Set rstPreco = db.OpenRecordset(" SELECT * FROM Preços " & _
                                        " WHERE Produto = '" & strCodigoProduto & "'" & _
                                        " AND Tabela = '" & rstParamFilial.Fields("Consignacao_TabelaPrecos") & "'", dbOpenSnapshot)
        
        .AddNew
        
        .Columns("Codigo").Text = grdLista.Columns("Codigo").Text
        .Columns("Nome").Text = grdLista.Columns("Nome").Text
        .Columns("Qtde").Text = grdLista.Columns("Qtde").Text
        
        If Not (rstPreco.BOF And rstPreco.EOF) Then
          .Columns("Preco").Text = Format(CStr(CDbl(rstPreco.Fields("Preço"))), FORMAT_VALUE)
        End If
        
        rstPreco.Close
        Set rstPreco = Nothing
        
        lngTotal = lngTotal + .Columns("Qtde").Text
        If Not IsNumeric(.Columns("Preco").Text) Then .Columns("Preco").Text = 0
        dblTotalMonetario = dblTotalMonetario + .Columns("Preco").Text * .Columns("Qtde").Text
        .Update
      End If
      
      grdLista.MoveNext
    Next intX
    
    .Redraw = True
  End With
  
  rstParamFilial.Close
  Set rstParamFilial = Nothing
  
  frmManutencaoConsignacao.lblQtdeTotal.Caption = lngTotal
  frmManutencaoConsignacao.lblTotalMonetario.Caption = Format(CStr(dblTotalMonetario), FORMAT_VALUE)
  Unload Me
End Sub

Private Sub cmdLimpar_Click()
  txtQtde.Text = ""
  txtNomeProduto.Text = ""
  cboProdutos.Text = ""
  cboProdutos.SetFocus
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  datProdutos.DatabaseName = gsQuickDBFileName
End Sub

'02/07/2004 - mpdea
'Seleciona todo o texto do controle
Private Sub txtQtde_GotFocus()
  Call SelectAllText(Me.ActiveControl)
End Sub

Private Sub txtQtde_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    PreencheGrid
  End If
End Sub

Private Sub PreencheGrid()
  Dim intX      As Integer
  Dim blnExiste As Boolean
  Dim dblQtdeTotal As Long
  
  If Len(Trim(cboProdutos.Text)) <= 0 Then Exit Sub
  
  If Not IsNumeric(txtQtde.Text) Or txtQtde.Text = "0" Then
    MsgBox "Quantidade inválida, verifique.", vbExclamation, "Quick Store"
    txtQtde.SetFocus
    Exit Sub
  End If
  
  '17/01/2005 - Daniel
  'Adicionado tratamento para não emprestar Qtde Superior
  'ao estoque
  'Primeiro.: Verificar se o produto existe ou usa grade
  'Segundo..: Verificar a qtde disponível no estoque
  Dim Cód         As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor     As Integer
  Dim Aux_Edição  As Long
  Dim Aux_Produto As String
  Dim Aux_Tipo    As Integer
  Dim Aux_Erro    As Integer
  Dim Cancel      As Integer
  
  Cód = Trim(CStr(cboProdutos.Text))
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  
  Call Acha_Produto(Cód, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Tipo, Aux_Erro)
  If Aux_Erro <> 0 Then
    Cancel = True
    If Aux_Erro = 1 Then
      DisplayMsg "Produto não existe."
    ElseIf Aux_Erro = 2 Then
      DisplayMsg "Produto com grade, digite tamanho e cor."
    ElseIf Aux_Erro = 3 Then
      DisplayMsg "Produto com edição, digite a edição também."
    End If
    Exit Sub
  End If
  
  If Trim(frmManutencaoConsignacao.cboTipoOperacao.Text) = "1 - Consignação" Then
    If EstoqueInsuficiente(Aux_Produto, Aux_Cor, Aux_Tamanho) Then Exit Sub
  Else '2 - Devolução
    'Para a Devolução devemos se preocupar se X produto
    'está consignado, ou seja, se pertence ou não ao conjunto
    'de ítens. Abaixo criamos um tratamento para verificação se
    'determinado produto pertence a este conjunto.
    Dim rstClientes          As Recordset
    Dim rstSaidas            As Recordset
    Dim rstSaProdutos        As Recordset
    Dim lngConsignacaoMestre As Long
    Dim lngSequenciaSaida    As Long
    Dim blnExisteProd        As Boolean
    Dim dblQtdeDevolvida     As Double
    Dim strMensagem          As String
    Dim dblQtdeTotalDevolv   As Double
    Dim dblQtdeTotalEmpres   As Double
    Dim dlbQtdeGrid          As Double
    Dim bytCont              As Byte
    
    Set rstClientes = db.OpenRecordset(" SELECT Código, UltimaConsignacao, ConsignacaoFechada, Nome " & _
                                       " FROM Cli_For WHERE Código = " & lngCodigoCliente, dbOpenSnapshot)
    
    If rstClientes.Fields("ConsignacaoFechada") Then Exit Sub
    
    lngConsignacaoMestre = IIf(IsNumeric(rstClientes.Fields("UltimaConsignacao")), rstClientes.Fields("UltimaConsignacao"), 0)
    
    Set rstSaidas = db.OpenRecordset(" SELECT * FROM Saídas " & _
                                     " WHERE ConsignacaoMestre = " & lngConsignacaoMestre & " AND Filial = " & gnCodFilial, dbOpenSnapshot)
    
    
    With rstSaidas
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do While Not .EOF
          lngSequenciaSaida = .Fields("Sequência")
          Set rstSaProdutos = db.OpenRecordset(" SELECT * FROM [Saídas - Produtos] WHERE Sequência = " & lngSequenciaSaida & " AND Filial = " & gnCodFilial, dbOpenSnapshot)

          With rstSaProdutos
            If Not (.BOF And .EOF) Then
              .MoveFirst
              
              Do While Not .EOF
              
                If Cód = rstSaProdutos.Fields("Código").Value Then
                  blnExisteProd = True
                  
                  bytCont = bytCont + 1
                  
                  If bytCont = 1 Then
                    'Verificação se não estão querendo devolver uma quantidade
                    'superior a que foi emprestada
                    Call VerificaDevolucoes(lngConsignacaoMestre, Cód, dblQtdeDevolvida)
                  
                    dblQtdeTotalDevolv = dblQtdeTotalDevolv + dblQtdeDevolvida
                  End If
                  
                  dblQtdeTotalEmpres = dblQtdeTotalEmpres + .Fields("Qtde").Value
                  
                  Exit Do
                End If
              
                .MoveNext
              Loop
            
            End If

            If Not rstSaProdutos Is Nothing Then .Close
            Set rstSaProdutos = Nothing
          End With
          
          .MoveNext
        Loop
        
      End If
    End With
    
    rstSaidas.Close
    Set rstSaidas = Nothing
    rstClientes.Close
    Set rstClientes = Nothing
    
    If Not blnExisteProd Then
      MsgBox "O produto " & Cód & " não foi emprestado, verifique.", vbCritical, "Quick Store"
      cboProdutos.SetFocus
      Exit Sub
    End If
    
    'Verificação se não estão querendo devolver uma quantidade
    'superior a que foi emprestada
    '
    'Buscar a Qtde do produto caso ele já tenha sido adicionado na grid
    Call BuscarQtdeProdGrid(Cód, dlbQtdeGrid)
    '
    '
    If dblQtdeTotalEmpres < (CDbl(txtQtde.Text) + dblQtdeTotalDevolv + dlbQtdeGrid) Then
      strMensagem = strMensagem & "Foram Emprestados: " & dblQtdeTotalEmpres & vbCrLf
      strMensagem = strMensagem & "Devolvidos até o momento: " & dblQtdeTotalDevolv & vbCrLf
      strMensagem = strMensagem & "Impossível tentar devolver: " & CDbl((txtQtde.Text) + dlbQtdeGrid) & vbCrLf
      
      MsgBox strMensagem, vbExclamation, "Atenção" & " - Produto " & Cód
      Exit Sub
    End If
    
    
  End If
  '------------------------------------------------------------------------------------------
  
  With grdLista
    If Not IsNumeric(txtQtde.Text) Then
      MsgBox "Quantidade inválida !", vbCritical, "Quick Store"
      Exit Sub
    End If
    
    cboProdutos_LostFocus
    
    .Redraw = False
    .MoveFirst

    blnExiste = False
    
    For intX = 0 To .Rows - 1
      If UCase(.Columns(0).Text) = UCase(cboProdutos.Text) Then
        .Columns("Qtde").Text = CDbl(.Columns("Qtde").Text) + CDbl(txtQtde.Text)
        blnExiste = True
      End If
      
      .MoveNext
    Next intX
    
    .Redraw = True
    
    If Not blnExiste Then
      If Not (rsProdutos.BOF And rsProdutos.EOF) Then
        .AddItem cboProdutos.Text & vbTab & _
                 rsProdutos.Fields("Nome") & vbTab & _
                 txtQtde.Text
      End If
    End If
    
    .Redraw = False
    .MoveFirst
    dblQtdeTotal = 0
    
    For intX = 0 To .Rows - 1
      dblQtdeTotal = dblQtdeTotal + .Columns("Qtde").Text
      
      .MoveNext
    Next intX
    
    .Redraw = True
    txtTotal.Text = Format(dblQtdeTotal, FORMAT_VALUE)
  End With
  
  cboProdutos.Text = ""
  txtNomeProduto.Text = ""
  txtQtde.Text = "1"
  
  '02/07/2004 - mpdea
  'Seta o foco para a combo de produtos
  cboProdutos.SetFocus
  
End Sub

Private Sub txtQtde_LostFocus()
  '01/03/2005 - Daniel
  'Comentado em 01/03/2005 para habilitarmos o recurso de limpeza
  'do cabeçalho através do evento click do objeto cmdLimpar
  '
  'PreencheGrid
End Sub

Private Function EstoqueInsuficiente(ByVal CodProduto As String, ByVal Cor As Integer, ByVal Tamanho As Integer) As Boolean
  '17/01/2005 - Daniel
  'Adicionado Function para não emprestar Qtde Superior
  'ao estoque
  Dim rstEstoqueFinal As Recordset
  Dim strSQL          As String
  
  strSQL = "SELECT [Estoque Atual] FROM [Estoque Final] "
  strSQL = strSQL & " WHERE Filial = " & gnCodFilial
  strSQL = strSQL & " AND Produto = '" & CodProduto & "'"
  strSQL = strSQL & " AND Tamanho = " & Tamanho
  strSQL = strSQL & " AND Cor = " & Cor
  
  Set rstEstoqueFinal = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstEstoqueFinal
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      If .Fields("Estoque Atual").Value < CSng(txtQtde.Text) Then
        EstoqueInsuficiente = True
        MsgBox "Estoque insuficiente, a quantidade existente é de " & (.Fields("Estoque Atual").Value), vbExclamation, "Quick Store"
        txtQtde.SetFocus
      End If
      
    End If
    .Close
  End With
  
  Set rstEstoqueFinal = Nothing
  
End Function

Private Sub VerificaDevolucoes(ByVal lngConsignacaoMestre As Long, ByVal CodProd As String, ByRef dblQtdeDevolvida As Double)
  '19/01/2005 - Daniel
  'Adicionado Private para não devolver Qtde Superior
  'aquilo que já foi devolvido
  Dim rstEntradas         As Recordset
  Dim rstEnProdutos       As Recordset
  Dim lngSequenciaEntrada As Long

  dblQtdeDevolvida = 0

  Set rstEntradas = db.OpenRecordset(" SELECT * FROM Entradas " & _
                                     " WHERE ConsignacaoMestre = " & lngConsignacaoMestre & " AND Filial = " & gnCodFilial, dbOpenSnapshot)

  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do While Not .EOF
        lngSequenciaEntrada = .Fields("Sequência")
        Set rstEnProdutos = db.OpenRecordset(" SELECT * FROM [Entradas - Produtos] WHERE Sequência = " & lngSequenciaEntrada & " AND Filial = " & gnCodFilial, dbOpenSnapshot)

        With rstEnProdutos
          If Not (.BOF And .EOF) Then
            .MoveFirst

            Do While Not .EOF
              If CodProd = .Fields("Código").Value Then
                dblQtdeDevolvida = dblQtdeDevolvida + .Fields("Qtde").Value
              End If
              
              .MoveNext
            Loop
          End If

          If Not rstEnProdutos Is Nothing Then .Close
          Set rstEnProdutos = Nothing
        End With 'With rstEnProdutos

        .MoveNext
      Loop

    End If
    .Close
  End With 'With rstEntradas

  Set rstEntradas = Nothing

End Sub

Private Sub BuscarQtdeProdGrid(ByVal CodProd As String, ByRef dblQtdeGrid As Double)
  '28/01/2005 - Daniel
  Dim intX As Integer
  
  dblQtdeGrid = 0
  
  grdLista.MoveFirst
  
  For intX = 0 To (grdLista.Rows - 1)
    If grdLista.Columns("Codigo").Text = CodProd Then dblQtdeGrid = CDbl(grdLista.Columns("Qtde").Text)
  Next intX
  
End Sub

Private Sub grdLista_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
  '01/03/2005 - Daniel
  'Tratamento para Exclusão de linha(s) da grid
  DispPromptMsg = False
  If gbPodeApagar = False Then
    Beep
    Cancel = True
    Exit Sub
  End If
  
  If bGridBeforeDelete Then
    Call StatusMsg("Seleção de itens apagada.")
    Cancel = False
  Else
    Cancel = True
  End If

End Sub

Public Function bGridBeforeDelete() As Boolean
  '01/03/2005 - Daniel
  'Tratamento para Exclusão de linha(s) da grid
  Dim intI    As Integer
  Dim varBook As Variant
  
  gsTitle = LoadResString(201)
  gsMsg = "Apagar a seleção atual?"
  gnStyle = vbYesNo + vbQuestion + vbDefaultButton2
  gnResponse = MsgBox(gsMsg, gnStyle, gsTitle)
  
  If gnResponse = vbNo Then
    bGridBeforeDelete = False
  Else
    bGridBeforeDelete = True
    
    For intI = 0 To (grdLista.SelBookmarks.Count - 1)
      varBook = grdLista.SelBookmarks(intI)
      grdLista.Bookmark = varBook
    
      txtTotal.Text = CDbl(txtTotal.Text) - CDbl(grdLista.Columns("Qtde").CellValue(varBook))
    Next intI
    
  End If
End Function

