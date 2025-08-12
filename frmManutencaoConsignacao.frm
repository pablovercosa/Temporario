VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManutencaoConsignacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção de Consignação"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManutencaoConsignacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   8895
   Begin VB.Data datClientes 
      Caption         =   "datClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome, Tipo FROM Cli_For"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   -120
      TabIndex        =   9
      Top             =   -240
      Width           =   9135
      Begin VB.TextBox txtConsignacao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtCreditoDisponivel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtLimiteCredito 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdContagem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Contagem"
         Height          =   735
         Left            =   1440
         Picture         =   "frmManutencaoConsignacao.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdHistorico 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Historico"
         Height          =   735
         Left            =   240
         Picture         =   "frmManutencaoConsignacao.frx":0B2A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7320
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Left            =   7320
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignação"
         Height          =   255
         Left            =   5880
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito Disponível"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Limite"
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdEfetivar 
      BackColor       =   &H0000C0C0&
      Caption         =   "Efetivar Movimentação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   2535
   End
   Begin VB.ComboBox cboTipoOperacao 
      Height          =   315
      ItemData        =   "frmManutencaoConsignacao.frx":103E
      Left            =   120
      List            =   "frmManutencaoConsignacao.frx":1048
      TabIndex        =   7
      Text            =   "1 - Consignação"
      Top             =   6000
      Width           =   2415
   End
   Begin SSDataWidgets_B.SSDBGrid grdMovimentacao 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   8655
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
      Col.Count       =   4
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Caption=   "Codigo"
      Columns(0).Name =   "Codigo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   6085
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2302
      Columns(2).Caption=   "Qtde"
      Columns(2).Name =   "Qtde"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   2381
      Columns(3).Caption=   "Preço"
      Columns(3).Name =   "Preco"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      _ExtentX        =   15266
      _ExtentY        =   7011
      _StockProps     =   79
      Caption         =   "Resumo da Movimentação"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskDataProximoAcerto 
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtNomeCliente 
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
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5055
   End
   Begin SSDataWidgets_B.SSDBCombo cboCliente 
      Bindings        =   "frmManutencaoConsignacao.frx":106C
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
      DataFieldList   =   "Código"
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
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "Código"
      Columns(0).Name =   "Código"
      Columns(0).DataField=   "Código"
      Columns(0).FieldLen=   256
      Columns(1).Width=   7038
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1).FieldLen=   256
      Columns(2).Width=   1244
      Columns(2).Caption=   "Tipo"
      Columns(2).Name =   "Tipo"
      Columns(2).DataField=   "Tipo"
      Columns(2).FieldLen=   256
      _ExtentX        =   3625
      _ExtentY        =   556
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Código"
   End
   Begin VB.Label lblTotalMonetario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Total R$:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label lblQtdeTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Qtde Total:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo da Operação"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Próximo Acerto"
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "frmManutencaoConsignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Public rstParamFilial As Recordset
  Public rstConsignacao As Recordset
  Public rstConsProduto As Recordset
  
  Public Enum enuMovimentacao
    enuEntrada = 1
    enuSaida = 2
  End Enum
  
  Dim lngSequencia As Long
  Dim lngbuffer As Long

  Dim blnErro As Boolean

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(0).Text
  cboCliente_LostFocus
End Sub

Private Sub cboCliente_LostFocus()
  Dim rstX As Recordset
  
  txtNomeCliente.Text = ""
  If Not IsNumeric(cboCliente.Text) Then Exit Sub
  
  Set rstX = db.OpenRecordset(" SELECT Código, Nome, [Limite Crédito] , DataProxAcertoConsignacao, ConsignacaoFechada, UltimaConsignacao FROM Cli_For WHERE Código = " & cboCliente.Text, dbOpenDynaset)
  
  With rstX
    If Not (.BOF And .EOF) Then
      txtNomeCliente.Text = .Fields("Nome") & ""
      txtConsignacao.Text = .Fields("UltimaConsignacao") & ""
      
      If .Fields("ConsignacaoFechada") Then
        lblStatus.Caption = "Fechada"
      ElseIf Not .Fields("ConsignacaoFechada") Then
        lblStatus.Caption = "Aberta"
      End If
      
      txtLimiteCredito.Text = Format(CStr(.Fields("Limite Crédito") & ""), FORMAT_VALUE)
      txtCreditoDisponivel.Text = Format(CStr(.Fields("Limite Crédito") - getLimiteUsado(cboCliente.Text)), FORMAT_VALUE)
      
      If IsDate(.Fields("DataProxAcertoConsignacao")) Then
        mskDataProximoAcerto.Text = CDate(.Fields("DataProxAcertoConsignacao"))
      Else
        mskDataProximoAcerto.Mask = ""
        mskDataProximoAcerto.Text = ""
        mskDataProximoAcerto.Mask = "##/##/####"
      End If
    End If
    
    If Not rstX Is Nothing Then .Close
    Set rstX = Nothing
  End With
End Sub

Private Sub cmdContagem_Click()
  '18/01/2005 - Daniel
  'Tratamento para em caso de Devolução não digitar
  'produtos que não fazem parte do conjunto de ítens
  'consignados
  If cboTipoOperacao.Text = "2 - Devolução" Then
    If Len(txtNomeCliente.Text) > 0 Then
      frmContagemProdutos.lngCodigoCliente = CLng(cboCliente.Text)
    Else
      MsgBox "Escolha o cliente.", vbExclamation, "Quick Store"
      cboCliente.SetFocus
      Exit Sub
    End If
  End If
  
  frmContagemProdutos.Show vbModal
  
End Sub

Private Sub cmdEfetivar_Click()
  If (Not IsNumeric(lblTotalMonetario.Caption)) Then
    MsgBox "Valor total da consignação inválido !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If (Not IsNumeric(txtCreditoDisponivel.Text)) Then
    MsgBox "Valor do crédito disponível inválido !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If GetCodigoCombos(cboTipoOperacao.Text) = "1" Then
    If CDbl(lblTotalMonetario.Caption) > CDbl(txtCreditoDisponivel.Text) Then
      MsgBox "O valor total da consignação está passando o saldo disponível para novas compras, verifique !", vbCritical, "Quick Store"
      Exit Sub
    End If
  End If
  
  blnErro = False
  
  If grdMovimentacao.Rows <= 0 Then
    MsgBox "Movimentação sem produtos, verifique !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  Select Case GetCodigoCombos(cboTipoOperacao.Text)
    Case "1"  'Consignação
      GeraMovimentacao enuSaida
    Case "2"  'Devolução
      GeraMovimentacao enuEntrada
    Case Else
      MsgBox "Opção inválida, verifique !", vbCritical, "Quick Store"
  End Select
  
  If Not blnErro Then
    Select Case GetCodigoCombos(cboTipoOperacao.Text)
      Case "1"
        If MsgBox("Operação efetivada com sucesso ! Deseja abrir a tela de saídas ? ", vbQuestion + vbYesNo, App.Title) = vbYes Then
          With frmSaidas
            .txtSeq.Text = lngSequencia
            .SearchRecord
            .Show
          End With
        End If
      Case "2"
        If MsgBox("Operação efetivada com sucesso ! Deseja abrir a tela de entradas ? ", vbQuestion + vbYesNo, App.Title) = vbYes Then
          With frmEntrada
            .txtSeq.Text = lngSequencia
            .SearchRecord
            .Show
          End With
        End If
    End Select
    ClearScreen
  End If
End Sub

Private Sub GeraMovimentacao(TipoMovimentacao As enuMovimentacao)
  Dim lngCliente            As Long
  Dim lngConsignacaoMestre  As Long
  Dim blnInTransaction      As Boolean
  Dim intX                  As Integer
  Dim rstParametros         As Recordset
  Dim strCodProdSemGrade    As String
  
  If (Len(Trim(txtNomeCliente.Text))) <= 0 Then
    MsgBox "Cliente incorreto, verifique !!", vbCritical, "Quick Store"
    blnErro = True
    Exit Sub
  End If
  
  lngCliente = cboCliente.Text & ""
  
  lngConsignacaoMestre = getNumeroConsignacao(lngCliente)
  lngSequencia = gnGetNextSequencia(gnCodFilial)
  
  ws.BeginTrans: blnInTransaction = True
  
  Set rstParametros = db.OpenRecordset("SELECT [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenDynaset)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      .Edit
      .Fields("[Última Movimentação]") = lngSequencia
      .Update
    End If
    .Close
    Set rstParametros = Nothing
  End With
  
  If TipoMovimentacao = enuSaida Then
    Set rstConsignacao = db.OpenRecordset(" SELECT * FROM Saídas WHERE Cliente = " & lngCliente & " AND Filial = " & gnCodFilial, dbOpenDynaset)
  ElseIf TipoMovimentacao = enuEntrada Then
    Set rstConsignacao = db.OpenRecordset(" SELECT * FROM Entradas WHERE Fornecedor = " & lngCliente & " AND Filial = " & gnCodFilial, dbOpenDynaset)
  End If
  
  With rstConsignacao
    .AddNew
    
    .Fields("Filial") = gnCodFilial
    .Fields("Data") = Data_Atual
    .Fields("Sequência") = lngSequencia
    .Fields("ConsignacaoMestre") = lngConsignacaoMestre
    .Fields("Caixa") = rstParamFilial.Fields("Consignacao_Caixa")
    .Fields("Digitador") = gnUserCode
    
    If TipoMovimentacao = enuSaida Then
      .Fields("Tabela") = rstParamFilial.Fields("Consignacao_TabelaPrecos")
      .Fields("Operador") = gnUserCode
      .Fields("Recebimento") = False
      .Fields("Cliente") = lngCliente
      .Fields("Operação") = rstParamFilial.Fields("Consignacao_OpSaida")
      .Fields("Serviços") = 0
      .Fields("DescontoSubTotal") = 0
      .Fields("Serviços") = 0
    ElseIf TipoMovimentacao = enuEntrada Then
      .Fields("Fornecedor") = lngCliente
      .Fields("Operação") = rstParamFilial.Fields("Consignacao_OpEntrada")
    End If
    
    .Fields("Total") = lblTotalMonetario.Caption
    .Fields("Produtos") = lblTotalMonetario.Caption
    
    .Update
  End With
  
  If TipoMovimentacao = enuSaida Then
    Set rstConsProduto = db.OpenRecordset("SELECT * FROM [Saídas - Produtos]", dbOpenDynaset)
  ElseIf TipoMovimentacao = enuEntrada Then
    Set rstConsProduto = db.OpenRecordset("SELECT * FROM [Entradas - Produtos]", dbOpenDynaset)
  End If
  
  With grdMovimentacao
    .MoveFirst
    
    For intX = 0 To .Rows - 1
      rstConsProduto.AddNew
      
      rstConsProduto.Fields("Filial") = gnCodFilial
      rstConsProduto.Fields("Sequência") = lngSequencia
      rstConsProduto.Fields("Linha") = intX + 1
      rstConsProduto.Fields("Código") = .Columns("Codigo").Text
      rstConsProduto.Fields("Qtde") = .Columns("Qtde").Text
      rstConsProduto.Fields("Preço") = IIf(IsNumeric(.Columns("Preco").Text), .Columns("Preco").Text, 0)
      rstConsProduto.Fields("Preço Final") = CDbl(rstConsProduto.Fields("Preço") * .Columns("Qtde").Text)
      '18/01/2005 - Daniel
      'Tratamento para produtos que utilizam a Grade
      Call BuscarCodProdSem(.Columns("Codigo").Text, strCodProdSemGrade)
      rstConsProduto.Fields("Código Sem Grade") = strCodProdSemGrade
      'Antiga linha comentada em 18/01/2005
      'rstConsProduto.Fields("Código Sem Grade") = .Columns("Codigo").Text
      
      'Antiga linha comentada em 18/01/2005
      'rstConsProduto.Fields("Icm") = GetFieldInProduto(.Columns("Codigo").Text, "Percentual ICM")
      rstConsProduto.Fields("Icm") = GetFieldInProduto(strCodProdSemGrade, "Percentual ICM")
      
      If TipoMovimentacao = enuSaida Then
        'Antiga linha comentada em 18/01/2005
        'rstConsProduto.Fields("Situação Tributária") = GetFieldInProduto(.Columns("Codigo").Text, "Situação Tributária")
        rstConsProduto.Fields("Situação Tributária") = GetFieldInProduto(strCodProdSemGrade, "Situação Tributária")
        
        If Len(rstConsProduto.Fields("Situação Tributária")) <= 0 Then
          rstConsProduto.Fields("Situação Tributária") = " "
        End If
        
        'Antiga linha comentada em 18/01/2005
        'rstConsProduto.Fields("Unidade Venda") = GetFieldInProduto(.Columns("Codigo").Text, "Unidade Venda")
        rstConsProduto.Fields("Unidade Venda") = GetFieldInProduto(strCodProdSemGrade, "Unidade Venda")
        If Len(Trim(rstConsProduto.Fields("Unidade Venda").Value)) <= 0 Then
          rstConsProduto.Fields("Unidade Venda") = "un"
        End If
        
        rstConsProduto.Fields("QtdeEntregue") = 0
      End If
      
      rstConsProduto.Update
      
      .MoveNext
    Next intX
  End With
  
  If TipoMovimentacao = enuEntrada Then
    Efetiva_Entrada gnCodFilial, lngSequencia
  ElseIf TipoMovimentacao = enuSaida Then
    Efetiva_Saída gnCodFilial, lngSequencia
  End If
  
  ws.CommitTrans: blnInTransaction = False
  
  If Not rstConsignacao Is Nothing Then rstConsignacao.Close
  If Not rstConsProduto Is Nothing Then rstConsProduto.Close
  
  Set rstConsignacao = Nothing
  Set rstConsProduto = Nothing
  
  Exit Sub
  
Erro:
  If blnInTransaction Then: ws.Rollback
  blnErro = True
End Sub

Private Sub cmdPrintTicket_Click()
  PrintTicketConsignacao
End Sub

Private Sub cmdHistorico_Click()
  If Len(Trim(txtNomeCliente.Text)) <= 0 Then
    MsgBox "Cliente incorreto, verifique !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  With frmHistoricoConsignacao
    .lngCodigoCliente = CLng(cboCliente.Text)
    .Show
  End With
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  Set rstParamFilial = db.OpenRecordset(" SELECT [Parâmetros Filial].Filial, [Parâmetros Filial].Consignacao_OpEntrada, [Parâmetros Filial].Consignacao_OpSaida, [Parâmetros Filial].Consignacao_Caixa, [Parâmetros Filial].Consignacao_TabelaPrecos " & _
                                        " FROM [Parâmetros Filial] " & _
                                        " WHERE ([Parâmetros Filial].Filial = " & gnCodFilial & ") ", dbOpenDynaset)
  With rstParamFilial
    If (.BOF And .EOF) Then
      MsgBox "Atenção, a filial usada é inválida !", vbCritical, "Quick Store"
      Unload Me
    Else
      VerificaParametros rstParamFilial
    End If
  End With
  
  datClientes.DatabaseName = gsQuickDBFileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not rstParamFilial Is Nothing Then rstParamFilial.Close
  Set rstParamFilial = Nothing
End Sub

Private Sub PrintTicketConsignacao()
  Dim dblQtdeTotal As Double
  Dim rstClientes As Recordset
  Dim lngConsignacaoMestre As Long
  Dim bytTipoMovimentacao As Byte    ' 1 - Entrada, 2 - Saída
  
  If Len(Trim(txtNomeCliente.Text)) <= 0 Then
    MsgBox "Cliente incorreto, verifique !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  Set rstClientes = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código = " & cboCliente.Text, dbOpenSnapshot)
  
  With rstClientes
    If Not (.BOF And .EOF) Then
      lngConsignacaoMestre = .Fields("UltimaConsignacao") & ""
    End If
  End With
  
  Set rstConsignacao = db.OpenRecordset("SELECT * FROM Entradas WHERE Sequência = " & lngSequencia, dbOpenSnapshot)
  
  With rstConsignacao
    If (.BOF And .EOF) Then
      If Not rstConsignacao Is Nothing Then .Close
      Set rstConsignacao = Nothing
      
      Set rstConsignacao = db.OpenRecordset("SELECT * FROM Saídas WHERE Sequência = " & lngSequencia, dbOpenSnapshot)
      
      With rstConsignacao
        If (.BOF And .EOF) Then
          MsgBox "Sequência de movimentação inválida !", vbCritical, "Quick Store"
          
          If Not rstConsignacao Is Nothing Then .Close
          Set rstConsignacao = Nothing
          
          Exit Sub
        Else
          bytTipoMovimentacao = 2
        End If
      End With
    Else
      bytTipoMovimentacao = 1
    End If
  End With
  
  With rstConsignacao
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not rstConsProduto Is Nothing Then rstConsProduto.Close
      Set rstConsProduto = Nothing
      
      Set rstConsProduto = db.OpenRecordset("SELECT * FROM [Entradas - Produtos] WHERE Sequência = " & lngSequencia, dbOpenSnapshot)
      
      If rstConsProduto.EOF Then
        Set rstConsProduto = db.OpenRecordset("SELECT * FROM [Saídas - Produtos] WHERE Sequência = " & lngSequencia, dbOpenSnapshot)
      End If
      
      With rstConsProduto
        If (.BOF And .EOF) Then
          MsgBox "Movimentação sem produtos a exibir !", vbCritical, "Quick Store"
          Exit Sub
        End If
      End With
      
      lngbuffer = FreeFile
      
      Open "LPT1" For Output As #lngbuffer
      
      LetPrinter Chr$(27) & "@"
      
      LetPrinter ""
      LetPrinter "|----------------------------------------------|"
      LetPrinter "|           Sistema de Consignacao             |"
      LetPrinter "|----------------------------------------------|"
      LetPrinter ""
      
      LetPrinter " Data        : " & rstConsignacao.Fields("Data")
      LetPrinter " Consignacao : " & rstConsignacao.Fields("ConsignacaoMestre")
      LetPrinter " Sequencia   : " & rstConsignacao.Fields("Sequência")
      
      If bytTipoMovimentacao = 1 Then
        LetPrinter " Vendedor    : " & rstConsignacao.Fields("Digitador") & " - " & getNomeFuncionario(rstConsignacao.Fields("Digitador"))
      Else
        LetPrinter " Vendedor    : " & rstConsignacao.Fields("Operador") & " - " & getNomeFuncionario(rstConsignacao.Fields("Operador"))
      End If
      
      Print #lngbuffer, Chr$(27) & Chr(69);   'Liga o negrito
      If bytTipoMovimentacao = 1 Then
        LetPrinter " Cliente     : " & rstConsignacao.Fields("Fornecedor") & " - " & getNomeCliente(rstConsignacao.Fields("Fornecedor"))
      Else
        LetPrinter " Cliente     : " & rstConsignacao.Fields("Cliente") & " - " & getNomeCliente(rstConsignacao.Fields("Cliente"))
      End If
      Print #lngbuffer, Chr$(27) & Chr(70);   'Desliga o negrito
      
      LetPrinter " Saldo       : " & "R$ 1234,12"
      
      LetPrinter ""
      LetPrinter " Produto                                   Qtde "
      LetPrinter "------------------------------------------------"
      
      dblQtdeTotal = 0
      
      rstConsProduto.MoveFirst
      Do While Not rstConsProduto.EOF
        LetPrinter Left(rstConsProduto.Fields("Código") & Space(15), 14), True
        LetPrinter Left(getNomeProduto(rstConsProduto.Fields("Código")) & Space(30), 29), True
        LetPrinter Right(Space(6) & rstConsProduto.Fields("Qtde"), 5)
        
        dblQtdeTotal = dblQtdeTotal + rstConsProduto.Fields("Qtde")
        
        rstConsProduto.MoveNext
      Loop
      
      LetPrinter "|----------------------------------------------|"
      LetPrinter "|                                Total: " & Right(Space(6) & Format(dblQtdeTotal, "#####"), 6) & " |"
      LetPrinter "|----------------------------------------------|"
      
      '---[Dá o espaçamento final]---'
        Dim intEs As Integer
        For intEs = 1 To 12
          LetPrinter ""
        Next intEs
      '---[Dá o espaçamento final]---'
    
      Close #lngbuffer
      
      .MoveNext
    Loop
  End With
End Sub

Private Sub LetPrinter(strString As String, Optional blnSameLine As Boolean = False)
  If blnSameLine Then
    Print #lngbuffer, strString;
  Else
    Print #lngbuffer, strString
  End If
End Sub

Private Function getNomeProduto(strCodigo As String) As String
  Dim rstProdutos As Recordset
  
  Set rstProdutos = db.OpenRecordset("SELECT Nome FROM Produtos WHERE Código = '" & strCodigo & "'", dbOpenSnapshot)
  
  With rstProdutos
    If (.BOF And .EOF) Then
      getNomeProduto = "<Produto_sem_nome>"
    Else
      getNomeProduto = .Fields("Nome") & ""
    End If
    
    If Not rstProdutos Is Nothing Then .Close
    Set rstProdutos = Nothing
  End With
End Function

Private Function getNomeFuncionario(lngCodigo As Long) As String
  Dim rstFuncionarios As Recordset
  
  Set rstFuncionarios = db.OpenRecordset("SELECT Nome FROM Funcionários WHERE Código = " & lngCodigo, dbOpenSnapshot)
  
  With rstFuncionarios
    If (.BOF And .EOF) Then
      getNomeFuncionario = "<Funcionario_sem_nome>"
    Else
      getNomeFuncionario = .Fields("Nome") & ""
    End If
    
    If Not rstFuncionarios Is Nothing Then .Close
    Set rstFuncionarios = Nothing
  End With
End Function

Private Function getNomeCliente(lngCodigo As Long) As String
  Dim rstClientes As Recordset
  
  Set rstClientes = db.OpenRecordset("SELECT Nome FROM Cli_For WHERE Código = " & lngCodigo, dbOpenSnapshot)
  
  With rstClientes
    If (.BOF And .EOF) Then
      getNomeCliente = "<Cliente_sem_nome>"
    Else
      getNomeCliente = .Fields("Nome") & ""
    End If
    
    If Not rstClientes Is Nothing Then .Close
    Set rstClientes = Nothing
  End With
End Function

Private Function getNumeroConsignacao(ByVal lngCliente As Long) As Long
  Dim rstCliente As Recordset
  Dim lngConsignacaoMestre As Long
  Dim strDataProximoAcerto As String
  
  Dim intDia As Integer
  Dim intMes As Integer
  Dim intAno As Integer
  
  Dim strInputData As String
  
  Set rstCliente = db.OpenRecordset("SELECT * FROM Cli_For WHERE Código = " & lngCliente, dbOpenDynaset)
  
  With rstCliente
    If (Not .Fields("ConsignacaoFechada")) Then
      If IsNull(.Fields("UltimaConsignacao")) Then
        lngConsignacaoMestre = gnGetNextConsignacao(gnCodFilial)
      Else
        lngConsignacaoMestre = .Fields("UltimaConsignacao")
      End If
    Else
      lngConsignacaoMestre = gnGetNextConsignacao(gnCodFilial)
      
      .Edit
      .Fields("UltimaConsignacao") = lngConsignacaoMestre
      .Fields("ConsignacaoFechada") = False
      
      If IsNull(.Fields("DiaBaseConsignacao")) Then
        .Fields("DiaBaseConsignacao") = 1
      End If
      
      intDia = .Fields("DiaBaseConsignacao")
      intMes = Month(Data_Atual)
      intAno = Year(Data_Atual)
      
      If intDia <= Day(Data_Atual) Then
        If intMes = 12 Then
          intAno = Year(Data_Atual) + 1
          intMes = 1
        Else
          intMes = intMes + 1
        End If
      End If
      
      strDataProximoAcerto = intDia & "/" & _
                            intMes & "/" & _
                            intAno
      
      Do Until IsDate(strInputData)
        strInputData = InputBox("Confirma a data do próximo acerto ?", "Quick Store", CDate(strDataProximoAcerto))
      Loop
      
      .Fields("DataProxAcertoConsignacao") = CDate(strInputData)
      .Update
    End If
  End With
  
  rstCliente.Edit
  rstCliente.Fields("UltimaConsignacao") = lngConsignacaoMestre
  rstCliente.Update
  
  If Not rstCliente Is Nothing Then rstCliente.Close
  Set rstCliente = Nothing
  
  getNumeroConsignacao = lngConsignacaoMestre
End Function

Private Sub ClearScreen()
  cboCliente.Text = ""
  txtNomeCliente.Text = ""
  mskDataProximoAcerto.Mask = ""
  mskDataProximoAcerto.Text = ""
  mskDataProximoAcerto.Mask = "##/##/####"
  lblQtdeTotal.Caption = "0"
  lblTotalMonetario.Caption = "0"
  
  grdMovimentacao.Redraw = False
  grdMovimentacao.RemoveAll
  grdMovimentacao.Redraw = True
End Sub

Private Function VerificaParametros(ByRef rstParametros As Recordset) As Boolean
  Dim blnErro As Boolean: blnErro = False
  
  With rstParametros
    '---[ Verificação dos parâmetros do sistema ]---'
      If (Not IsNumeric(.Fields("Consignacao_OpEntrada"))) Then
        blnErro = True
      Else
        If .Fields("Consignacao_OpEntrada") = 0 Then
          blnErro = True
        End If
      End If
      
      If blnErro Then
        MsgBox "Operação de entrada incorreta !! Você pode selecionar uma operação de entrada na tela de parâmetros/ filial", vbCritical, "Quick Store"
        VerificaParametros = False
        Exit Function
      End If
      
      '-------------------------------------------------------------------------'
      
      If (Not IsNumeric(.Fields("Consignacao_OpSaida"))) Then
        blnErro = True
      Else
        If .Fields("Consignacao_OpSaida") = 0 Then
          blnErro = True
        End If
      End If
      
      If blnErro Then
        MsgBox "Operação de saída incorreta !! Você pode selecionar uma operação de saída na tela de parâmetros/ filial", vbCritical, "Quick Store"
        VerificaParametros = False
        Exit Function
      End If
      
      '-------------------------------------------------------------------------'
      
      If (Not IsNumeric(.Fields("Consignacao_Caixa"))) Then
        blnErro = True
      Else
        If .Fields("Consignacao_Caixa") = 0 Then
          blnErro = True
        End If
      End If
      
      If blnErro Then
        MsgBox "Caixa incorreto !! Você pode selecionar um caixa na tela de parâmetros/ filial", vbCritical, "Quick Store"
        VerificaParametros = False
        Exit Function
      End If
      
      '-------------------------------------------------------------------------'
      
      blnErro = Len(Trim(.Fields("Consignacao_TabelaPrecos"))) <= 0
      
      If blnErro Then
        MsgBox "Tabela de preços incorreta !! Você pode selecionar uma tabela de preços na tela de parâmetros/ filial", vbCritical, "Quick Store"
        VerificaParametros = False
        Exit Function
      End If
    '---[ Verificação dos parâmetros do sistema ]---'
  End With
End Function

Private Function GetFieldInProduto(strCodigoProduto As String, strField As String) As String
  Dim rstProdutos As Recordset
  
  Set rstProdutos = db.OpenRecordset(" SELECT [" & strField & "] FROM Produtos " & _
                                     " WHERE Código = '" & strCodigoProduto & "'", dbOpenSnapshot)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      GetFieldInProduto = .Fields(strField).Value & ""
    Else
      GetFieldInProduto = ""
    End If
    
    .Close
    Set rstProdutos = Nothing
  End With
End Function

Private Function getLimiteUsado(lngCodigoCliente As Long) As Double
  Dim rstClientes As Recordset
  Dim rstConsignacao As Recordset
  Dim dblValorUsado As Double
  
  Set rstClientes = db.OpenRecordset("SELECT Código, UltimaConsignacao, ConsignacaoFechada FROM Cli_For WHERE Código = " & lngCodigoCliente, dbOpenSnapshot)
  With rstClientes
    If Not (.BOF And .EOF) Then
      If Not IsNumeric(.Fields("UltimaConsignacao").Value) Then
        getLimiteUsado = 0
        Exit Function
      End If
      
      If Not .Fields("ConsignacaoFechada") Then
        'Pega as saídas
        Set rstConsignacao = db.OpenRecordset("SELECT Total FROM Saídas WHERE ConsignacaoMestre = " & .Fields("UltimaConsignacao") & " AND Filial = " & gnCodFilial, dbOpenSnapshot)
        With rstConsignacao
          If Not (.BOF And .EOF) Then
            If Not .EOF Then .MoveFirst
            
            Do Until .EOF
              dblValorUsado = dblValorUsado + .Fields("Total")
              .MoveNext
            Loop
          End If
          
          .Close
          Set rstConsignacao = Nothing
        End With
        
        'Pega as entradas
        Set rstConsignacao = db.OpenRecordset("SELECT Total FROM Entradas WHERE ConsignacaoMestre = " & .Fields("UltimaConsignacao") & " AND Filial = " & gnCodFilial, dbOpenSnapshot)
        With rstConsignacao
          If Not (.BOF And .EOF) Then
            If Not .EOF Then .MoveFirst
            
            Do Until .EOF
              dblValorUsado = dblValorUsado - .Fields("Total")
              .MoveNext
            Loop
          End If
          
          .Close
          Set rstConsignacao = Nothing
        End With
      End If
    End If
    
    .Close
    Set rstClientes = Nothing
  End With
  
  getLimiteUsado = dblValorUsado
End Function

Private Sub BuscarCodProdSem(ByVal CodProd As String, ByRef CodProdSemGrade As String)
  '18/01/2005 - Daniel
  'Private criada para atender o tratamento de produtos que
  'utilizam a grade
  Dim Cód         As String
  Dim Aux_Tamanho As Integer
  Dim Aux_Cor     As Integer
  Dim Aux_Edição  As Long
  Dim Aux_Produto As String
  Dim Aux_Tipo    As Integer
  Dim Aux_Erro    As Integer
  Dim Cancel      As Integer
  
  Cód = CodProd
  Aux_Tamanho = 0
  Aux_Cor = 0
  Aux_Edição = 0
  
  Call Acha_Produto(Cód, Aux_Produto, Aux_Tamanho, Aux_Cor, Aux_Edição, Aux_Tipo, Aux_Erro)
  
  CodProdSemGrade = Aux_Produto

End Sub
