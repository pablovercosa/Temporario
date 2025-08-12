VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmHistoricoConsignacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico das consignações"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   Icon            =   "frmHistoricoConsignacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8865
   Begin VB.CommandButton cmdFechamento 
      Caption         =   "Fechamento"
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
      TabIndex        =   13
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Totais das Quantidades "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   8655
      Begin VB.TextBox txtTotalPagar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtTotalConsignado 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTotalDevolvido 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtTotalSaldo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Total R$:"
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
         Left            =   6720
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Consignado:"
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
         TabIndex        =   9
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Devolvido:"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Saldo:"
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
         Left            =   4680
         TabIndex        =   7
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
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
      Left            =   7320
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin SSDataWidgets_B.SSDBGrid grdResumoConsignacoes 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   480
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
      Col.Count       =   8
      AllowAddNew     =   -1  'True
      RowHeight       =   423
      Columns.Count   =   8
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "Data"
      Columns(0).Name =   "Data"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2514
      Columns(1).Caption=   "Produto"
      Columns(1).Name =   "Produto"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3836
      Columns(2).Name =   "ProdutoNome"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1482
      Columns(3).Caption=   "Saída"
      Columns(3).Name =   "QtdeSaida"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   1482
      Columns(4).Caption=   "Entrada"
      Columns(4).Name =   "QtdeEntrada"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   1482
      Columns(5).Caption=   "Saldo"
      Columns(5).Name =   "Saldo"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1561
      Columns(6).Caption=   "Preço"
      Columns(6).Name =   "Preco"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   1561
      Columns(7).Caption=   "Vl Unit"
      Columns(7).Name =   "PrecoUnitario"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      _ExtentX        =   15266
      _ExtentY        =   6588
      _StockProps     =   79
      Caption         =   "Resumo das Consignações"
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
   Begin VB.Label lblNomeCliente 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmHistoricoConsignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  Public lngCodigoCliente As Long
  Dim lngbuffer As Long
  
Private Sub cmdFechamento_Click()
  Dim rstConsignacao        As Recordset
  Dim rstConsProduto        As Recordset
  Dim rstCliente            As Recordset
  Dim lngConsignacaoMestre  As Long
  Dim blnInTransaction      As Boolean
  Dim intX                  As Integer
  Dim blnErro               As Boolean
  Dim lngSequencia          As Long
  Dim rstParamFilial        As Recordset
  Dim strCodProdSemGrade    As String
  
  If lngCodigoCliente <= 0 Then
    MsgBox "Cliente incorreto, verifique !!", vbCritical, "Quick Store"
    blnErro = True
    Exit Sub
  End If
  
  Set rstCliente = db.OpenRecordset(" SELECT UltimaConsignacao, ConsignacaoFechada FROM Cli_For WHERE Código = " & lngCodigoCliente, dbOpenDynaset)
  
  If rstCliente.Fields("ConsignacaoFechada") Then
    MsgBox "ATENÇÃO" & vbCrLf & vbCrLf & "O cliente não tem consignação em aberto", vbInformation, "Quick Store"
    Exit Sub
  End If
  
  
  Set rstParamFilial = db.OpenRecordset("SELECT Consignacao_Caixa, Consignacao_TabelaPrecos, Consignacao_OpFechamento FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenSnapshot)
  
  lngSequencia = gnGetNextSequencia(gnCodFilial)
  ws.BeginTrans: blnInTransaction = True
  
  Set rstConsignacao = db.OpenRecordset(" SELECT * FROM Saídas WHERE Cliente = " & lngCodigoCliente, dbOpenDynaset)
  
  With rstConsignacao
    .AddNew
    
    .Fields("Filial") = gnCodFilial
    .Fields("Data") = Data_Atual
    .Fields("Sequência") = lngSequencia
    .Fields("ConsignacaoMestre") = lngConsignacaoMestre
    .Fields("Caixa") = rstParamFilial.Fields("Consignacao_Caixa")
    .Fields("Digitador") = gnUserCode
    .Fields("Tabela") = rstParamFilial.Fields("Consignacao_TabelaPrecos")
    .Fields("Operador") = gnUserCode
    .Fields("Recebimento") = False
    .Fields("Cliente") = lngCodigoCliente
    .Fields("Operação") = rstParamFilial.Fields("Consignacao_OpFechamento")
    .Fields("Total") = txtTotalPagar.Text
    .Fields("Produtos") = txtTotalPagar.Text
    .Fields("Serviços") = 0
    .Fields("Efetivada") = False
    
    .Update
  End With
  
  Set rstConsProduto = db.OpenRecordset("SELECT * FROM [Saídas - Produtos]", dbOpenDynaset)
  
  With grdResumoConsignacoes
    .MoveFirst
    
    For intX = 0 To .Rows - 1
      If .Columns("Saldo").Text > 0 Then
        rstConsProduto.AddNew
        rstConsProduto.Fields("Filial") = gnCodFilial
        rstConsProduto.Fields("Sequência") = lngSequencia
        rstConsProduto.Fields("Linha") = intX + 1
        rstConsProduto.Fields("Código") = .Columns("Produto").Text
        rstConsProduto.Fields("Qtde") = .Columns("Saldo").Text
        rstConsProduto.Fields("Preço") = CDbl(.Columns("Preco").Text) / CDbl(.Columns("Saldo").Text)
        rstConsProduto.Fields("Preço Final") = CDbl(rstConsProduto.Fields("Preço"))
        
        '18/01/2005 - Daniel
        'Tratamento para produtos que usam a Grade
        Call BuscarCodProdSem(.Columns("Produto").Text, strCodProdSemGrade)
        rstConsProduto.Fields("Código Sem Grade") = strCodProdSemGrade
        'Antiga linha comentada em 18/01/2005
        'rstConsProduto.Fields("Código Sem Grade") = .Columns("Produto").Text
        
        rstConsProduto.Update
      End If
      .MoveNext
    Next intX
  End With
  
  With rstCliente
    .Edit
    .Fields("ConsignacaoFechada") = True
    .Fields("UltimaConsignacao") = Null
    .Update
  End With
  
  ws.CommitTrans: blnInTransaction = False

  If Not rstConsignacao Is Nothing Then rstConsignacao.Close
  If Not rstConsProduto Is Nothing Then rstConsProduto.Close
  If Not rstCliente Is Nothing Then rstCliente.Close
  
  Set rstConsignacao = Nothing
  Set rstConsProduto = Nothing
  Set rstCliente = Nothing
  
  If Not blnErro Then
    If MsgBox("Fechamento efetuado com sucesso ! Deseja abrir a tela de saída para fazer o recebimento ?", vbQuestion + vbYesNo, "Quick Store") = vbYes Then
      With frmSaidas
        .txtSeq.Text = lngSequencia
        .SearchRecord
        .Show
      End With
    End If
  End If
  
  Exit Sub
  
Erro:
  If blnInTransaction Then
    ws.Rollback
    blnInTransaction = False
    blnErro = True
  End If
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim intX As Integer
  Dim strPath As String
  
  On Error GoTo ErrView
  
  lngbuffer = FreeFile
  strPath = InputBox("Insira a porta onde será impresso o extrato", "Quick Store", "LPT1")
  
  If Len(Trim(strPath)) <= 0 Then Exit Sub
  
  Open strPath For Output As #lngbuffer
  
  'LetPrinter Chr$(27) & "@"
  LetPrinter ""
  LetPrinter "|----------------------------------------------------------------|"
  LetPrinter "|  Extrato das Consignacoes                                      |"
  LetPrinter "|----------------------------------------------------------------|"
  LetPrinter ""
  LetPrinter "Cliente.....: " & Left(lblNomeCliente.Caption, 45) & ""
  LetPrinter "Data........: " & Format(Data_Atual, "dd/mm/yyyy")
  LetPrinter ""
  LetPrinter "Produto                      Cons  Dev  Sal     Preco     Unid"
  With grdResumoConsignacoes
    .MoveFirst
    
    For intX = 0 To .Rows - 1
      LetPrinter Left(.Columns("Produto").Text & " - " & _
                      .Columns("ProdutoNome").Text & _
                       Space(28), 28) & _
                 Right(Space(5) & .Columns("QtdeSaida").Text, 5) & _
                 Right(Space(5) & .Columns("QtdeEntrada").Text, 5) & _
                 Right(Space(5) & (.Columns("QtdeSaida").Text - .Columns("QtdeEntrada").Text), 5) & _
                 Right(Space(10) & .Columns("Preco").Text, 10) & _
                 Right(Space(9) & .Columns("PrecoUnitario").Text, 9)
      .MoveNext
    Next
  End With
  
  LetPrinter ""
  LetPrinter " Totalizadores "
  LetPrinter "|----------------------------------------------------------------|"
  LetPrinter "| Consignado | Devolvido |  Saldo  |                      Total  |"
  LetPrinter "|" & Right(Space(12) & txtTotalConsignado.Text, 12) & _
             "|" & Right(Space(11) & txtTotalDevolvido.Text, 11) & _
             "|" & Right(Space(9) & txtTotalSaldo.Text, 9) & _
             "|" & Right(Space(29) & txtTotalPagar.Text, 29) & "|"
  
  LetPrinter "|----------------------------------------------------------------|"
  LetPrinter "|  Extrato das Consignacoes                                      |"
  LetPrinter "|----------------------------------------------------------------|"
  
  Close #lngbuffer
  
  Exit Sub
  
ErrView:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
    Exit Sub
    
End Sub

Private Sub Form_Load()
  Dim rstClientes As Recordset
  Dim lngConsignacaoMestre As Long
  Dim rstEntradas As Recordset
  Dim rstEnProdutos As Recordset
  Dim rstSaidas As Recordset
  Dim rstSaProdutos As Recordset
  Dim lngSequenciaEntrada As Long
  Dim lngSequenciaSaida   As Long
  Dim intX As Integer
  Dim blnExisteNoGrid As Boolean
  Dim dblPreco As Double
  Dim strCodProdSemGrade As String
  
  Call CenterForm(Me)
  
  Set rstClientes = db.OpenRecordset(" SELECT Código, UltimaConsignacao, ConsignacaoFechada, Nome " & _
                                     " FROM Cli_For WHERE Código = " & lngCodigoCliente, dbOpenSnapshot)
  
  With rstClientes
    If Not (.BOF And .EOF) Then
      lblNomeCliente.Caption = .Fields("Código") & " - " & .Fields("Nome")
      
      If .Fields("ConsignacaoFechada") Then Exit Sub
      lngConsignacaoMestre = IIf(IsNumeric(.Fields("UltimaConsignacao")), .Fields("UltimaConsignacao"), 0)
      
      Set rstSaidas = db.OpenRecordset(" SELECT * FROM Saídas " & _
                                         " WHERE ConsignacaoMestre = " & lngConsignacaoMestre & " AND Filial = " & gnCodFilial, dbOpenSnapshot)
      
      With rstSaidas
        If Not (.BOF And .EOF) Then
          .MoveFirst
          
          grdResumoConsignacoes.Redraw = False
          grdResumoConsignacoes.RemoveAll
          
          Do While Not .EOF
            lngSequenciaSaida = .Fields("Sequência")
            Set rstSaProdutos = db.OpenRecordset(" SELECT * FROM [Saídas - Produtos] WHERE Sequência = " & lngSequenciaSaida & " AND Filial = " & gnCodFilial, dbOpenSnapshot)

            With rstSaProdutos
              If Not (.BOF And .EOF) Then
                .MoveFirst
                
                Do While Not .EOF
                  blnExisteNoGrid = False
                  dblPreco = 0
                  
                  grdResumoConsignacoes.MoveFirst
                  For intX = 0 To grdResumoConsignacoes.Rows - 1
                    If UCase(grdResumoConsignacoes.Columns("Produto").Text) = UCase(.Fields("Código")) Then
                      blnExisteNoGrid = True
                      grdResumoConsignacoes.Columns("QtdeSaida").Text = grdResumoConsignacoes.Columns("QtdeSaida").Text + .Fields("Qtde")
                      grdResumoConsignacoes.Columns("Preco").Text = CDbl(grdResumoConsignacoes.Columns("Preco").Text) + CDbl(.Fields("Preço Final"))
                    End If
                    
                    grdResumoConsignacoes.MoveNext
                  Next intX
                  
                  If Not blnExisteNoGrid Then
                    grdResumoConsignacoes.AddNew
                    grdResumoConsignacoes.Columns("Produto").Text = UCase(.Fields("Código"))
                    '18/01/2005 - Daniel
                    'Tratamento para produtos que usam a Grade
                    Call BuscarCodProdSem(.Fields("Código"), strCodProdSemGrade)
                    grdResumoConsignacoes.Columns("ProdutoNome").Text = getNomeProduto(strCodProdSemGrade)
                    'Antiga linha comentada em 18/01/2005
                    'grdResumoConsignacoes.Columns("ProdutoNome").Text = GetNomeProduto(.Fields("Código"))
                    grdResumoConsignacoes.Columns("QtdeSaida").Text = .Fields("Qtde")
                    grdResumoConsignacoes.Columns("Preco").Text = CDbl(.Fields("Preço Final"))
                  End If
                  
                  grdResumoConsignacoes.Update
                  .MoveNext
                Loop
              End If

              If Not rstSaProdutos Is Nothing Then .Close
              Set rstSaProdutos = Nothing
            End With
            
            .MoveNext
          Loop
          
          grdResumoConsignacoes.Redraw = True
        End If
      End With
      
      '---------------------------------------------------------------------------------------'
      
      Set rstEntradas = db.OpenRecordset(" SELECT * FROM Entradas " & _
                                         " WHERE ConsignacaoMestre = " & lngConsignacaoMestre & " AND Filial = " & gnCodFilial, dbOpenSnapshot)

      With rstEntradas
        If Not (.BOF And .EOF) Then
          .MoveFirst

          grdResumoConsignacoes.Redraw = False

          Do While Not .EOF
            lngSequenciaEntrada = .Fields("Sequência")
            Set rstEnProdutos = db.OpenRecordset(" SELECT * FROM [Entradas - Produtos] WHERE Sequência = " & lngSequenciaEntrada & " AND Filial = " & gnCodFilial, dbOpenSnapshot)

            With rstEnProdutos
              If Not (.BOF And .EOF) Then
                .MoveFirst

                Do While Not .EOF
                  grdResumoConsignacoes.MoveFirst

                  For intX = 0 To grdResumoConsignacoes.Rows - 1
                    If UCase(grdResumoConsignacoes.Columns("Produto").Text) = UCase(rstEnProdutos.Fields("Código")) Then
                      blnExisteNoGrid = True
                      
                      If Not IsNumeric(grdResumoConsignacoes.Columns("QtdeEntrada").Text) Then grdResumoConsignacoes.Columns("QtdeEntrada").Text = 0
                      
                      grdResumoConsignacoes.Columns("QtdeEntrada").Text = grdResumoConsignacoes.Columns("QtdeEntrada").Text + rstEnProdutos.Fields("Qtde")
                      grdResumoConsignacoes.Columns("Saldo").Text = _
                        grdResumoConsignacoes.Columns("QtdeSaida").Text - _
                        grdResumoConsignacoes.Columns("QtdeEntrada").Text
                      
                      If Not IsNumeric(grdResumoConsignacoes.Columns("Preco").Text) Then grdResumoConsignacoes.Columns("Preco").Text = 0
                      
                      grdResumoConsignacoes.Columns("Preco").Text = CDbl(grdResumoConsignacoes.Columns("Preco").Text) - (.Fields("Preço Final") + (.Fields("Preço Final") * .Fields("IPI")))
                      
                      grdResumoConsignacoes.Update
                    End If

                    grdResumoConsignacoes.MoveNext
                  Next intX
                  .MoveNext
                Loop
              End If

              If Not rstEnProdutos Is Nothing Then .Close
              Set rstEnProdutos = Nothing
            End With

            .MoveNext
          Loop

          grdResumoConsignacoes.Redraw = True
        End If
      End With
    End If
  End With

  With grdResumoConsignacoes
    .MoveFirst

    Dim lngTotalConsignado As Long
    Dim lngTotalDevolvido As Long
    Dim lngTotalSaldo As Long

    For intX = 0 To .Rows - 1
      If Not IsNumeric(.Columns("QtdeEntrada").Text) Then
        .Columns("QtdeEntrada").Text = 0
        .Columns("Saldo").Text = .Columns("QtdeSaida").Text - .Columns("QtdeEntrada").Text
        .Update
      End If
      
      If Not IsNumeric(.Columns("Preco").Text) Then .Columns("Preco").Text = 0
      
      .Columns("Preco").Text = Format(CStr(.Columns("Preco").Text), FORMAT_VALUE)
      
      lngTotalConsignado = lngTotalConsignado + .Columns("QtdeSaida").Text
      lngTotalDevolvido = lngTotalDevolvido + .Columns("QtdeEntrada").Text
      lngTotalSaldo = lngTotalSaldo + .Columns("Saldo").Text
      dblPreco = dblPreco + .Columns("Preco").Text
      
      .MoveNext
    Next intX

    .MoveFirst
  End With

  '02/03/2005 - Daniel
  'Adicionada a coluna preço unitário
  'Solicitante: Aura Prata
  Call PreencherPrecoUnitario
  '----------------------------------
  
  txtTotalPagar.Text = Format(CStr(dblPreco), FORMAT_VALUE)
  txtTotalConsignado.Text = lngTotalConsignado
  txtTotalDevolvido.Text = lngTotalDevolvido
  txtTotalSaldo.Text = lngTotalSaldo
End Sub

Private Sub PreencherPrecoUnitario()
  '02/03/2005 - Daniel
  'Adicionada a coluna preço unitário
  'Solicitante: Aura Prata
  Dim intX            As Integer
  Dim strProdSemGrade As String
  Dim dblPrecoUnit    As Double
  'Nota: Vamos buscar o preço unitário da tabela
  'que estiver cadastrada no Parâmetros
  Dim rstParametros As Recordset
  Dim strTabela     As String
  
  Set rstParametros = db.OpenRecordset("SELECT Consignacao_TabelaPrecos FROM [Parâmetros Filial] WHERE Filial = " & gnCodFilial, dbOpenSnapshot)
  
  With rstParametros
    If Not (.BOF And .EOF) Then
      .MoveFirst
      strTabela = .Fields("Consignacao_TabelaPrecos").Value & ""
    End If
    .Close
  End With
  
  Set rstParametros = Nothing
  
  With grdResumoConsignacoes
    .MoveFirst
  
    For intX = 0 To .Rows - 1
      If Len(.Columns("Produto").Text) > 0 Then
        Call Acha_Produto(.Columns("Produto").Text, strProdSemGrade, 0, 0, 0, 0, 0)
        Call BuscarPrecoUnitario(strProdSemGrade, strTabela, dblPrecoUnit)
        
        .Columns("PrecoUnitario").Text = Format(dblPrecoUnit, FORMAT_VALUE)
      End If

      .MoveNext
    Next intX
    
    .MoveFirst
  End With
  
End Sub

Private Sub BuscarPrecoUnitario(ByVal CodProd As String, ByVal Tabela As String, ByRef PrecoUnit As Double)
  '02/03/2005 - Daniel
  'Adicionada a coluna preço unitário
  'Solicitante: Aura Prata
  Dim rstPrecos As Recordset
  Dim strSQL    As String
  
  PrecoUnit = 0
  
  strSQL = "SELECT * FROM Preços WHERE Produto = '" & CodProd & "'"
  strSQL = strSQL & " AND Tabela = '" & Tabela & "'"

  Set rstPrecos = db.OpenRecordset(strSQL, dbOpenSnapshot)

  With rstPrecos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      PrecoUnit = .Fields("Preço").Value
    End If
    .Close
  End With

  Set rstPrecos = Nothing

End Sub
Private Function getNomeProduto(strCodigo As String) As String
  Dim rstProdutos As Recordset

  Set rstProdutos = db.OpenRecordset(" SELECT Nome FROM Produtos " & _
                                     " WHERE Código = '" & strCodigo & "'", dbOpenSnapshot)

  With rstProdutos
    If (.BOF And .EOF) Then
      getNomeProduto = "<Produto_Sem_Nome>"
    Else
      getNomeProduto = .Fields("Nome") & ""
    End If

    If Not rstProdutos Is Nothing Then .Close
    Set rstProdutos = Nothing
  End With
End Function

Private Sub LetPrinter(strString As String, Optional blnSameLine As Boolean = False)
  If blnSameLine Then
    Print #lngbuffer, strString;
  Else
    Print #lngbuffer, strString
  End If
End Sub

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
