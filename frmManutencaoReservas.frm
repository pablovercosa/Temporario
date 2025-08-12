VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmManutencaoReservas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção de Reservas"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManutencaoReservas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   8820
   Begin VB.Frame fraY 
      Caption         =   "Estornar a Reserva"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2280
      TabIndex        =   12
      Top             =   4800
      Width           =   2120
      Begin VB.CommandButton cmdEstornar 
         Caption         =   "&Estornar"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame fraF 
      Caption         =   "Faturar a Reserva"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4440
      TabIndex        =   11
      Top             =   4800
      Width           =   4275
      Begin VB.TextBox txtOperacao 
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Width           =   3135
      End
      Begin VB.Data datOperacao 
         Caption         =   "datOperacao"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome, Tipo FROM [Operações Saída] ORDER BY Código"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.CommandButton cmdFaturar 
         Caption         =   "&Faturar"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperacao 
         Bindings        =   "frmManutencaoReservas.frx":058A
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   735
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Operação de Saída"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame fraX 
      Caption         =   "Ações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   2120
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdCarregar 
         BackColor       =   &H0000C0C0&
         Caption         =   "&Carregar"
         Default         =   -1  'True
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdResultado 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   8655
      ScrollBars      =   2
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
      Col.Count       =   5
      BevelColorFrame =   -2147483632
      BevelColorHighlight=   -2147483633
      BevelColorShadow=   -2147483633
      AllowRowSizing  =   0   'False
      RowHeight       =   423
      ExtraHeight     =   26
      Columns.Count   =   5
      Columns(0).Width=   2355
      Columns(0).Caption=   "Data"
      Columns(0).Name =   "Data"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   7
      Columns(0).FieldLen=   256
      Columns(1).Width=   2355
      Columns(1).Caption=   "Sequência"
      Columns(1).Name =   "Sequencia"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      Columns(2).Width=   4604
      Columns(2).Caption=   "Cliente"
      Columns(2).Name =   "Cliente"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2355
      Columns(3).Caption=   "Total"
      Columns(3).Name =   "Total"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   6
      Columns(3).NumberFormat=   "CURRENCY"
      Columns(3).FieldLen=   256
      Columns(4).Width=   2355
      Columns(4).Caption=   "Validade"
      Columns(4).Name =   "Validade"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   7
      Columns(4).FieldLen=   256
      _ExtentX        =   15266
      _ExtentY        =   4895
      _StockProps     =   79
      Caption         =   "Reservas Existentes"
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
   Begin VB.Frame fraResultados 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   9135
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ao clicar em Faturar, você estará a partir da Reserva, criando uma Saída, apta para emitir nota e realizar recebimento."
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   1440
         Width           =   7980
      End
      Begin VB.Label lblDica 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmManutencaoReservas.frx":05A4
         ForeColor       =   &H00808080&
         Height          =   795
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   7980
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Visualização das Reservas Cadastradas no Sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmManutencaoReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Para cálculo do Valor do ICM
Dim m_dblValorICM      As Double

Private Sub cboOperacao_CloseUp()
  cboOperacao.Text = cboOperacao.Columns(0).Text
  cboOperacao_LostFocus
End Sub

Private Sub cboOperacao_LostFocus()
  Dim rstOperacoesSaida As Recordset

  txtOperacao.Text = ""

  If Not IsNumeric(cboOperacao.Text) Then Exit Sub

  Set rstOperacoesSaida = db.OpenRecordset("SELECT Código, Nome, Tipo FROM [Operações Saída] WHERE Código = " & CInt(cboOperacao.Text) & " ORDER BY Código", dbOpenDynaset)

  With rstOperacoesSaida
    If Not (.BOF And .EOF) Then
      txtOperacao.Text = .Fields("Nome") & ""
    End If
  End With

  rstOperacoesSaida.Close
  Set rstOperacoesSaida = Nothing

End Sub

Private Sub cmdCarregar_Click()
    Dim rstSaidas As Recordset
    Dim strSQL    As String

    strSQL = " SELECT Data, Sequência, Cliente, Total, [Data Validade] "
    strSQL = strSQL & " FROM Saídas "
    strSQL = strSQL & " WHERE NOT IsNull(Saídas.[Data Validade]) "
    strSQL = strSQL & " AND NOT [Movimentação Desfeita] "
    strSQL = strSQL & " AND NOT Recebimento "
    strSQL = strSQL & " AND NOT FaturaSourceReserva " 'Posteriormente caso esta saída seja clonada para uma venda este se tornará TRUE
    strSQL = strSQL & " ORDER BY Saídas.Sequência "

    Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)

    With grdResultado
      'Não permite atualizar o layout do grid
      .Redraw = False
      'Limpa o grid
      .RemoveAll
      'Permite atualizar o layout do grid
      .Redraw = True
    End With

    With rstSaidas
      'Se o recordset estiver vazio
      If (.BOF And .EOF) Then
        MsgBox "Nenhuma reserva carregada.", vbInformation, "Quick Store"
        Exit Sub
      End If

      .MoveFirst

      Do Until .EOF
        grdResultado.AddNew

        grdResultado.Columns("Data").Text = .Fields("Data").Value
        grdResultado.Columns("Sequencia").Text = .Fields("Sequência").Value
        grdResultado.Columns("Cliente").Text = .Fields("Cliente").Value & " - " & GetNomeCliFor(.Fields("Cliente").Value)
        grdResultado.Columns("Total").Text = .Fields("Total").Value
        grdResultado.Columns("Validade").Text = .Fields("Data Validade").Value

        grdResultado.Update

      .MoveNext
      Loop

    End With

    rstSaidas.Close
    Set rstSaidas = Nothing

End Sub

Private Sub cmdEstornar_Click()
'Para cada linha selecionada na grid será realizado o estorno
'o campo Saídas.[Movimentação Desfeita] será atualizado para True
  Dim intAuxi        As Integer
  Dim varBook        As Variant
  Dim lngSeq         As Long
  Dim rstSaidas      As Recordset
  Dim strSQL         As String
  Dim blnTransaction As Boolean

  Dim rstEntradas    As Recordset
  Dim strSQLEntradas As String
  Dim lngSeqEntradas As Long

  Dim rstOperacoesEntrada    As Recordset

  Dim rstEntradasProdutos    As Recordset
  Dim lngSeqEntradasProdutos As Long
  Dim blnSeq                 As Boolean 'Amarrará a Seq para sucessivas criações em Entradas - Produtos

  Dim rstSaidasProdutos      As Recordset

  'Tratamento para Parâmetros e a função Efetiva_Entrada
  Dim rstParametros          As Recordset
  Dim nSequencia             As Long
  Dim nRet                   As Integer

  If ExaminaSelecao Then Exit Sub

  On Error GoTo ErrHandler

  'Inicia a Transação
  ws.BeginTrans
  blnTransaction = True

  Screen.MousePointer = vbHourglass

  For intAuxi = 0 To (grdResultado.SelBookmarks.Count - 1)
    varBook = grdResultado.SelBookmarks(intAuxi)
    grdResultado.Bookmark = varBook
    lngSeq = grdResultado.Columns("Sequencia").CellValue(book)

    strSQL = " SELECT * "
    strSQL = strSQL & " FROM Saídas "
    strSQL = strSQL & " WHERE Saídas.Sequência =" & CLng(lngSeq)
    strSQL = strSQL & " AND Saídas.Filial =" & CByte(gnCodFilial)

    Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)

'Colocado este código mais para o final...
'
'    With rstSaidas
'      If Not (.BOF And .EOF) Then
'        .Edit
'        .Fields("Movimentação Desfeita").Value = True
'        .Update
'      End If
'    End With

    'Próximo passo será a criação da entradas e em seguida entrada - produtos
      strSQLEntradas = " SELECT * FROM Entradas ORDER BY Sequência "

      Set rstEntradas = db.OpenRecordset(strSQLEntradas, dbOpenDynaset)

        With rstEntradas
          .MoveLast
          .MoveFirst

          .AddNew
          .Fields("Filial").Value = rstSaidas.Fields("Filial").Value
          .Fields("Data").Value = rstSaidas.Fields("Data").Value

          '----- [rstOperacoesEntrada]-----
          Set rstOperacoesEntrada = db.OpenRecordset(" SELECT Código, Estorno FROM [Operações Entrada] WHERE Estorno = True ")

          With rstOperacoesEntrada
            If Not (.BOF And .EOF) Then
              .MoveFirst
              rstEntradas.Fields("Operação").Value = .Fields("Código").Value
            Else

              rstOperacoesEntrada.Close
              rstSaidas.Close
              Set rstOperacoesEntrada = Nothing
              Set rstSaidas = Nothing

              MsgBox "Impossível criar Entrada, não há uma Operação de Entrada que possua o campo (Realiza Estorno da Reserva) marcado.", vbExclamation, "Quick Store"
              Exit Sub
            End If
          End With
          rstOperacoesEntrada.Close
          Set rstOperacoesEntrada = Nothing
          '----- Fim de rstOperacoesEntrada -----

          nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("Última Movimentação") + 1

          .Fields("Sequência").Value = nSequencia
          .Fields("Digitador").Value = rstSaidas.Fields("Digitador").Value
          .Fields("Fornecedor").Value = rstSaidas.Fields("Cliente").Value
          .Fields("Observações").Value = "Estorno de Reserva " & " - " & "Sequência da Saída: " & rstSaidas.Fields("Sequência").Value & " - " & "Dia: " & rstSaidas.Fields("Data").Value
          '.Fields("Nota Fiscal").Value =
          .Fields("Data Emissão").Value = Data_Atual
          '.Fields("Pedido").Value =
          '.Fields("Forma Pagto").Value =
          .Fields("Produtos").Value = rstSaidas.Fields("Produtos").Value
          .Fields("Desconto").Value = rstSaidas.Fields("Desconto").Value
          .Fields("IPI").Value = rstSaidas.Fields("IPI").Value
          .Fields("Frete").Value = rstSaidas.Fields("Frete").Value
          .Fields("Base ICM").Value = rstSaidas.Fields("Base ICM").Value
          .Fields("Valor ICM").Value = rstSaidas.Fields("Valor ICM").Value
          .Fields("Base ICM Subs").Value = rstSaidas.Fields("Base ICM Subs").Value
          .Fields("Valor ICM Subs").Value = rstSaidas.Fields("Valor ICM Subs").Value
          .Fields("Total").Value = rstSaidas.Fields("Total").Value
          .Fields("Dinheiro Caixa").Value = 0
          .Fields("Cheque Caixa").Value = 0
          .Fields("Caixa").Value = 0
          .Fields("Conta").Value = 0
          .Fields("Num Cheque").Value = ""
          '.Fields("Bom para").Value =
          .Fields("Valor Cheque").Value = 0
          .Fields("Descrição").Value = "Estorno"
          .Fields("Efetivada").Value = True
          .Fields("Nota Impressa").Value = 0
          .Fields("Nota Cancelada").Value = False
          '.Fields("Data Acerto Empréstimo").Value =
          '.Fields("WebOrderFormID").Value =
          '.Fields("Centro Custo").Value =
          '.Fields("ConsignacaoMestre").Value =
          .Fields("obs_infCpl1").Value = ""
          .Fields("obs_infCpl2").Value = ""
'          .Fields("obs_Obs1").Value = ""
'          .Fields("obs_Obs2").Value = ""
'          .Fields("obs_Obs3").Value = ""
'          .Fields("obs_Obs4").Value = ""
'          .Fields("obs_Obs5").Value = ""
'          .Fields("obs_Obs6").Value = ""
'          .Fields("obs_Obs7").Value = ""
'          .Fields("obs_Obs8").Value = ""
          .Fields("obs_Transportadora").Value = ""
          .Fields("obs_Placa").Value = ""
          .Fields("obs_UF").Value = ""
          .Fields("obs_Qtde").Value = ""
          .Fields("obs_Especie").Value = ""
          .Fields("obs_Marca").Value = ""
          .Fields("obs_PesoLiquido").Value = 0
          .Fields("obs_PesoBruto").Value = 0
          .Fields("obs_FretePago").Value = 0

          .Update

        End With 'rstEntradas
        rstEntradas.Close
        Set rstEntradas = Nothing

        '----- [rstSaidasProdutos]-----
        Set rstSaidasProdutos = db.OpenRecordset(" SELECT * FROM [Saídas - Produtos] WHERE [Saídas - Produtos].Sequência=" & CLng(lngSeq) & " AND [Saídas - Produtos].Filial=" & CByte(gnCodFilial) & " ORDER BY [Saídas - Produtos].Linha ", dbOpenDynaset)

          With rstSaidasProdutos
            .MoveFirst

            If Not (.BOF And .EOF) Then
              Do Until .EOF
                '----- [rstSaidasProdutos]-----
                Set rstEntradasProdutos = db.OpenRecordset(" SELECT * FROM [Entradas - Produtos] ", dbOpenDynaset)

                With rstEntradasProdutos
                  .MoveLast
                  .MoveFirst

                  lngSeqEntradasProdutos = 0

                  Do Until .EOF
                    If lngSeqEntradasProdutos < .Fields("Sequência").Value Then
                      lngSeqEntradasProdutos = .Fields("Sequência").Value
                    End If
                  .MoveNext
                  Loop

                  If Not blnSeq Then lngSeqEntradasProdutos = lngSeqEntradasProdutos + 1
                  blnSeq = True 'Passou uma não passará mais vezes...

                  .AddNew
                  .Fields("Filial").Value = rstSaidasProdutos.Fields("Filial").Value
                  .Fields("Sequência").Value = nSequencia  'lngSeqEntradasProdutos
                  .Fields("Linha").Value = rstSaidasProdutos.Fields("Linha").Value
                  .Fields("Código").Value = rstSaidasProdutos.Fields("Código").Value
                  .Fields("Qtde").Value = rstSaidasProdutos.Fields("Qtde").Value
                  .Fields("Preço").Value = rstSaidasProdutos.Fields("Preço").Value
                  .Fields("Desconto").Value = rstSaidasProdutos.Fields("Desconto").Value
                  .Fields("ICM").Value = rstSaidasProdutos.Fields("ICM").Value
                  .Fields("IPI").Value = rstSaidasProdutos.Fields("IPI").Value
                  .Fields("Preço Final").Value = rstSaidasProdutos.Fields("Preço Final").Value
                  .Fields("Etiqueta").Value = rstSaidasProdutos.Fields("Etiqueta").Value
                  .Fields("Código Sem Grade").Value = rstSaidasProdutos.Fields("Código sem Grade").Value
                  .Fields("InGeradoViaConsig").Value = False

                  .Update

                End With
                rstEntradasProdutos.Close
                Set rstEntradasProdutos = Nothing

              rstSaidasProdutos.MoveNext
              Loop 'rstSaidasProdutos
            End If
          End With
          'Finalizar o rstSaidasProdutos
          rstSaidasProdutos.Close
          Set rstSaidasProdutos = Nothing


          'Chamada da função do Quick Efetiva_Entrada para Efetivação
          Call StatusMsg("Efetivando entrada...")
            nRet = Efetiva_Entrada(gnCodFilial, nSequencia)
            If nRet <> 0 Then
              Select Case nRet
                Case -1 'Ação cancelada
                  Call StatusMsg("Ação cancelada.")
                Case 1
                  Call DisplayMsg("Código da operação inexistente.")
                Case 2
                  Call DisplayMsg("Funcionário inexistente.")
                Case 3
                  Call DisplayMsg("Fornecedor inexistente.")
                Case Else
                  Call DisplayMsg("Operação NÃO efetivada. Erro" & str(nRet))
              End Select
              Screen.MousePointer = vbDefault
              blnTransaction = True
              'ws.Rollback
              Call StatusMsg("")
              Exit Sub
            Else
              '-----[Tudo OK]-----

              'Agora podemos atualizar a tabela Saídas
              With rstSaidas
                If Not (.BOF And .EOF) Then
                  .Edit
                  .Fields("Movimentação Desfeita").Value = True
                  .Update
                End If
              End With
              '-------------------------------------------------

              'Tratamento Parâmetros
              Set rstParametros = db.OpenRecordset(" SELECT [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial =" & gnCodFilial, dbOpenDynaset)

                rstParametros.Edit
                rstParametros.Fields("Última Movimentação").Value = nSequencia
                rstParametros.Update
                rstParametros.Close

              Set rstParametros = Nothing
              '----- Fim do Tratamento Parâmetros -----

              'Call ws.CommitTrans
              Call StatusMsg("Movimentação de Entrada realizada com sucesso.")
            End If

          Call StatusMsg("")
          '----- Fim Efetivar a Entrada


  Next intAuxi

  Screen.MousePointer = vbDefault

  'Finaliza Transação
  ws.CommitTrans
  blnTransaction = False

  rstSaidas.Close
  Set rstSaidas = Nothing

  MsgBox "Estorno realizado com sucesso.", vbInformation, "Quick Store"

  cmdCarregar_Click

  Exit Sub

ErrHandler:
  'Desfaz a transação
  If blnTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmdFaturar_Click()
  Dim intAuxi                As Integer
  Dim varBook                As Variant
  Dim lngSeq                 As Long
  Dim rstSaidas              As Recordset
  Dim strSQL                 As String
  Dim rstSaidasClone         As Recordset
  Dim rstSaidasProdutos      As Recordset
  Dim rstSaidasProdutosClone As Recordset
  Dim blnTransaction         As Boolean

  'Tratamento para Parâmetros e a função Efetiva_Entrada
  Dim rstParametros          As Recordset
  Dim nSequencia             As Long
  Dim nRet                   As Integer


  '-------------------------[Validações]----------------------------------------
  If ExaminaSelecao Then Exit Sub

  If Len(txtOperacao.Text) <= 0 Then
    MsgBox "Operação de Saída inválida, verifique.", vbExclamation, "Quick Store"
    cboOperacao.SetFocus
    Exit Sub
  End If

  'Verificar se [Operações Saída].Estoque está True,
  'não poderemos baixar duas vezes
  If VerificaOperacoesSaidaEstoque Then
    MsgBox "Selecione uma Operação que esteja desmarcado o campo 'Diminui Estoque' em 'Operações de Saída' para não ocorrer baixa duas vezes no Estoque.", vbExclamation, "Quick Store"
    cboOperacao.SetFocus
    Exit Sub
  End If
  '-------------------------[Fim Validações]------------------------------------

  On Error GoTo ErrHandler

  '-------------------------[Inicia a Transação]--------------------------------
  ws.BeginTrans
  blnTransaction = True

  Screen.MousePointer = vbHourglass

  For intAuxi = 0 To (grdResultado.SelBookmarks.Count - 1)
    varBook = grdResultado.SelBookmarks(intAuxi)
    grdResultado.Bookmark = varBook
    lngSeq = grdResultado.Columns("Sequencia").CellValue(book)

    strSQL = " SELECT * "
    strSQL = strSQL & " FROM Saídas "
    strSQL = strSQL & " WHERE Saídas.Sequência =" & CLng(lngSeq)
    strSQL = strSQL & " AND Saídas.Filial =" & CByte(gnCodFilial)

    Set rstSaidas = db.OpenRecordset(strSQL, dbOpenDynaset)

    Set rstSaidasClone = db.OpenRecordset("Saídas", dbOpenDynaset)

    'Buscar uma próxima Sequência
    nSequencia = gnGetNextSequencia(gnCodFilial) 'rsParametros("Última Movimentação") + 1

    With rstSaidas
      If Not (.BOF And .EOF) Then
        .MoveFirst

        .Edit
        .Fields("FaturaSourceReserva").Value = True 'Para não misturar-se com os registros do botão carregar
        .Update

        'Criação da nova Saída [Populado somente os campos necessários no momento]
        rstSaidasClone.AddNew
          rstSaidasClone.Fields("Filial").Value = .Fields("Filial").Value
          rstSaidasClone.Fields("Data").Value = .Fields("Data").Value
          rstSaidasClone.Fields("Sequência").Value = nSequencia
          rstSaidasClone.Fields("Operação").Value = CInt(cboOperacao.Text)
          rstSaidasClone.Fields("Caixa").Value = .Fields("Caixa").Value
          rstSaidasClone.Fields("Tabela").Value = .Fields("Tabela").Value
          rstSaidasClone.Fields("Digitador").Value = .Fields("Digitador").Value
          rstSaidasClone.Fields("Operador").Value = .Fields("Operador").Value
          rstSaidasClone.Fields("Cliente").Value = .Fields("Cliente").Value
          rstSaidasClone.Fields("Observações").Value = "Saída oriunda da Saída de Reserva de Sequência " & .Fields("Sequência").Value
          rstSaidasClone.Fields("Produtos").Value = .Fields("Produtos").Value
          rstSaidasClone.Fields("Desconto").Value = .Fields("Desconto").Value
          rstSaidasClone.Fields("Serviços").Value = .Fields("Serviços").Value
          rstSaidasClone.Fields("Base ISS").Value = .Fields("Base ISS").Value
          rstSaidasClone.Fields("Valor ISS").Value = .Fields("Valor ISS").Value
          rstSaidasClone.Fields("Perc IR Sobre ISS").Value = .Fields("Perc IR Sobre ISS").Value
          rstSaidasClone.Fields("Valor IR Sobre ISS").Value = .Fields("Valor IR Sobre ISS").Value
          rstSaidasClone.Fields("IPI").Value = .Fields("IPI").Value
          rstSaidasClone.Fields("Frete").Value = .Fields("Frete").Value
          rstSaidasClone.Fields("Base ICM").Value = .Fields("Total").Value
          'rstSaidasClone.Fields("Valor ICM").Value [Vamos popular mais adiante]
          rstSaidasClone.Fields("Base ICM Subs").Value = .Fields("Base ICM Subs").Value
          rstSaidasClone.Fields("Valor ICM Subs").Value = .Fields("Valor ICM Subs").Value
          rstSaidasClone.Fields("Total").Value = .Fields("Total").Value
          rstSaidasClone.Fields("Efetivada").Value = False
          rstSaidasClone.Fields("Recebimento").Value = False
          rstSaidasClone.Fields("FaturaSourceReserva").Value = True 'Para não misturar-se com os registros do botão carregar

            'Criação da [Saídas - Produtos]
            Set rstSaidasProdutos = db.OpenRecordset("SELECT * FROM [Saídas - Produtos] WHERE [Saídas - Produtos].Sequência =" & (.Fields("Sequência").Value), dbOpenDynaset)
            Set rstSaidasProdutosClone = db.OpenRecordset("Saídas - Produtos", dbOpenDynaset)
            
            If Not (rstSaidasProdutos.BOF And rstSaidasProdutos.EOF) Then
              rstSaidasProdutos.MoveFirst
              
              Do Until rstSaidasProdutos.EOF
                
                'Rotina de Cálculo do ICM [CodProduto - Preço Final]
                Call CalculoICM(rstSaidasProdutos.Fields("Código").Value, rstSaidasProdutos.Fields("Preço Final").Value)
                
                rstSaidasProdutosClone.AddNew
                  rstSaidasProdutosClone.Fields("Filial").Value = rstSaidasProdutos.Fields("Filial").Value
                  rstSaidasProdutosClone.Fields("Sequência").Value = nSequencia
                  rstSaidasProdutosClone.Fields("Linha").Value = rstSaidasProdutos.Fields("Linha").Value
                  rstSaidasProdutosClone.Fields("Código").Value = rstSaidasProdutos.Fields("Código").Value
                  rstSaidasProdutosClone.Fields("Qtde").Value = rstSaidasProdutos.Fields("Qtde").Value
                  rstSaidasProdutosClone.Fields("Preço").Value = rstSaidasProdutos.Fields("Preço").Value
                  rstSaidasProdutosClone.Fields("Desconto").Value = rstSaidasProdutos.Fields("Desconto").Value
                  rstSaidasProdutosClone.Fields("Desconto Valor").Value = rstSaidasProdutos.Fields("Desconto Valor").Value
                  rstSaidasProdutosClone.Fields("ICM").Value = rstSaidasProdutos.Fields("ICM").Value
                  rstSaidasProdutosClone.Fields("IPI").Value = rstSaidasProdutos.Fields("IPI").Value
                  rstSaidasProdutosClone.Fields("Preço Final").Value = rstSaidasProdutos.Fields("Preço Final").Value
                  rstSaidasProdutosClone.Fields("Etiqueta").Value = rstSaidasProdutos.Fields("Etiqueta").Value
                  rstSaidasProdutosClone.Fields("Código sem Grade").Value = rstSaidasProdutos.Fields("Código sem Grade").Value
                  rstSaidasProdutosClone.Fields("InGeradoViaConsig").Value = rstSaidasProdutos.Fields("InGeradoViaConsig").Value
                  rstSaidasProdutosClone.Fields("Situação Tributária").Value = rstSaidasProdutos.Fields("Situação Tributária").Value
                  rstSaidasProdutosClone.Fields("Unidade Venda").Value = rstSaidasProdutos.Fields("Unidade Venda").Value
                  rstSaidasProdutosClone.Fields("Descricao Adicional").Value = rstSaidasProdutos.Fields("Descricao Adicional").Value
                  rstSaidasProdutosClone.Fields("QtdeEntregue").Value = rstSaidasProdutos.Fields("QtdeEntregue").Value
                rstSaidasProdutosClone.Update
                
                rstSaidasProdutos.MoveNext
              Loop
    
            End If
            rstSaidasProdutos.Close
            rstSaidasProdutosClone.Close
            Set rstSaidasProdutos = Nothing
            Set rstSaidasProdutosClone = Nothing
            'Fim da Criação da [Saídas - Produtos]

        rstSaidasClone.Fields("Valor ICM").Value = m_dblValorICM
        rstSaidasClone.Update
        'Fim da Criação da nova Saída...

      End If
      
    End With 'rstSaidas
    rstSaidas.Close
    rstSaidasClone.Close
    Set rstSaidas = Nothing
    Set rstSaidasClone = Nothing
    
    'Efetivar Saída
    Call StatusMsg("Aguarde, efetivando venda...")

    nRet = Efetiva_Saída(gnCodFilial, nSequencia)

    If nRet <> 0 Then
      Select Case nRet
        Case -1
          'Ação cancelada
          Call StatusMsg("Ação cancelada.")
        Case 5
          Call DisplayMsg("Tabela de preços inexistente.")
        Case Else
          Call DisplayMsg("Operação NÃO efetivada. Erro" & str(nRet))
      End Select
      'Cancelamento da transação
      ws.Rollback
      Exit Sub
    End If
    'Fim de Efetivar Saída
    
    'Tratamento para Atualização de Parâmetros
    Set rstParametros = db.OpenRecordset(" SELECT [Última Movimentação] FROM [Parâmetros Filial] WHERE Filial =" & gnCodFilial, dbOpenDynaset)

      rstParametros.Edit
      rstParametros.Fields("Última Movimentação").Value = nSequencia
      rstParametros.Update
      rstParametros.Close

    Set rstParametros = Nothing
    'Fim do Tratamento para Atualização de Parâmetros
    
    m_dblValorICM = 0
    
  Next intAuxi
      
  'Fim da transação
  ws.CommitTrans
  blnTransaction = False
  Call StatusMsg("")

  Screen.MousePointer = vbDefault
  MsgBox "Gerado saída(s) para faturamento com sucesso.", vbInformation, "Quick Store"
  cboOperacao.Text = ""
  txtOperacao.Text = ""
  cmdCarregar_Click

  Exit Sub

ErrHandler:
  'Desfaz a transação
  If blnTransaction Then ws.Rollback
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"

End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    datOperacao.DatabaseName = gsQuickDBFileName

    Call CenterForm(Me)
End Sub

Private Function GetNomeCliFor(lngCodigo As Long) As String
  Dim rstCliFor As Recordset

  Set rstCliFor = db.OpenRecordset("SELECT Nome FROM Cli_For WHERE Código = " & lngCodigo, dbOpenDynaset)

  With rstCliFor
    GetNomeCliFor = IIf((.BOF And .EOF), "<_não_cadastrado>", .Fields("Nome").Value & "")
    .Close
  End With


  Set rstCliFor = Nothing
End Function

Private Sub grdResultado_DblClick()
'Ao dar um duplo click na linha da grid desejada,
'carregará o registro na tela de saidas

    With frmSaidas
      .txtSeq.Text = grdResultado.Columns("Sequencia").Text
      .SearchRecord
      .Show
    End With

End Sub

Private Function VerificaOperacoesSaidaEstoque() As Boolean
  Dim rstOperacoesSaida As Recordset

  Set rstOperacoesSaida = db.OpenRecordset("SELECT Estoque FROM [Operações Saída] WHERE [Operações Saída].Código =" & CInt(cboOperacao.Text), dbOpenDynaset)

  With rstOperacoesSaida
    If Not (.BOF And .EOF) Then
      VerificaOperacoesSaidaEstoque = .Fields("Estoque").Value
    End If
    .Close
  End With

  Set rstOperacoesSaida = Nothing

End Function
Private Function ExaminaSelecao() As Boolean
  If grdResultado.SelBookmarks.Count < 1 Then
    MsgBox "Favor selecionar alguma saída da grid.", vbExclamation, "Quick Store"
    ExaminaSelecao = True
  End If
End Function

Private Function CalculoICM(ByVal CodProduto As String, PrecoFinal As Single) As String
  Dim rstProdutos   As Recordset
  Dim sngPercentual As Single
  Dim dblAuxi       As Double
  
  Set rstProdutos = db.OpenRecordset("SELECT [Percentual Icm Saida] FROM Produtos WHERE Código ='" & CodProduto & "'", dbOpenDynaset)

  With rstProdutos
    If Not (.BOF And .EOF) Then
      sngPercentual = .Fields("Percentual Icm Saida").Value
    End If
    .Close
  End With

  Set rstProdutos = Nothing
  
  If sngPercentual <> 0 Then
    dblAuxi = (PrecoFinal * sngPercentual) / 100
    dblAuxi = Format(dblAuxi, "##,###,###.##")
  Else
    dblAuxi = 0
  End If
  
  m_dblValorICM = m_dblValorICM + dblAuxi
  
  dblAuxi = 0

End Function
