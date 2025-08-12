VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAlteracaoPrecoCusto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Preços Calculado"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmAlteracaoPrecoCusto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   8040
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   -360
      TabIndex        =   17
      Top             =   -120
      Width           =   8535
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAlteracaoPrecoCusto.frx":058A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alteração de preços calculado"
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
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   600
         Picture         =   "frmAlteracaoPrecoCusto.frx":0627
         Top             =   320
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdPesquisa 
      Caption         =   "&Pesquisa"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordenação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
      Begin VB.OptionButton optOrdenaCodigo 
         Caption         =   "Por Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   280
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optOrdenaNome 
         Caption         =   "Por Nome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   6
         Top             =   280
         Width           =   1215
      End
   End
   Begin VB.Data datSubClasse 
      Caption         =   "datSubClasse"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Sub Classes]"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Data datClasse 
      Caption         =   "datClasse"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM Classes"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Data datTabelaDestino 
      Caption         =   "datTabelaDestino"
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
      RecordSource    =   "SELECT DISTINCTROW Tabela FROM [Tabela de Preços]"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdAlterarPrecos 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Alterar Preços"
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
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBGrid grdProdutos 
      Bindings        =   "frmAlteracaoPrecoCusto.frx":0A27
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   7815
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   3
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "Produto"
      Columns(0).Name =   "Produto"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   6350
      Columns(1).Caption=   "Nome"
      Columns(1).Name =   "Nome"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Caption=   "Preço Lista"
      Columns(2).Name =   "PrecoLista"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   13785
      _ExtentY        =   4260
      _StockProps     =   79
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
   Begin VB.Frame Frame1 
      Height          =   1400
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   7815
      Begin VB.TextBox txtSubClasse 
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   960
         Width           =   4215
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmAlteracaoPrecoCusto.frx":0A42
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2778
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.TextBox txtClasse 
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   4215
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmAlteracaoPrecoCusto.frx":0A5D
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   1575
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2778
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.TextBox txtFabricante 
         Height          =   285
         Left            =   4440
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin SSDataWidgets_B.SSDBCombo cboTabelaDestino 
         Bindings        =   "frmAlteracaoPrecoCusto.frx":0A75
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         DataFieldList   =   "Tabela"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2778
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Tabela"
      End
      Begin VB.Label Label4 
         Caption         =   "Sub classe:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Classe:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Fabricante:"
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
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tabela de Destino:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   255
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAlteracaoPrecoCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------------------------
  Dim rsResultado As Recordset
  Dim rsPrecos    As Recordset
  
  Dim dblCustoPrecoValor      As Double       'Preço de Lista sem IPI
  Dim dblVendaPrecoValor      As Single       'Preço de venda sem os impostos
  Dim sngCustoDesconto        As Single       'Desconto dado pelo fornecedor
  Dim sngCustoFrete           As Single       'Frete
  Dim sngCustoICMSCompra      As Single       'ICMS de Compra
  Dim sngCustoIPICompra       As Single       'IPI de Compra
  Dim sngCustoCustoFinanceiro As Single       'Custo Financeiro
  Dim sngCustoOutrosCompra    As Single       'Outros Custos na compra
  Dim sngCustoPercSemImpostos As Single       'Percentual de produtos comprado sem nota
  Dim sngVendaPercSemImpostos As Single       'Percentual de produtos que serão vendidos sem nota
  
  Dim dblPrecoCustoCalculado  As Single       'Preço de custo calculado
  Dim dblPrecoVendaCalculado  As Single       'Preço de venda calculado
  
  Dim sngVendaICMS            As Single       'Percentual de ICMS para venda
  Dim sngVendaIPI             As Single       'Percentual de IPI para venda
  Dim sngVendaOutrosImpostos  As Single       'Outros Impostos para venda
  Dim sngVendaOutros          As Single       'Outros Valores da venda

  Dim sngMargemLucro          As Single       'Margem de lucro para cálculo do preço de venda
  Dim sngSaldoICMS            As Single       'Saldo para cálculo do preço de venda
'-------------------------------------------------------------------------------------------------

Private Sub cboClasse_CloseUp()
  cboClasse.Text = cboClasse.Columns(1).Text
End Sub

Private Sub cboClasse_LostFocus()
  Dim rsTMPClasse As Recordset
  
  txtClasse.Text = ""
  If Not IsNumeric(cboClasse.Text) Then Exit Sub
  Set rsTMPClasse = db.OpenRecordset("SELECT * FROM Classes WHERE Código = " & cboClasse.Text, dbOpenSnapshot)
  
  With rsTMPClasse
    If Not (.BOF And .EOF) Then
      txtClasse.Text = !Nome & ""
    End If
  End With
  
  If Not rsTMPClasse Is Nothing Then rsTMPClasse.Close
  Set rsTMPClasse = Nothing
End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.Text = cboSubClasse.Columns(1).Text
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim rsTMPSubClasse As Recordset
  
  txtSubClasse.Text = ""
  If Not IsNumeric(cboSubClasse.Text) Then Exit Sub
  Set rsTMPSubClasse = db.OpenRecordset("SELECT * FROM [Sub Classes] WHERE Código = " & cboClasse.Text, dbOpenSnapshot)
  
  With rsTMPSubClasse
    If Not (.BOF And .EOF) Then
      txtSubClasse.Text = !Nome & ""
    End If
  End With
  
  If Not rsTMPSubClasse Is Nothing Then rsTMPSubClasse.Close
  Set rsTMPSubClasse = Nothing
End Sub

Private Sub cmdAlterarPrecos_Click()
  Dim intCont As Integer
  
  If Len(Trim(cboTabelaDestino.Text)) <= 0 Then
    MsgBox "Tabela de destino inválida, verifique !!", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  grdProdutos.MoveFirst
  
  For intCont = 0 To grdProdutos.Rows - 1
    '------------------------------------
      With rsResultado
        .FindFirst "Código = '" & grdProdutos.Columns("Produto").Text & "'"
        
        If Not .NoMatch Then
          Set rsPrecos = db.OpenRecordset(" SELECT * FROM [Preços] WHERE Produto = '" & grdProdutos.Columns("Produto").Text & "'")
          
          .Edit                                                         'Coloca o registro do produto em modo de edição para atualizar os valores
          ![Custo Preço Valor] = grdProdutos.Columns("PrecoLista").Text 'Atualiza o novo [preço de Lista sem IPI]
          
          dblCustoPrecoValor = grdProdutos.Columns("PrecoLista").Text   'Pega o preço de lista digitado no GRID
          
          sngCustoDesconto = IIf(UCase(![Custo Desconto Fixo]) = UCase("V"), ![Custo Desconto Valor], (![Custo Desconto Perc] / 100) * dblCustoPrecoValor)
          dblCustoPrecoValor = dblCustoPrecoValor - sngCustoDesconto
          
          sngCustoFrete = IIf(UCase(![Custo Frete Fixo]) = UCase("V"), ![Custo Frete Valor], (![Custo Frete Perc] / 100) * dblCustoPrecoValor)
          
          sngCustoPercSemImpostos = (1 - (![Custo Perc Compra Sem] / 100))
          sngVendaPercSemImpostos = (1 - (![Custo Perc Venda Sem] / 100))
          
          sngCustoICMSCompra = IIf(UCase(![Custo ICM Compra Fixo]) = UCase("V"), ![Custo ICM Compra Valor], (![Custo ICM Compra Perc] / 100) * dblCustoPrecoValor)
'           sngCustoICMSCompra = sngCustoPercCompSemImpostos * sngCustoICMSCompra
          
          sngCustoIPICompra = IIf(UCase(![Custo IPI Compra Fixo]) = UCase("V"), ![Custo IPI Compra Valor], (![Custo IPI Compra Perc] / 100) * dblCustoPrecoValor)
'          sngCustoIPICompra = sngCustoPercCompSemImpostos * sngCustoIPICompra
          
          sngCustoCustoFinanceiro = IIf(UCase(![Custo Custo Finan Fixo]) = UCase("V"), ![Custo Custo Finan Valor], (![Custo Custo Finan Perc] / 100) * dblCustoPrecoValor)
          sngCustoOutrosCompra = IIf(UCase(![Custo Outros Compra Fixo]) = UCase("V"), ![Custo Outros Compra Valor], (![Custo Outros Compra Perc] / 100) * dblCustoPrecoValor)
          
          dblPrecoCustoCalculado = dblCustoPrecoValor + _
                                   sngCustoFrete + _
                                   sngCustoIPICompra + _
                                   sngCustoCustoFinanceiro + _
                                   sngCustoOutrosCompra
          
          '---[ Altera o preço de custo do produto ]---'
            rsPrecos.FindFirst "Tabela = 'CUSTO'"
            
            If rsPrecos.NoMatch Then
              rsPrecos.AddNew
            Else
              rsPrecos.Edit
            End If
            rsPrecos.Fields("Preço") = dblPrecoCustoCalculado
            ![Custo Custo Calculado] = dblPrecoCustoCalculado
            rsPrecos.Update
          '---[ Altera o preço de custo do produto ]---'
          
          
          '===[ Cálculo do preço de venda do produto ]==='
            
            'Select Case UCase(.Fields("Custo Manter"))
            '  Case UCase("P")
                Dim nCalculoGeral As Double
                Dim nAliquotas    As Single
                
                '---[ Cálculo do preço de venda por Percentual ]---'
                  sngMargemLucro = ![Custo Lucro Perc] / 100
                  dblPrecoVendaCalculado = dblPrecoCustoCalculado + (dblPrecoCustoCalculado * sngMargemLucro)
                  
                  nCalculoGeral = dblPrecoVendaCalculado - sngCustoICMSCompra
                  nAliquotas = (.Fields("Custo ICM Venda Perc") / 100) + (.Fields("Custo Impostos Perc") / 100) + (.Fields("Custo Outros Venda Perc") / 100)
                  
                  nAliquotas = 1 - nAliquotas
                  
                  dblPrecoVendaCalculado = nCalculoGeral / nAliquotas
                '---[ Cálculo do preço de venda por Percentual ]---'
            'End Select
          '===[ Cálculo do preço de venda do produto ]==='
            
'                        '                          '          sngVendaICMS = IIf(UCase(![Custo ICM Venda Fixo]) = UCase("V"), ![Custo ICM Venda Valor], (![Custo ICM Venda Perc] / 100) * dblPrecoVendaCalculado)
'                        '                          '          sngVendaIPI = IIf(UCase(![Custo IPI Venda Fixo]) = UCase("V"), ![Custo IPI Venda Valor], (![Custo IPI Venda Perc] / 100) * dblPrecoVendaCalculado)
'                        '                          '          sngVendaOutrosImpostos = IIf(UCase(![Custo Impostos Fixo]) = UCase("V"), ![Custo Impostos Valor], (![Custo Impostos Perc] / 100) * dblPrecoVendaCalculado)
'                        '                          '          sngVendaOutros = IIf(UCase(![Custo Outros Venda Fixo]) = UCase("V"), ![Custo Outros Venda Valor], (![Custo Outros Venda Perc] / 100) * dblPrecoVendaCalculado)
'                        '                          '
'                        '                          '          sngSaldoICMS = (sngVendaICMS * sngVendaPercSemImpostos) - (sngCustoICMSCompra * sngCustoPercSemImpostos)
'                        '                          '
'                        '                          '          dblPrecoVendaCalculado = dblPrecoCustoCalculado + _
'                        '                          '                                   sngSaldoICMS + _
'                        '                          '                                   sngVendaIPI + _
'                        '                          '                                   (sngVendaOutrosImpostos * sngVendaPercSemImpostos) + _
'                        '                          '                                   sngVendaOutros + _
'                        '                          '                                   sngMargemLucro
          
          '---[ Altera o preço da tabela selecionada ]---'
            rsPrecos.FindFirst "Tabela = '" & Trim(cboTabelaDestino.Text) & "'"
            
            If rsPrecos.NoMatch Then
              rsPrecos.AddNew
              rsPrecos.Fields("TABELA") = Trim(cboTabelaDestino.Text)
              rsPrecos.Fields("PRODUTO") = grdProdutos.Columns("Produto").Text
            Else
              rsPrecos.Edit
            End If
            rsPrecos.Fields("Preço") = dblPrecoVendaCalculado
            rsPrecos.Update
          '---[ Altera o preço da tabela selecionada ]---'
          ![Custo Preço Venda] = dblPrecoVendaCalculado
          .Update
          
          
        End If
      End With
      
      grdProdutos.MoveNext
    '------------------------------------
  Next intCont
  
  MsgBox "Preços calculados com sucesso !!", vbInformation, "Quick Store"
End Sub

Private Sub cmdPesquisa_Click()
  Dim strSQL As String
  
  strSQL = " SELECT * FROM Produtos WHERE Código <> '0' "
  
  If Len(txtClasse.Text) > 0 Then strSQL = strSQL & " AND Classe = " & cboClasse.Text
  If Len(txtSubClasse.Text) > 0 Then strSQL = strSQL & " AND [Sub Classe] = " & cboSubClasse.Text
  If Len(txtFabricante.Text) > 0 Then strSQL = strSQL & " AND Fabricante = '" & txtFabricante.Text & "'"
  
  If optOrdenaCodigo.Value Then
    strSQL = strSQL & " ORDER BY [Código Ordenação] "
  ElseIf optOrdenaNome.Value Then
    strSQL = strSQL & " ORDER BY Nome "
  End If
  
  Set rsResultado = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rsResultado
    grdProdutos.RemoveAll
    grdProdutos.Redraw = False
    If Not (.BOF And .EOF) Then .MoveFirst
    
    Do While Not .EOF
      grdProdutos.AddItem !Código & vbTab & !Nome & vbTab & ![Custo Preço Valor]
      .MoveNext
    Loop
    grdProdutos.Redraw = True
  End With
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datTabelaDestino.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
End Sub







