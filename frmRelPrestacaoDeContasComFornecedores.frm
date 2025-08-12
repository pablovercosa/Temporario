VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelPrestacaoDeContasComFornecedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendas de Produtos Consignados"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelPrestacaoDeContasComFornecedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   7110
   Begin VB.Data datFiliais 
      Caption         =   "datFiliais"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
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
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Data datFornecedor 
      Caption         =   "datFornecedor"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
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
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' ORDER BY Nome"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   0
      TabIndex        =   15
      Top             =   1440
      Width           =   7095
      Begin VB.TextBox txtNomeFilial 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtNomeFornecedor 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelPrestacaoDeContasComFornecedores.frx":058A
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1335
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelPrestacaoDeContasComFornecedores.frx":05A6
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         DataFieldList   =   "Filial"
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
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Filial"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   3550
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   340
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Imprimir"
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2625
      Width           =   1575
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   3105
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "Período das Saídas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   3495
      Begin MSMask.MaskEdBox mskDataFinalSaidas 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataInicioSaidas 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1800
         TabIndex        =   13
         Top             =   540
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   -120
      Width           =   9615
      Begin VB.Image Image1 
         Height          =   1170
         Left            =   360
         Picture         =   "frmRelPrestacaoDeContasComFornecedores.frx":05BF
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Prestação de Contas com o Fornecedor"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelPrestacaoDeContasComFornecedores.frx":1FF1
         ForeColor       =   &H00808080&
         Height          =   855
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   4095
      End
   End
   Begin Crystal.CrystalReport crtRelPrestacao 
      Left            =   5400
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRelPrestacaoDeContasComFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()

  If Not ValidarDados Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde pesquisando no Banco de Dados...")
  
  Call CriarAcertoConsignacaoEntradaTemp
  
  Call EditarValoresEmAcertoConsignacaoEntradaTemp
  
  If Len(txtNomeFornecedor.Text) > 0 Then dbTemp.Execute "DELETE * FROM AcertoConsignacaoEntrada WHERE Fornecedor <> " & CLng(cboFornecedor.Text)
  
  '10/12/2004 - Daniel
  '
  'Solicitação de Agrupamento das informações
  'para isto foi necessário criar mais uma tabela
  'no banco temporário após a análise inicial
  'tabela temporária Acerto
  Call StatusMsg("Agrupando as informações...")
  Call TratarAgrupamento
  '-----------------------------------------------
  
  Call StatusMsg("Aguarde montando o relatório...")
  
  Call CriarRelatorio

  Call StatusMsg("")
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFiliais.DatabaseName = gsQuickDBFileName
  datFornecedor.DatabaseName = gsQuickDBFileName

End Sub

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstFiliais As Recordset
  
  txtNomeFilial.Text = ""
  If Not IsNumeric(cboFilial.Text) Then Exit Sub
  
  Set rstFiliais = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
  
  With rstFiliais
    If Not (.BOF And .EOF) Then
      txtNomeFilial.Text = .Fields("Nome") & ""
    End If
    
    If Not rstFiliais Is Nothing Then .Close
    Set rstFiliais = Nothing
  End With
End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

Private Sub cboFornecedor_LostFocus()
  txtNomeFornecedor.Text = ""
  If Not IsNumeric(cboFornecedor.Text) Then Exit Sub
  
  datFornecedor.Recordset.FindFirst "Código = " & cboFornecedor.Text
  
  If Not datFornecedor.Recordset.NoMatch Then
    txtNomeFornecedor.Text = datFornecedor.Recordset.Fields("Nome") & ""
  End If
End Sub

Private Sub mskDataFinalSaidas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinalSaidas.Text = frmCalendario.gsDateCalender(mskDataFinalSaidas.Text)
  End If
End Sub

Private Sub mskDataFinalSaidas_LostFocus()
  mskDataFinalSaidas.Text = Ajusta_Data(mskDataFinalSaidas.Text)
End Sub
Private Sub mskDataInicioSaidas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicioSaidas.Text = frmCalendario.gsDateCalender(mskDataInicioSaidas.Text)
  End If
End Sub

Private Sub mskDataInicioSaidas_LostFocus()
  mskDataInicioSaidas.Text = Ajusta_Data(mskDataInicioSaidas.Text)
End Sub

Private Function ValidarDados() As Boolean
  ValidarDados = True
  
  If Len(txtNomeFilial.Text) <= 0 Then
    ValidarDados = False
    MsgBox "Filial inválida, verifique", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataInicioSaidas.Text) Then
    ValidarDados = False
    MsgBox "Data Inicial das Saídas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataInicioSaidas.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataFinalSaidas.Text) Then
    ValidarDados = False
    MsgBox "Data Final das Saídas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataFinalSaidas.SetFocus
    Exit Function
  End If
  
  
End Function

Private Sub CriarAcertoConsignacaoEntradaTemp()
  Dim rstAcerto     As Recordset
  Dim rstAcertoTemp As Recordset
  Dim strQuery      As String
  
  dbTemp.Execute "DELETE * FROM AcertoConsignacaoEntrada "
  
  '10/12/2004 - Daniel
  'Limpar a table Acerto
  dbTemp.Execute "DELETE * FROM Acerto "
  '-------------------------------------
  
  Set rstAcertoTemp = dbTemp.OpenRecordset("AcertoConsignacaoEntrada", dbOpenDynaset)
  
  strQuery = "SELECT * FROM AcertoConsignacaoEntrada "
  strQuery = strQuery & " WHERE Filial = " & CByte(cboFilial.Text)
  strQuery = strQuery & " AND DataAcerto >= #" & Format(mskDataInicioSaidas.Text, "MM/DD/YYYY") & "#"
  strQuery = strQuery & " AND DataAcerto <= #" & Format(mskDataFinalSaidas.Text, "MM/DD/YYYY") & "#"
  
  strQuery = strQuery & " ORDER BY DataAcerto "
  
  Set rstAcerto = db.OpenRecordset(strQuery, dbOpenDynaset)
  
  With rstAcerto
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        rstAcertoTemp.AddNew
          rstAcertoTemp.Fields("Filial").Value = .Fields("Filial").Value
          rstAcertoTemp.Fields("Sequencia").Value = .Fields("Sequencia").Value
          rstAcertoTemp.Fields("DataAcerto").Value = .Fields("DataAcerto").Value
          rstAcertoTemp.Fields("LinhaProduto").Value = .Fields("LinhaProduto").Value
          rstAcertoTemp.Fields("CodigoProduto").Value = .Fields("CodigoProduto").Value & ""
          rstAcertoTemp.Fields("QtdeVendida").Value = .Fields("QtdeVendida").Value
          rstAcertoTemp.Fields("FilialVenda").Value = .Fields("FilialVenda").Value
          rstAcertoTemp.Fields("SequenciaVenda").Value = .Fields("SequenciaVenda").Value
          rstAcertoTemp.Fields("PrecoCusto").Value = .Fields("PrecoCusto").Value
        rstAcertoTemp.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstAcerto = Nothing
  
End Sub

Private Sub EditarValoresEmAcertoConsignacaoEntradaTemp()
  'Nesta Private buscaremos os valores para os campos
  'Fornecedor e PrecoVenda
  Dim rstAcertoTemp As Recordset
  Dim lngFornecedor As Long
  Dim PrecoSaida    As Double

  Set rstAcertoTemp = dbTemp.OpenRecordset("AcertoConsignacaoEntrada", dbOpenDynaset)

  With rstAcertoTemp
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        Call BuscarFornecedor(.Fields("Filial").Value, .Fields("Sequencia").Value, lngFornecedor)
        Call BuscarPrecoVenda(.Fields("FilialVenda").Value, .Fields("SequenciaVenda").Value, .Fields("CodigoProduto").Value, PrecoSaida)
      
        .Edit
        .Fields("Fornecedor").Value = lngFornecedor
        .Fields("PrecoVenda").Value = PrecoSaida
        .Update
      
        .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstAcertoTemp = Nothing
  
  '22/11/2004 - Daniel
  'Existem alguns fornecedores que emprestam materiais para a Resultado
  'e que não são 'livros' quando entra no estoque entra como operação 50 de empréstimo,
  'é criado o registro na tabela AcertoConsignacaoEntrada ao efetivar;
  'Foi solicitado para filtrar isto, então o(s) fornecedor(es) que identificamos até o
  'momento foram: 13 - Mister Pasta
  'para não exibir estas informações daremos um delete na tabela temporária.
  If Len(txtNomeFornecedor.Text) <= 0 Then
    dbTemp.Execute "DELETE * FROM AcertoConsignacaoEntrada WHERE Fornecedor = " & 13
  End If
  
End Sub

Private Sub BuscarFornecedor(ByVal Filial As Byte, ByVal Sequencia As Long, ByRef Fornecedor As Long)
  Dim rstEntradas As Recordset
  Dim strSQL      As String
  
  strSQL = "SELECT Fornecedor FROM Entradas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequência = " & Sequencia
  
  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Fornecedor = .Fields("Fornecedor").Value
      
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing
  
End Sub

Private Sub BuscarPrecoVenda(ByVal Filial As Byte, ByVal Sequencia As Long, ByVal Produto As String, ByRef PrecoSaida As Double)
  Dim rstSaidasProdutos As Recordset
  Dim strSQL            As String
  
  strSQL = "SELECT [Preço Final], Qtde FROM [Saídas - Produtos] "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequência = " & Sequencia
  strSQL = strSQL & " AND Código = '" & Produto & "'"

  Set rstSaidasProdutos = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstSaidasProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      PrecoSaida = Format((.Fields("Preço Final").Value / .Fields("Qtde").Value), FORMAT_VALUE)
    End If
    .Close
  End With
  
  Set rstSaidasProdutos = Nothing

End Sub

Private Sub CriarRelatorio()
  Dim strReport As String
  
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptRelPrestacaoDeContasComFornecedor.rpt"
  
  With crtRelPrestacao
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crtRelPrestacao
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsTempDBFileName
    .DataFiles(3) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .Formulas(1) = "Periodo = '" & "Período de " & mskDataInicioSaidas.Text & " até " & mskDataFinalSaidas.Text & "'"
    .SortFields(0) = "+{AcertoConsignacaoEntrada.Fornecedor}" 'Ordenação
    .SortFields(1) = "+{AcertoConsignacaoEntrada.DataAcerto}"
    
    .WindowState = crptMaximized
    .Destination = IIf(optSaidaVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crtRelPrestacao)
  
    .Action = 1
  End With

  Screen.MousePointer = vbDefault
  
  Call StatusMsg("")


End Sub

Private Sub TratarAgrupamento()
  Dim rstAcertoConsigEnt As Recordset
  Dim rstAcerto          As Recordset
  Dim strSQL             As String
  
  strSQL = "SELECT Filial, Fornecedor, DataAcerto, CodigoProduto, SUM(QtdeVendida) AS Qtde, SUM(PrecoCusto) AS CUSTO, SUM(PrecoVenda) AS VENDA "
  strSQL = strSQL & " FROM AcertoConsignacaoEntrada "
  strSQL = strSQL & " GROUP BY Filial, Fornecedor, DataAcerto, CodigoProduto "
  strSQL = strSQL & " ORDER BY DataAcerto"
  
  Set rstAcertoConsigEnt = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
  Set rstAcerto = dbTemp.OpenRecordset("Acerto", dbOpenDynaset)
  
  With rstAcertoConsigEnt
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        rstAcerto.AddNew
          rstAcerto.Fields("Filial").Value = .Fields("Filial").Value
          rstAcerto.Fields("DataAcerto").Value = .Fields("DataAcerto").Value
          rstAcerto.Fields("CodigoProduto").Value = .Fields("CodigoProduto").Value & ""
          rstAcerto.Fields("QtdeVendida").Value = .Fields("Qtde").Value
          rstAcerto.Fields("PrecoCusto").Value = .Fields("CUSTO").Value
          rstAcerto.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
          rstAcerto.Fields("PrecoVenda").Value = .Fields("VENDA").Value
        rstAcerto.Update
      
        .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstAcertoConsigEnt = Nothing

  
  'Remontamos a tabela a ser exibida no relatório
  'com as informações já agrupadas
  dbTemp.Execute "DELETE * FROM AcertoConsignacaoEntrada "
  
  Set rstAcertoConsigEnt = dbTemp.OpenRecordset("AcertoConsignacaoEntrada", dbOpenDynaset)
  
  With rstAcerto
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        
        rstAcertoConsigEnt.AddNew
          rstAcertoConsigEnt.Fields("Filial").Value = .Fields("Filial").Value
          rstAcertoConsigEnt.Fields("DataAcerto").Value = .Fields("DataAcerto").Value
          rstAcertoConsigEnt.Fields("CodigoProduto").Value = .Fields("CodigoProduto").Value & ""
          rstAcertoConsigEnt.Fields("QtdeVendida").Value = .Fields("QtdeVendida").Value
          rstAcertoConsigEnt.Fields("PrecoCusto").Value = .Fields("PrecoCusto").Value
          rstAcertoConsigEnt.Fields("Fornecedor").Value = .Fields("Fornecedor").Value
          rstAcertoConsigEnt.Fields("PrecoVenda").Value = .Fields("PrecoVenda").Value
        rstAcertoConsigEnt.Update
        
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstAcerto = Nothing
  
  rstAcertoConsigEnt.Close
  Set rstAcertoConsigEnt = Nothing
  
End Sub
