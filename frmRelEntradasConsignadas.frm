VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEntradasConsignadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rel. de Entradas Consignadas"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelEntradasConsignadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   7080
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
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
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   5040
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
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' ORDER BY Nome"
      Top             =   5040
      Width           =   2415
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
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Entradas oriundas de Empréstimos de Fornecedores ( Consignações )"
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
         Left            =   480
         TabIndex        =   24
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtros Obrigatórios: Filial, Fornecedor e Datas."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   600
         Width           =   3495
      End
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
      TabIndex        =   17
      Top             =   840
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
         TabIndex        =   19
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   600
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelEntradasConsignadas.frx":058A
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
         Bindings        =   "frmRelEntradasConsignadas.frx":05A6
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
         TabIndex        =   21
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   660
         Width           =   825
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Período das Entradas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   0
      TabIndex        =   14
      Top             =   1950
      Width           =   3495
      Begin MSMask.MaskEdBox mskDataFinalEntradas 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   360
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
      Begin MSMask.MaskEdBox mskDataInicioEntradas 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   360
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
         TabIndex        =   16
         Top             =   420
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   255
      End
   End
   Begin VB.Frame fraNF 
      Caption         =   "Nota Fiscal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3600
      TabIndex        =   11
      Top             =   1950
      Width           =   3495
      Begin VB.TextBox txtNFIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         MaxLength       =   8
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtNFFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   5
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   390
         Width           =   300
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
      Height          =   810
      Left            =   0
      TabIndex        =   10
      Top             =   2760
      Width           =   3495
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "Im&primir"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin Crystal.CrystalReport crtRelEntradasConsignadas 
      Left            =   5400
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRelEntradasConsignadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()

  If ValidarDados Then Exit Sub

  Screen.MousePointer = vbHourglass
  Call StatusMsg("Aguarde, montando as informações...")
  Call CriarRecords
  Call StatusMsg("Verificando estoque dos Produtos...")
  '17/12/2004 - Daniel
  'Private comentada pois a Livraria Resultado
  'precisa saber as informações de Estoque de
  'produtos consignados e não mais de estoque
  'geral (compras + consignados)
  '
  'Call AtualizarTabela
  Call PesquisarEstoqueConsignado 'Private adicionada em 17/12/2004
  Call StatusMsg("Montando o relatório...")
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
Private Sub mskDataFinalEntradas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinalEntradas.Text = frmCalendario.gsDateCalender(mskDataFinalEntradas.Text)
  End If
End Sub

Private Sub mskDataFinalEntradas_LostFocus()
  mskDataFinalEntradas.Text = Ajusta_Data(mskDataFinalEntradas.Text)
End Sub
Private Sub mskDataInicioEntradas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicioEntradas.Text = frmCalendario.gsDateCalender(mskDataInicioEntradas.Text)
  End If
End Sub

Private Sub mskDataInicioEntradas_LostFocus()
  mskDataInicioEntradas.Text = Ajusta_Data(mskDataInicioEntradas.Text)
End Sub

Private Sub CriarRecords()
  Dim rstEntradasConsignadas As Recordset
  Dim rstEntradas            As Recordset
  Dim strSQL                 As String
  
  dbTemp.Execute "DELETE * FROM EntradasConsignadas"
  
  Set rstEntradasConsignadas = dbTemp.OpenRecordset("EntradasConsignadas", dbOpenDynaset)
  
  strSQL = "SELECT Entradas.[Nota Fiscal] AS NF, Entradas.Fornecedor AS Fornec, [Entradas - Produtos].Qtde AS QtdeOriginal, [Entradas - Produtos].Filial AS Fil, "
  strSQL = strSQL & " [Entradas - Produtos].Sequência AS Seq, [Entradas - Produtos].Linha AS Lin, [Entradas - Produtos].Código AS Cod, "
  strSQL = strSQL & " [Entradas - Produtos].Preço AS Custo, [Entradas - Produtos].[Preço Final] AS PrecoFinal "
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Entradas.Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Entradas.Operação = " & 50 'Operação de Empréstimo / Consignações
  
  If Len(txtNomeFornecedor.Text) > 0 Then strSQL = strSQL & " AND Entradas.Fornecedor = " & CLng(cboFornecedor.Text)
  
  strSQL = strSQL & " AND Entradas.Data >= #" & Format(mskDataInicioEntradas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Entradas.Data <= #" & Format(mskDataFinalEntradas.Text, "MM/DD/YYYY") & "#"
  
  If Len(txtNFIni.Text) > 0 And Len(txtNFFin.Text) > 0 Then
    If CLng(txtNFIni.Text) <= CLng(txtNFFin.Text) Then
      strSQL = strSQL & " AND NF >= " & Trim(txtNFIni.Text)
      strSQL = strSQL & " AND NF <= " & Trim(txtNFFin.Text)
    End If
  End If
  
  strSQL = strSQL & " AND [Entradas - Produtos].Filial = Entradas.Filial "
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = Entradas.Sequência "
  strSQL = strSQL & " ORDER BY Entradas.Fornecedor, [Entradas - Produtos].Filial, [Entradas - Produtos].Sequência, [Entradas - Produtos].Linha "
  
  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
      
        rstEntradasConsignadas.AddNew
          rstEntradasConsignadas.Fields("Filial").Value = .Fields("Fil").Value
          rstEntradasConsignadas.Fields("CodProduto").Value = .Fields("Cod").Value
          rstEntradasConsignadas.Fields("Sequencia").Value = .Fields("Seq").Value
          rstEntradasConsignadas.Fields("Linha").Value = .Fields("Lin").Value
          rstEntradasConsignadas.Fields("Custo").Value = Format((.Fields("Custo").Value), FORMAT_VALUE)
          rstEntradasConsignadas.Fields("Fornecedor").Value = .Fields("Fornec").Value
          rstEntradasConsignadas.Fields("QtdeOriginal").Value = .Fields("QtdeOriginal").Value
          rstEntradasConsignadas.Fields("PrecoFinal").Value = Format((.Fields("PrecoFinal").Value), FORMAT_VALUE)
        rstEntradasConsignadas.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing
  
  rstEntradasConsignadas.Close
  Set rstEntradasConsignadas = Nothing

End Sub

Private Sub AtualizarTabela()
  Dim rstEntradasConsignadas As Recordset
  Dim dblQtdeAtual           As Double
  
  Set rstEntradasConsignadas = dbTemp.OpenRecordset("EntradasConsignadas", dbOpenDynaset)
  
  With rstEntradasConsignadas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        Call ConsultarEstoque(.Fields("Filial").Value, .Fields("CodProduto").Value, dblQtdeAtual)
        
        .Edit
        .Fields("QtdeAtual").Value = dblQtdeAtual
        .Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstEntradasConsignadas = Nothing

End Sub

Private Sub ConsultarEstoque(ByVal Filial As Byte, ByVal Produto As String, ByRef QtdeAtual As Double)
  Dim rstEstoqueFinal As Recordset
  Dim strSQL          As String

  strSQL = "SELECT [Estoque Atual] FROM [Estoque Final]"
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Produto = '" & Produto & "'"
  
  Set rstEstoqueFinal = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEstoqueFinal
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      QtdeAtual = .Fields("Estoque Atual").Value
      
    End If
    .Close
  End With
  
  Set rstEstoqueFinal = Nothing
  
End Sub

Private Sub CriarRelatorio()
  Dim strReport As String
  
  'Nome do arquivo .rpt
  'strReport = gsReportPath & "rptRelEntradasConsignadas.rpt"
  strReport = gsReportPath & "rptRelEntradasConsignadas2.rpt"
  
  With crtRelEntradasConsignadas
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crtRelEntradasConsignadas
    
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsTempDBFileName
    .DataFiles(3) = gsTempDBFileName
    
    '.SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & "Empresa: " & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .Formulas(1) = "nome_filial = '" & "Filial: " & Trim(txtNomeFilial.Text) & "'"
    .Formulas(2) = "Periodo = '" & "Período de " & mskDataInicioEntradas.Text & " até " & mskDataFinalEntradas.Text & "'"
    .SortFields(0) = "+{EntradasConsignadas.Fornecedor}" 'Ordenação
    
    .WindowState = crptMaximized
    .Destination = IIf(optSaidaVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", crtRelEntradasConsignadas)
  
    .Action = 1
  End With

  Screen.MousePointer = vbDefault
  
  Call StatusMsg("")


End Sub

Private Function ValidarDados() As Boolean

  If Len(txtNomeFilial.Text) <= 0 Then
    ValidarDados = True
    MsgBox "Filial inválida, verifique", vbExclamation, "Quick Store"
    cboFilial.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataInicioEntradas.Text) Then
    ValidarDados = True
    MsgBox "Data Inicial das Entradas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataInicioEntradas.SetFocus
    Exit Function
  End If
  
  If Not IsDate(mskDataFinalEntradas.Text) Then
    ValidarDados = True
    MsgBox "Data Final das Entradas inválida, verifique.", vbExclamation, "Quick Store"
    mskDataFinalEntradas.SetFocus
    Exit Function
  End If

  If CDate(mskDataFinalEntradas.Text) < CDate(mskDataInicioEntradas.Text) Then
    ValidarDados = True
    MsgBox "Data Final das Entradas menor que a Inicial, verifique.", vbExclamation, "Quick Store"
    mskDataFinalEntradas.SetFocus
    Exit Function
  End If

  '17/12/2004 - Daniel
  'Para buscarmos o estoque correto de cada produto por
  'Fornecedor, será necessário selecionarmos o Fornecedor
  'a partir disso o campo Fornecedor será obrigatório
  If Len(txtNomeFornecedor.Text) <= 0 Then
    ValidarDados = True
    MsgBox "Fornecedor inválido, verifique.", vbExclamation, "Quick Store"
    cboFornecedor.SetFocus
    Exit Function
  End If

End Function

Private Sub PesquisarEstoqueConsignado()
  '17/12/2004 - Daniel
  'Criação da Private PesquisarEstoqueConsignado
  Dim rstEntradasConsignadas As Recordset
  Dim dblEstoqueConsignado   As Double
  
  Set rstEntradasConsignadas = dbTemp.OpenRecordset("EntradasConsignadas", dbOpenDynaset)
  
  With rstEntradasConsignadas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        Call BuscarEstoqueConsignado(.Fields("Filial").Value, .Fields("CodProduto").Value, dblEstoqueConsignado)
        
        .Edit
        If dblEstoqueConsignado <= 0 Then
          .Fields("EstoqueConsignado").Value = 0
          .Fields("Valor").Value = 0
        Else
          .Fields("EstoqueConsignado").Value = dblEstoqueConsignado
          .Fields("Valor").Value = Format((.Fields("EstoqueConsignado").Value * .Fields("Custo").Value), FORMAT_VALUE)
        End If
        .Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstEntradasConsignadas = Nothing

End Sub

Private Sub BuscarEstoqueConsignado(ByVal Filial As Byte, ByVal CodProd As String, ByRef EstoqueConsignado As Double)
  '17/12/2004 - Daniel
  'Private BuscarEstoqueConsignado
  Dim rstEntradas        As Recordset
  Dim dblQtdeEntrou      As Double
  Dim rstAcerto          As Recordset
  Dim dblQtdeSaiu        As Double
  Dim strSQL             As String
  Dim rstConsigEntra     As Recordset
  Dim dblQtdeConsigEntra As Double
  
  dblQtdeEntrou = 0
  dblQtdeSaiu = 0
  dblQtdeConsigEntra = 0
  
  strSQL = "SELECT SUM([Entradas - Produtos].Qtde) AS QtdeTotal "
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos] "
  strSQL = strSQL & " WHERE Entradas.Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND Entradas.Operação = 50 "
  
  If Len(txtNomeFornecedor.Text) > 0 Then
    strSQL = strSQL & " AND Entradas.Fornecedor = " & CLng(cboFornecedor.Text)
  End If
  
  strSQL = strSQL & " AND Entradas.Data >= #" & Format(mskDataInicioEntradas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND Entradas.Data <= #" & Format(mskDataFinalEntradas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND [Entradas - Produtos].Filial = Entradas.Filial "
  strSQL = strSQL & " AND [Entradas - Produtos].Sequência = Entradas.Sequência "
  strSQL = strSQL & " AND [Entradas - Produtos].Código = '" & CodProd & "'"
  
  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstEntradas.RecordCount > 0 Then
  
    With rstEntradas
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        dblQtdeEntrou = .Fields("QtdeTotal").Value
        
      End If
      .Close
    End With
  
    Set rstEntradas = Nothing
  Else
    rstEntradas.Close
    Set rstEntradas = Nothing
  End If
  
  '-----------------------------[Buscar as saídas]-----------------------------
  
  strSQL = ""
  strSQL = "SELECT * FROM AcertoConsignacaoEntrada "
  strSQL = strSQL & " WHERE Filial = " & CByte(cboFilial.Text)
  strSQL = strSQL & " AND DataAcerto >= #" & Format(mskDataInicioEntradas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND DataAcerto <= #" & Format(mskDataFinalEntradas.Text, "MM/DD/YYYY") & "#"
  strSQL = strSQL & " AND CodigoProduto = '" & CodProd & "'"

  Set rstAcerto = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  If rstAcerto.RecordCount > 0 Then
  
    With rstAcerto
      If Not (.BOF And .EOF) Then
        .MoveFirst
        
        Do Until .EOF
          
          If FornecedorEscolhido(.Fields("Filial").Value, .Fields("Sequencia").Value) Then
            dblQtdeSaiu = dblQtdeSaiu + .Fields("QtdeVendida").Value
          End If
          
         .MoveNext
        Loop
        
      End If
      .Close
    End With
  
    Set rstAcerto = Nothing
  
  Else
    rstAcerto.Close
    Set rstAcerto = Nothing
  End If
  
  '04/01/2004 - Daniel
  'Adicionar rotina de busca dos acertos de feitos no período
  'de 01/06/2004 a 20/10/2004 quando o sistema ainda não mapeava
  'as vendas de produtos consignados
  'If (mskDataInicioEntradas.Text) >= "01/06/2004" Or (mskDataFinalEntradas.Text) <= "20/10/2004" Then
  
  
  'End If
  '---------------------------------------------------------------------------------------------------
  
  'Logo teremos:
  EstoqueConsignado = Format((dblQtdeEntrou - dblQtdeSaiu), FORMAT_VALUE)

End Sub

Private Function FornecedorEscolhido(ByVal Filial As Byte, ByVal Sequencia As Long) As Boolean
  '17/12/2004 - Daniel
  'Criação da Private Function FornecedorEscolhido
  Dim rstEntradas As Recordset
  Dim strSQL      As String
  
  strSQL = "SELECT Fornecedor FROM Entradas "
  strSQL = strSQL & " WHERE Filial = " & Filial
  strSQL = strSQL & " AND Sequência = " & Sequencia
  
  Set rstEntradas = db.OpenRecordset(strSQL, dbOpenDynaset)
  
  With rstEntradas
    If Not (.BOF And .EOF) Then
      If .Fields("Fornecedor").Value = CLng(cboFornecedor.Text) Then FornecedorEscolhido = True
    End If
    .Close
  End With
  
  Set rstEntradas = Nothing

End Function
