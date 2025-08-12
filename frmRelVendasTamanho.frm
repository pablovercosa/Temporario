VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendasTamanho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de vendas por tamanho"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelVendasTamanho.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   8070
   Begin Crystal.CrystalReport crpRelVendasPorTamanho 
      Left            =   6000
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraSaida 
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
      Height          =   970
      Left            =   3000
      TabIndex        =   23
      Top             =   5280
      Width           =   4935
      Begin VB.OptionButton optVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Data datFornecedor 
      Caption         =   "datFornecedor"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' ORDER BY Código"
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtFornecedor 
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
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox txtFilial 
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
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   360
      Width           =   5895
   End
   Begin VB.Data datParametros 
      Caption         =   "datParametros"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial] ORDER BY Filial"
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame fraOrdem 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   970
      Left            =   3000
      TabIndex        =   20
      Top             =   4200
      Width           =   4935
      Begin VB.OptionButton optRankingValor 
         Caption         =   "Ranking por Valor"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optRankingQtde 
         Caption         =   "Ranking por Quantidade"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optOrdemCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optOrdemNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "Imprimir"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   970
      Left            =   3000
      TabIndex        =   8
      Top             =   3120
      Width           =   4935
      Begin MSMask.MaskEdBox mskPeriodoFin 
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   375
         Width           =   1260
         _ExtentX        =   2223
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskPeriodoIni 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Pressione F2 para Calendário"
         Top             =   375
         Width           =   1260
         _ExtentX        =   2223
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Data Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   405
         Width           =   885
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Data Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   405
         Width           =   975
      End
   End
   Begin VB.ListBox lstTamanhos 
      Height          =   3210
      ItemData        =   "frmRelVendasTamanho.frx":058A
      Left            =   120
      List            =   "frmRelVendasTamanho.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ListBox lstClasses 
      Height          =   1185
      ItemData        =   "frmRelVendasTamanho.frx":058E
      Left            =   120
      List            =   "frmRelVendasTamanho.frx":0590
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1680
      Width           =   2775
   End
   Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
      Bindings        =   "frmRelVendasTamanho.frx":0592
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1815
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
      Columns(0).Width=   3200
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Código"
   End
   Begin SSDataWidgets_B.SSDBCombo cboFilial 
      Bindings        =   "frmRelVendasTamanho.frx":05AE
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
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
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Filial"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tamanhos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Classes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmRelVendasTamanho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFilial_CloseUp()
  cboFilial.Text = cboFilial.Columns(0).Text
  cboFilial_LostFocus
End Sub

Private Sub cboFilial_LostFocus()
  Dim rstParametros As Recordset

  txtFilial.Text = ""
  
  If Not IsNumeric(cboFilial.Text) Then Exit Sub

  Set rstParametros = db.OpenRecordset("SELECT Filial, Nome FROM [Parâmetros Filial] WHERE Filial = " & CByte(cboFilial.Text), dbOpenDynaset)

  With rstParametros
    If Not (.BOF And .EOF) Then
      txtFilial.Text = .Fields("Nome") & ""
    End If
  End With

  rstParametros.Close
  Set rstParametros = Nothing
End Sub

Private Sub cboFornecedor_CloseUp()
  cboFornecedor.Text = cboFornecedor.Columns(0).Text
  cboFornecedor_LostFocus
End Sub

Private Sub cboFornecedor_LostFocus()
  Dim rstFornecedor As Recordset

  txtFornecedor.Text = ""
  
  If Not IsNumeric(cboFornecedor.Text) Then Exit Sub
  
  Set rstFornecedor = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & CInt(cboFornecedor.Text), dbOpenDynaset)

  With rstFornecedor
    If Not (.BOF And .EOF) Then
      txtFornecedor.Text = .Fields("Nome") & ""
    End If
  End With

  rstFornecedor.Close
  Set rstFornecedor = Nothing
  
End Sub


Private Sub cmdImprimir_Click()
  Dim rstRelVendasTamanho       As Recordset
  Dim strSQL                    As String

  Dim dblQtdeTotalDev           As Double: dblQtdeTotalDev = 0
  Dim dblValorTotalDev          As Double: dblValorTotalDev = 0
  Dim dblTotalDescontoSubTotal  As Double: dblTotalDescontoSubTotal = 0
  
  Dim intX                      As Integer
  Dim intAuxi                   As Integer

'Validação para os campos tipo data

  If Not IsDate(mskPeriodoIni.Text) Then
    MsgBox "Período Inicial inválido, verifique.", vbExclamation, "Quick Store"
    mskPeriodoIni.SetFocus
    Exit Sub
  End If

  If Not IsDate(mskPeriodoFin.Text) Then
    MsgBox "Período Final inválido, verifique.", vbExclamation, "Quick Store"
    mskPeriodoFin.SetFocus
    Exit Sub
  End If

  If CDate(mskPeriodoIni.Text) > CDate(mskPeriodoFin.Text) Then
    MsgBox "Período Final menor que o Período Inicial, verifique.", vbExclamation, "Quick Store"
    mskPeriodoFin.SetFocus
    Exit Sub
  End If
'------------------------------------------------------------------------------------------------

  dbTemp.Execute "DELETE * FROM tblRelVendasTamanho"

  Call StatusMsg("Gerando as informações do tipo grade, aguarde . . . ")
  GeraRelatorio

  Call StatusMsg("")
  '---[ Chamada das funções para geração da tabela temporária ]---'

  Set rstRelVendasTamanho = dbTemp.OpenRecordset("SELECT * FROM tblRelVendasTamanho", dbOpenSnapshot)

  With rstRelVendasTamanho
    If Not (.BOF And .EOF) Then
      '---[ Gera o total de Descontos do sub-total ]---'
        Call StatusMsg("Analisando descontos no sub-total e devoluções, aguarde . . . ")
        ReturnDescontoSubTotal dblTotalDescontoSubTotal
        ReturnDevolucaoGrade dblValorTotalDev, dblQtdeTotalDev
      '---[ Gera o total de Descontos do sub-total ]---'

      With crpRelVendasPorTamanho
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized

        If optVideo.Value Then .Destination = crptToWindow
        If optImpressora.Value Then .Destination = crptToPrinter

        .SortFields(0) = "-{tblRelVendasTamanho.Tipo}"

        If optOrdemCodigo.Value Then .SortFields(1) = "+{Produtos.Código Ordenação}"
        If optOrdemNome.Value Then .SortFields(1) = "+{Produtos.Nome}"
        If optRankingQtde.Value Then .SortFields(1) = "+{tblRelVendasTamanho.Qtde Vendida}"
        If optRankingValor.Value Then .SortFields(1) = "+{tblRelVendasTamanho.Valor Vendido}"

        .ReportFileName = gsReportPath & "rptVendasTamanho.rpt"
        
        ' Modelo 1 ou 2
        'SetPrinterModeloPwd2 crpRelVendasPorTamanho

        'Foram incluídas 06 Tabelas no pacote do Crystal
        'tblRelVendasTamanho e Report do Temp.mdb e
        'Parâmetros Filial, Produtos, Tamanhos e Cores
        'do QuickStore.mdb
        .DataFiles(0) = gsQuickDBFileName
        .DataFiles(1) = gsQuickDBFileName
        .DataFiles(2) = gsQuickDBFileName
        .DataFiles(3) = gsQuickDBFileName
        .DataFiles(4) = gsTempDBFileName
        .DataFiles(5) = gsQuickDBFileName
        .DataFiles(6) = gsQuickDBFileName
        
        .Formulas(0) = "DescSubTotal = " & Replace(Format(CStr(dblTotalDescontoSubTotal), "###0.00"), gsCurrencyDecimal, ".")
        .Formulas(1) = "DevolucoesQtde = " & Replace(Format(CStr(dblQtdeTotalDev), "###0.00"), gsCurrencyDecimal, ".")
        .Formulas(2) = "DevolucoesValor = " & Replace(Format(CStr(dblValorTotalDev), "###0.00"), gsCurrencyDecimal, ".")

        '---[ Preenchimento dos campos de cabeçalho de filtro ]---'
          .Formulas(3) = "Periodo = '" & "De " & mskPeriodoIni.Text & " até " & mskPeriodoFin.Text & "'"

          If Len(Trim(txtFilial.Text)) > 0 Then .Formulas(4) = "Filtro_Filial = '" & txtFilial.Text & "'"
          If Len(Trim(txtFornecedor.Text)) > 0 Then .Formulas(5) = "Filtro_Fornecedor = '" & txtFornecedor.Text & "'"
          
          intAuxi = 6 'Seis Fórmulas até o momento: 0,1,2,3,4,5
          
          With lstClasses
            For intX = 0 To .ListCount - 1
              If .Selected(intX) Then
                intAuxi = intAuxi + 1
                
                crpRelVendasPorTamanho.Formulas(intAuxi) = "Filtro_Classe = '" & .List(intX) & "'"
                
              End If
            Next intX
          End With
          
          With lstTamanhos
            For intX = 0 To .ListCount - 1
              If .Selected(intX) Then
                intAuxi = intAuxi + 1
                
                crpRelVendasPorTamanho.Formulas(intAuxi) = "Filtro_Tamanho = '" & .List(intX) & "'"
                
              End If
            Next intX
          End With
          
          intAuxi = 0

        '---[ Preenchimento dos campos de cabeçalho de filtro ]---'

        'Seta a impressora para relatório
        Call SetPrinterName("REL", crpRelVendasPorTamanho)

        .Action = 1
'        pgbProgress.Value = 0
      End With
    Else
      MsgBox "Não existem informações a serem exibidas !", vbInformation, App.Title
    End If
  End With

  Call StatusMsg("")
End Sub
  
Private Sub GeraRelatorio()
  Dim strSQL                    As String
  Dim blnInTransaction          As Boolean
  
  Dim rstVendas                 As Recordset
  Dim rstRelVendasTamanho       As Recordset
  Dim rstProdutos               As Recordset
  
  Dim intTamanho                As Integer
  Dim intCor                    As Integer
  
  Dim blnProdutoOK              As Boolean
  
  Dim intX                      As Integer
  Dim intCodigoTamanhos         As Integer   'Código da tabela Tamanhos
  Dim strCodigoTamanhos         As String    'Código da tabela Tamanhos em string
  Dim rstTamanhos               As Recordset
  
  Dim intCodigoClasse           As Integer   'Código da Classe
  Dim rstClasses                As Recordset
  Dim intQtdeItensSelecionados  As Integer
  Dim intAuxi                   As Integer
  
  Dim blnAdicionaProduto        As Boolean
  
  strSQL = " SELECT Saídas.Filial, Saídas.Data, [Saídas - Produtos].Código, [Saídas - Produtos].[Código sem Grade], [Saídas - Produtos].Qtde, [Saídas - Produtos].[Preço Final], [Operações Saída].Tipo "    'Sum([Saídas - Produtos].Qtde) AS SomaQtde, Sum([Saídas - Produtos].[Preço Final]) AS SomaPrecoFinal
  
  strSQL = strSQL & " FROM Saídas, [Saídas - Produtos], Produtos, [Códigos da Grade], [Operações Saída] "
  strSQL = strSQL & " WHERE (Saídas.Sequência = [Saídas - Produtos].Sequência) "
  strSQL = strSQL & " AND (Saídas.Filial = [Saídas - Produtos].Filial) "
  strSQL = strSQL & " AND ([Saídas - Produtos].[Código sem Grade] = Produtos.Código) "
  '*** Esta Linha abaixo realiza o filtro necessário para [Códigos da Grade]
  strSQL = strSQL & " AND ([Códigos da Grade].[Código Original] = [Saídas - Produtos].[Código sem Grade]) "
  strSQL = strSQL & " AND (Saídas.Operação = [Operações Saída].Código) "
  strSQL = strSQL & " AND ( Saídas.Efetivada )  "
  strSQL = strSQL & " AND ( NOT Saídas.[Nota Cancelada]) "
  strSQL = strSQL & " AND ( [Operações Saída].Tipo = 'V' )  "
  strSQL = strSQL & " AND Produtos.Tipo = 'G' "
  strSQL = strSQL & " AND (Saídas.Data >= #" & Format(mskPeriodoIni.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Saídas.Data <= #" & Format(mskPeriodoFin.Text, "mm/dd/yyyy") & "#) "
  
  
  If Len(Trim(txtFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Filial = " & cboFilial.Text & ") "
  End If
  
                      
  '------ TRATAMENTO PARA TAMANHOS
  intQtdeItensSelecionados = lstTamanhos.SelCount
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & " AND ("
  End If
  
  With lstTamanhos
    
    For intX = 0 To .ListCount - 1 'ListCount para pegar todos da lista
    
      If .Selected(intX) Then
        
          Set rstTamanhos = db.OpenRecordset("SELECT Código, Nome FROM Tamanhos WHERE Nome = '" & Trim(.List(intX)) & "'")
          
          With rstTamanhos
            If Not (.BOF And .EOF) Then
              intCodigoTamanhos = .Fields("Código").Value
            End If
          End With
          
          strCodigoTamanhos = "0000000000" + CStr(intCodigoTamanhos)
          strCodigoTamanhos = Right(strCodigoTamanhos, 3)
        
                        
          If intQtdeItensSelecionados = 1 Then
            strSQL = strSQL & " ((Left(Right([Saídas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
          Else
            intAuxi = intAuxi + 1
            
            If intAuxi = 1 Then
              strSQL = strSQL & " ((Left(Right([Saídas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
            Else
              strSQL = strSQL & " OR ((Left(Right([Saídas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
            End If
          End If
          
          rstTamanhos.Close
          Set rstTamanhos = Nothing
  
      End If 'If .Selected
  
    Next intX 'For intX
  
  
    intAuxi = 0
  
  End With 'lstTamanhos
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & ")"
  End If
  '------ FIM DO TRATAMENTO PARA TAMANHOS

  '------ TRATAMENTO PARA CLASSES
  intQtdeItensSelecionados = lstClasses.SelCount
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & " AND ("
  End If
  
  With lstClasses
  
    For intX = 0 To .ListCount - 1 'ListCount para pegar todos da lista
      
      If .Selected(intX) Then
      
        Set rstClasses = db.OpenRecordset("SELECT Código, Nome FROM Classes WHERE Nome = '" & .List(intX) & "'", dbOpenDynaset)
        
        With rstClasses
          If Not (.BOF And .EOF) Then
            intCodigoClasse = .Fields("Código").Value
            
              If intQtdeItensSelecionados = 1 Then
                strSQL = strSQL & " (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
              Else
                intAuxi = intAuxi + 1
                
                If intAuxi = 1 Then
                  strSQL = strSQL & " (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
                Else
                  strSQL = strSQL & " OR (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
                End If
              End If
            
          End If 'If Not
        End With
        
        rstClasses.Close
        Set rstClasses = Nothing
   
      End If
  
    Next intX
    
    
    intAuxi = 0
  
  End With 'lstClasses
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & ")"
  End If
  '------ FIM DO TRATAMENTO PARA CLASSES
  
  strSQL = strSQL & " GROUP BY Saídas.Filial, [Saídas - Produtos].Código, [Saídas - Produtos].[Código sem Grade], Saídas.Efetivada, Saídas.[Nota Cancelada], [Operações Saída].Tipo, Saídas.Data, Saídas.Data, Saídas.Filial, Saídas.Cliente, [Saídas - Produtos].[Código sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Saídas.Digitador, [Saídas - Produtos].Qtde, [Saídas - Produtos].[Preço Final] "
  
  intQtdeItensSelecionados = 0
  
  '-----------------------------------------------------------------------------------------------
  
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If (.BOF And .EOF) Then
      Exit Sub
    End If
  End With
  
  rstVendas.MoveLast
  rstVendas.MoveFirst
  
'  pgbProgress.Min = 0
'  pgbProgress.Max = rstVendas.RecordCount + 1
  
  Set rstRelVendasTamanho = dbTemp.OpenRecordset("SELECT * FROM tblRelVendasTamanho", dbOpenDynaset)

  ws.BeginTrans
  blnInTransaction = True
  
  With rstVendas
    .MoveFirst
    
    Do While Not .EOF
      blnAdicionaProduto = False
      
      blnAdicionaProduto = Len(Trim(txtFornecedor.Text)) <= 0
      
      If Not blnAdicionaProduto Then
        blnAdicionaProduto = blnVerificaFornecedor(.Fields("Código Sem Grade"))
      End If
      
      If blnAdicionaProduto Then
        rstRelVendasTamanho.AddNew
        
        rstRelVendasTamanho.Fields("Filial") = .Fields("Filial")
        rstRelVendasTamanho.Fields("Data") = .Fields("Data")
        rstRelVendasTamanho.Fields("Produto") = .Fields("Código sem Grade")
        rstRelVendasTamanho.Fields("Tamanho") = Left(Right(.Fields("Código"), 6), 3)
        rstRelVendasTamanho.Fields("Cor") = Right(.Fields("Código"), 3)
        rstRelVendasTamanho.Fields("Edicao") = 0
        rstRelVendasTamanho.Fields("Qtde Vendida") = .Fields("Qtde")
        rstRelVendasTamanho.Fields("Valor Vendido") = .Fields("Preço Final")
        rstRelVendasTamanho.Fields("Tipo") = "G"
        
        rstRelVendasTamanho.Update
'      pgbProgress.Value = rstVendas.AbsolutePosition
      End If
      .MoveNext
    Loop
  End With
  
  ws.CommitTrans
  blnInTransaction = False
  
'  If Not rstRelVendas Is Nothing Then rstRelVendas.Close
'  Set rstRelVendas = Nothing
  
  If Not rstVendas Is Nothing Then rstVendas.Close
  Set rstVendas = Nothing


End Sub

Private Sub Form_Load()
  datParametros.DatabaseName = gsQuickDBFileName
  datFornecedor.DatabaseName = gsQuickDBFileName
  
  PreencheLstClasses
  PreencheLstTamanhos

  Call CenterForm(Me)
End Sub

Private Sub PreencheLstClasses()
  Dim rstClasses As Recordset

  Set rstClasses = db.OpenRecordset("SELECT Nome FROM Classes ORDER BY Código", dbOpenDynaset)

  With rstClasses
    If Not (.BOF And .EOF) Then
      .MoveFirst
      Do Until .EOF
        lstClasses.AddItem (.Fields("Nome").Value)
        .MoveNext
      Loop
    End If
  End With
  
  rstClasses.Close
  Set rstClasses = Nothing
End Sub

Private Sub PreencheLstTamanhos()
  Dim rstTamanhos As Recordset
  
  Set rstTamanhos = db.OpenRecordset("SELECT Nome FROM Tamanhos ORDER BY Código", dbOpenDynaset)

  With rstTamanhos
    If Not (.BOF And .EOF) Then
      .MoveFirst
      Do Until .EOF
        lstTamanhos.AddItem (.Fields("Nome").Value)
        .MoveNext
      Loop
    End If
  End With

  rstTamanhos.Close
  Set rstTamanhos = Nothing
End Sub

Private Function ReturnDescontoSubTotal(ByRef dblValorDesconto As Double) As Double
  Dim strSQL              As String
  Dim blnInTransaction    As Boolean
  
  Dim rstVendas           As Recordset
  Dim rstProdutos         As Recordset
  Dim rstDescontoSubTotal As Recordset
  
  Dim dblDescontoSubTotal As Double
  Dim dblDescontoSomar    As Double
  Dim blnProdutoOK        As Boolean
  
  Dim intX                      As Integer
  Dim intCodigoTamanhos         As Integer   'Código da tabela Tamanhos
  Dim strCodigoTamanhos         As String    'Código da tabela Tamanhos em string
  Dim rstTamanhos               As Recordset
  
  Dim intCodigoClasse           As Integer   'Código da Classe
  Dim rstClasses                As Recordset
  Dim intQtdeItensSelecionados  As Integer
  Dim intAuxi                   As Integer

  
  strSQL = " SELECT (Saídas.DescontoSubTotal) AS DescontoSubTotal, [Saídas - Produtos].[Código sem Grade], Saídas.Filial, Saídas.Sequência " 'SELECT SUM(Saídas.DescontoSubTotal) AS DescontoSubTotal .. Para tratamento individual
  strSQL = strSQL & " FROM Saídas, [Saídas - Produtos], Produtos, [Códigos da Grade], [Operações Saída] "
  strSQL = strSQL & " WHERE (Saídas.Sequência = [Saídas - Produtos].Sequência) "
  strSQL = strSQL & " AND (Saídas.Filial = [Saídas - Produtos].Filial) "
  strSQL = strSQL & " AND ([Saídas - Produtos].[Código sem Grade] = Produtos.Código) "
  '****
  strSQL = strSQL & " AND ([Códigos da Grade].[Código Original] = [Saídas - Produtos].[Código sem Grade]) "
  strSQL = strSQL & " AND (Saídas.Operação = [Operações Saída].Código) "
  strSQL = strSQL & " AND ( Saídas.Efetivada ) "
  strSQL = strSQL & " AND ( NOT Saídas.[Nota Cancelada]) "
  strSQL = strSQL & " AND ( [Operações Saída].Tipo = 'V' ) "
  strSQL = strSQL & " AND Saídas.DescontoSubTotal > 0 "
  strSQL = strSQL & " AND (Saídas.Data >= #" & Format(mskPeriodoIni.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Saídas.Data <= #" & Format(mskPeriodoFin.Text, "mm/dd/yyyy") & "#) "
  
  
  If Len(Trim(txtFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Filial = " & cboFilial.Text & ") "
  End If
  
  
  '------ TRATAMENTO PARA TAMANHOS
  intQtdeItensSelecionados = lstTamanhos.SelCount
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & " AND ("
  End If
  
  With lstTamanhos
    
    For intX = 0 To .ListCount - 1 'ListCount para pegar todos da lista
    
      If .Selected(intX) Then
        
          Set rstTamanhos = db.OpenRecordset("SELECT Código, Nome FROM Tamanhos WHERE Nome = '" & Trim(.List(intX)) & "'")
          
          With rstTamanhos
            If Not (.BOF And .EOF) Then
              intCodigoTamanhos = .Fields("Código").Value
            End If
          End With
          
          strCodigoTamanhos = "0000000000" + CStr(intCodigoTamanhos)
          strCodigoTamanhos = Right(strCodigoTamanhos, 3)
        
                        
          If intQtdeItensSelecionados = 1 Then
            strSQL = strSQL & " ((Left(Right([Saídas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
          Else
            intAuxi = intAuxi + 1
          
            If intAuxi = 1 Then
              strSQL = strSQL & " ((Left(Right([Saídas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
            Else
              strSQL = strSQL & " OR ((Left(Right([Saídas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
            End If
          End If
          
          rstTamanhos.Close
          Set rstTamanhos = Nothing
  
      End If 'If .Selected
  
    Next intX 'For intX
  
  
    intAuxi = 0
  
  End With 'lstTamanhos
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & ")"
  End If
  '------ FIM DO TRATAMENTO PARA TAMANHOS

  
  '------ TRATAMENTO PARA CLASSES
  intQtdeItensSelecionados = lstClasses.SelCount
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & " AND ("
  End If
  
  With lstClasses
  
    For intX = 0 To .ListCount - 1 'ListCount para pegar todos da lista
      
      If .Selected(intX) Then
      
        Set rstClasses = db.OpenRecordset("SELECT Código, Nome FROM Classes WHERE Nome = '" & .List(intX) & "'", dbOpenDynaset)
        
        With rstClasses
          If Not (.BOF And .EOF) Then
            intCodigoClasse = .Fields("Código").Value
            
              If intQtdeItensSelecionados = 1 Then
                strSQL = strSQL & " (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
              Else
                intAuxi = intAuxi + 1
                
                If intAuxi = 1 Then
                  strSQL = strSQL & " (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
                Else
                  strSQL = strSQL & " OR (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
                End If
              End If
            
          End If 'If Not
        End With
        
        rstClasses.Close
        Set rstClasses = Nothing
   
      End If
  
    Next intX
    
    
    intAuxi = 0
  
  End With 'lstClasses
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & ")"
  End If
  '------ FIM DO TRATAMENTO PARA CLASSES
  
  strSQL = strSQL & " GROUP BY Saídas.Filial, Saídas.Data, Saídas.Cliente, [Saídas - Produtos].[Código sem Grade], Saídas.Digitador, Produtos.Classe, Produtos.[Sub Classe], Saídas.Efetivada, Saídas.[Nota Cancelada], [Operações Saída].Tipo = 'V', Saídas.Sequência, Saídas.DescontoSubTotal "
  
  
  intQtdeItensSelecionados = 0
  
  '--------------------------------------------------------------------------
  
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If (.BOF And .EOF) Then
      Exit Function
    End If
    
    .MoveLast
    .MoveFirst
    
'    pgbProgress.Min = 0
'    pgbProgress.Max = .RecordCount + 1
  End With

  With rstVendas
    .MoveFirst
    
    dbTemp.Execute "DELETE * FROM tblRelVendasDescontoSubTotal"
    
    Do While Not .EOF
      strSQL = " SELECT * FROM tblRelVendasDescontoSubTotal WHERE filID = " & .Fields("Filial")
      strSQL = strSQL & " AND movSequencia = " & .Fields("Sequência")
      
      If CDbl(.Fields("DescontoSubTotal")) > 0 Then
        Set rstDescontoSubTotal = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        
        If (rstDescontoSubTotal.BOF And rstDescontoSubTotal.EOF) Then
          dblDescontoSomar = .Fields("DescontoSubTotal")
          
          rstDescontoSubTotal.AddNew
          rstDescontoSubTotal.Fields("filID") = .Fields("Filial")
          rstDescontoSubTotal.Fields("movSequencia") = .Fields("Sequência")
          rstDescontoSubTotal.Fields("movValorDesconto") = dblDescontoSomar
          rstDescontoSubTotal.Update
        Else
          dblDescontoSomar = 0
        End If
      Else
        dblDescontoSomar = 0
      End If
      rstDescontoSubTotal.Close
      Set rstDescontoSubTotal = Nothing
      
      Set rstProdutos = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE Código = '" & .Fields("Código Sem Grade") & "' ", dbOpenSnapshot)
      
      blnProdutoOK = Not (rstProdutos.BOF And rstProdutos.EOF)
      
      rstProdutos.Close
      Set rstProdutos = Nothing
      
      If blnProdutoOK Then
        If Len(Trim(txtFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaFornecedor(.Fields("Código Sem Grade"))
        End If
      End If
      
      If blnProdutoOK Then
        dblDescontoSubTotal = dblDescontoSubTotal + dblDescontoSomar
      End If
      
'      pgbProgress.Value = .AbsolutePosition
      .MoveNext
    Loop
  End With
  
  dblValorDesconto = dblDescontoSubTotal
  
  If Not rstVendas Is Nothing Then rstVendas.Close
  Set rstVendas = Nothing
End Function

Private Function ReturnDevolucaoGrade(ByRef dblValorDevolucao As Double, _
                                      ByRef dblQtdeDevolucao As Double) As Boolean
  
  Dim strSQL                    As String
  Dim rstDev                    As Recordset
  Dim blnProdutoOK              As Boolean
  
  Dim intX                      As Integer
  Dim intCodigoTamanhos         As Integer   'Código da tabela Tamanhos
  Dim strCodigoTamanhos         As String    'Código da tabela Tamanhos em string
  Dim rstTamanhos               As Recordset
  
  Dim intCodigoClasse           As Integer   'Código da Classe
  Dim rstClasses                As Recordset
  Dim intQtdeItensSelecionados  As Integer
  Dim intAuxi                   As Integer
  
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Códigos da Grade].[Código Original], [Entradas - Produtos].Qtde, [Entradas - Produtos].[Preço Final] "   'Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal
  strSQL = strSQL & " FROM Entradas, [Entradas - Produtos], [Operações Entrada], [Códigos da Grade], Produtos "
  strSQL = strSQL & " WHERE (Entradas.Filial = [Entradas - Produtos].Filial) "
  strSQL = strSQL & " AND (Entradas.Sequência = [Entradas - Produtos].Sequência) "
  strSQL = strSQL & " AND (Entradas.Operação = [Operações Entrada].Código) "
  strSQL = strSQL & " AND ([Entradas - Produtos].Código = [Códigos da Grade].Código) "
  strSQL = strSQL & " AND ([Códigos da Grade].[Código Original] = Produtos.Código) "
  strSQL = strSQL & " AND ([Operações Entrada].Tipo ='D') "

  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(mskPeriodoIni.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(mskPeriodoFin.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & cboFilial.Text & ") "
  End If
    
  
  '------ TRATAMENTO PARA TAMANHOS
  intQtdeItensSelecionados = lstTamanhos.SelCount
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & " AND ("
  End If
  
  With lstTamanhos
    
    For intX = 0 To .ListCount - 1 'ListCount para pegar todos da lista
    
      If .Selected(intX) Then
        
          Set rstTamanhos = db.OpenRecordset("SELECT Código, Nome FROM Tamanhos WHERE Nome = '" & Trim(.List(intX)) & "'")
          
          With rstTamanhos
            If Not (.BOF And .EOF) Then
              intCodigoTamanhos = .Fields("Código").Value
            End If
          End With
          
          strCodigoTamanhos = "0000000000" + CStr(intCodigoTamanhos)
          strCodigoTamanhos = Right(strCodigoTamanhos, 3)
        
                        
          If intQtdeItensSelecionados = 1 Then
            strSQL = strSQL & " ((Left(Right([Entradas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
          Else
            intAuxi = intAuxi + 1
          
            If intAuxi = 1 Then
              strSQL = strSQL & " ((Left(Right([Entradas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
            Else
              strSQL = strSQL & " OR ((Left(Right([Entradas - Produtos].Código,6),3)) ='" & strCodigoTamanhos & "'" & ") "
            End If
          End If
          
          rstTamanhos.Close
          Set rstTamanhos = Nothing
  
      End If 'If .Selected
  
    Next intX 'For intX
  
  
    intAuxi = 0
  
  End With 'lstTamanhos
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & ")"
  End If
  '------ FIM DO TRATAMENTO PARA TAMANHOS

  
  '------ TRATAMENTO PARA CLASSES
  intQtdeItensSelecionados = lstClasses.SelCount
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & " AND ("
  End If
  
  With lstClasses
  
    For intX = 0 To .ListCount - 1 'ListCount para pegar todos da lista
      
      If .Selected(intX) Then
      
        Set rstClasses = db.OpenRecordset("SELECT Código, Nome FROM Classes WHERE Nome = '" & .List(intX) & "'", dbOpenDynaset)
        
        With rstClasses
          If Not (.BOF And .EOF) Then
            intCodigoClasse = .Fields("Código").Value
            
              If intQtdeItensSelecionados = 1 Then
                strSQL = strSQL & " (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
              Else
                intAuxi = intAuxi + 1
                
                If intAuxi = 1 Then
                  strSQL = strSQL & " (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
                Else
                  strSQL = strSQL & " OR (Produtos.Classe = " & CInt(intCodigoClasse) & ") "
                End If
              End If
            
          End If 'If Not
        End With
        
        rstClasses.Close
        Set rstClasses = Nothing
   
      End If
  
    Next intX
    
    
    intAuxi = 0
  
  End With 'lstClasses
  
  If intQtdeItensSelecionados >= 1 Then
    strSQL = strSQL & ")"
  End If
  '------ FIM DO TRATAMENTO PARA CLASSES
  
  
  strSQL = strSQL & " GROUP BY Entradas.Filial, Entradas.Data, [Códigos da Grade].[Código Original], Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe], [Entradas - Produtos].Qtde, [Entradas - Produtos].[Preço Final] "
  
  
  intQtdeItensSelecionados = 0
  
  '--------------------------------------------------
  
  Set rstDev = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstDev
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        blnProdutoOK = True
        If Len(Trim(txtFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaFornecedor(.Fields("Código Original"))
        End If
        
        If blnProdutoOK Then
          dblValorDevolucao = dblValorDevolucao + CDbl(.Fields("Preço Final"))
          dblQtdeDevolucao = dblQtdeDevolucao + CDbl(.Fields("Qtde"))
        End If
        
        .MoveNext
      Loop
    End If
  End With
End Function

Private Function blnVerificaFornecedor(strCodigoProduto As String) As Boolean
  Dim rstFornProd As Recordset
  
  Set rstFornProd = db.OpenRecordset(" SELECT * FROM Forn_Prod " & _
                                     " WHERE Produto = '" & strCodigoProduto & "' " & _
                                     " AND Fornecedor = " & CLng(cboFornecedor.Text), dbOpenSnapshot)
  With rstFornProd
    blnVerificaFornecedor = Not (.BOF And .EOF)
    
    rstFornProd.Close
    Set rstFornProd = Nothing
  End With
End Function

Private Sub mskPeriodoFin_KeyDown(KeyCode As Integer, Shift As Integer)
'A tecla está pressionada para baixo
  If KeyCode = vbKeyF2 Then
    mskPeriodoFin.Text = frmCalendario.gsDateCalender(mskPeriodoFin.Text)
  End If
End Sub

Private Sub mskPeriodoFin_LostFocus()
  mskPeriodoFin.Text = Ajusta_Data(mskPeriodoFin.Text)
End Sub

Private Sub mskPeriodoIni_KeyDown(KeyCode As Integer, Shift As Integer)
'A tecla está pressionada para baixo
  If KeyCode = vbKeyF2 Then
    mskPeriodoIni.Text = frmCalendario.gsDateCalender(mskPeriodoIni.Text)
  End If
End Sub

Private Sub mskPeriodoIni_LostFocus()
  mskPeriodoIni.Text = Ajusta_Data(mskPeriodoIni.Text)
End Sub


