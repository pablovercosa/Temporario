VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelLocalizacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Localização de Produtos"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmRelLocalizacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7575
   Begin Crystal.CrystalReport rptLocalizacao 
      Left            =   4920
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
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
      Left            =   5640
      TabIndex        =   6
      Top             =   3285
      Width           =   1815
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
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
      Height          =   885
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   5415
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
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
         TabIndex        =   4
         Top             =   520
         Width           =   1215
      End
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
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
         TabIndex        =   3
         Top             =   260
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   7335
      Begin VB.Data datTabela 
         Caption         =   "datTabela"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Tabela FROM [Tabela de Preços] ORDER BY Tabela"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2340
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   255
         Width           =   3615
      End
      Begin VB.TextBox txtProduto 
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Width           =   3615
      End
      Begin VB.Data datProdutos 
         Caption         =   "datProdutos"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' ORDER BY Código"
         Top             =   600
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Data datFornecedor 
         Caption         =   "datFornecedor"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Código, Nome FROM Cli_For WHERE Tipo = 'F' ORDER BY Código"
         Top             =   120
         Visible         =   0   'False
         Width           =   1740
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelLocalizacao.frx":058A
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   255
         Width           =   2320
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   4092
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboProduto 
         Bindings        =   "frmRelLocalizacao.frx":05A6
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   667
         Width           =   2325
         DataFieldList   =   "Código"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   4101
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboTabela 
         Bindings        =   "frmRelLocalizacao.frx":05C0
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   2325
         DataFieldList   =   "Tabela"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   4101
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   16777215
         DataFieldToDisplay=   "Tabela"
      End
      Begin VB.Label Label2 
         Caption         =   "Tabela"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "Fabricante"
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
         Left            =   120
         TabIndex        =   15
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Produto"
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
         Left            =   120
         TabIndex        =   14
         Top             =   727
         Width           =   810
      End
   End
   Begin VB.Frame fraDica 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   8175
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Localização de Produtos"
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
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblDica 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelLocalizacao.frx":05D8
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
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmRelLocalizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Relatório desenvolvido por Daniel R. Rodrigues
'22/04/2004 - Case Ortociso
'Otimizar a visualização onde é guardado os produtos,
'isto é, em quais gavetas (localizações)

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

Private Sub cboProduto_CloseUp()
  cboProduto.Text = cboProduto.Columns(0).Text
  cboProduto_LostFocus
End Sub

Private Sub cboProduto_LostFocus()
  Dim rstProdutos As Recordset

  txtProduto.Text = ""
  
  Set rstProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & (cboProduto.Text) & "'", dbOpenDynaset)

  With rstProdutos
    If Not (.BOF And .EOF) Then
      txtProduto.Text = .Fields("Nome") & ""
    End If
  End With

  rstProdutos.Close
  Set rstProdutos = Nothing

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim strReport  As String
  Dim strSQL     As String
  
  Call StatusMsg("")
    
  If Len(cboTabela.Text) = 0 Then
    MsgBox "Selecione uma tabela.", vbExclamation, "Quick Store"
    cboTabela.SetFocus
    Exit Sub
  End If
  
  If Len(txtProduto.Text) > 0 Then
    strSQL = " {Produtos.Código} = '" & (cboProduto.Text) & "'"
  Else
    strSQL = " {Produtos.Código} <> '0' "
  End If

  If Len(txtFornecedor.Text) > 0 Then
    strSQL = strSQL & " AND {Forn_Prod.Fornecedor} = " & CInt(cboFornecedor.Text)
  End If

  If Len(cboTabela.Text) > 0 Then
    strSQL = strSQL & " AND {Preços.Tabela} = '" & cboTabela.Text & "'"
    
    If Len(txtProduto.Text) > 0 Then
      strSQL = strSQL & " AND {Preços.Produto} = '" & (cboProduto.Text) & "'"
    End If
  End If
  
  'Nome do arquivo .rpt
  strReport = gsReportPath & "rptProdutosLocalizacao.rpt"
  MousePointer = vbHourglass
  
  With rptLocalizacao
    .Reset
    .ReportFileName = strReport
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 rptLocalizacao
    
    'Quatro tabelas do QS
    'Uma tabela do QTemp
    .DataFiles(0) = gsQuickDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsQuickDBFileName
    .DataFiles(3) = gsQuickDBFileName
    .DataFiles(4) = gsTempDBFileName
    
    .SelectionFormula = strSQL
    .Formulas(0) = "nome_empresa = '" & gsNomeEmpresa & "'" 'Cadastra a fórmula no crystal também
    .SortFields(0) = "+{Forn_Prod.Fornecedor}" 'Ordenação
    .SortFields(1) = "+{Produtos.Código}"
    
    .WindowState = crptMaximized
    .Destination = IIf(optSaidaVideo.Value, crptToWindow, crptToPrinter)
    Call StatusMsg("Aguarde, imprimindo...")
    
    'Seta a impressora para relatório
    Call SetPrinterName("REL", rptLocalizacao)
  
    .Action = 1
  End With
  
  MousePointer = vbDefault
  
  Call StatusMsg("")

End Sub

Private Sub Form_Load()
  datFornecedor.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
  datTabela.DatabaseName = gsQuickDBFileName

  Call CenterForm(Me)
End Sub
