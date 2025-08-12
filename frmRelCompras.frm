VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Compras"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   7305
   Begin Crystal.CrystalReport crpView 
      Left            =   240
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tipo dos Produtos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   3840
      TabIndex        =   7
      Top             =   3840
      Width           =   3375
      Begin VB.CheckBox chkTipoNormal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTipoGrade 
         Caption         =   "Grade"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkTipoEdicao 
         Caption         =   "Edição"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   900
      End
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Filial, Nome FROM [Parâmetros Filial]"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Data datClientes 
      Caption         =   "datClientes"
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
      RecordSource    =   "SELECT Código, Nome FROM Cli_For ORDER BY Nome"
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Data datClasse 
      Caption         =   "datClasse"
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Classes ORDER BY Nome"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Data datSubClasse 
      Caption         =   "datSubClasse"
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM [Sub Classes] ORDER BY Nome"
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Data datProdutos 
      Caption         =   "datProdutos"
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
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0' ORDER BY Nome"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Frame Frame6 
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
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   3615
      Begin VB.OptionButton optOrdemCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optOrdemNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optRankingUnidade 
         Caption         =   "Ranking por unidade"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optRankingValor 
         Caption         =   "Ranking por valor"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   650
      Left            =   120
      TabIndex        =   36
      Top             =   3840
      Width           =   3615
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   240
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
      Begin MSMask.MaskEdBox mskDataInicio 
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   240
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
      Begin VB.Label Label3 
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "até:"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Imprimir"
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
      TabIndex        =   19
      Top             =   4640
      Width           =   1575
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
      Height          =   885
      Left            =   3840
      TabIndex        =   16
      Top             =   4560
      Width           =   1695
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1215
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
      Left            =   -120
      TabIndex        =   33
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "• Caso não queira utilizar algum filtro, basta não preencher o campo"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2280
         TabIndex        =   39
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "frmRelCompras.frx":058A
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Compras"
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
         TabIndex        =   35
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Utilize os campos abaixo como filtro e ordenação do relatório. O relatório mostra todos os produtos que tiveram compra no período."
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   2280
         TabIndex        =   34
         Top             =   600
         Width           =   5055
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
      Height          =   2175
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   7095
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtNomeSubClasse 
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
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtNomeClasse 
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   960
         Width           =   4455
      End
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
         TabIndex        =   23
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelCompras.frx":23F2
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1335
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
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboProduto 
         Bindings        =   "frmRelCompras.frx":240C
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
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
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmRelCompras.frx":2426
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
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
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmRelCompras.frx":2441
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   1335
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
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelCompras.frx":2459
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
      Begin VB.Label Label7 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Sub-Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filial"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   630
         Width           =   975
      End
   End
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   5520
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "frmRelCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboClasse_CloseUp()
  cboClasse.Text = cboClasse.Columns(0).Text
  cboClasse_LostFocus
End Sub

Private Sub cboClasse_LostFocus()
  Dim rstClasses As Recordset
  
  txtNomeClasse.Text = ""
  
  If Not IsNumeric(cboClasse.Text) Then Exit Sub
  
  Set rstClasses = db.OpenRecordset("SELECT Código, Nome FROM Classes WHERE Código = " & cboClasse.Text, dbOpenSnapshot)
  
  With rstClasses
    If Not (.BOF And .EOF) Then
      txtNomeClasse.Text = .Fields("Nome") & ""
    End If
    
    If Not rstClasses Is Nothing Then .Close
    Set rstClasses = Nothing
  End With
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
  
  datClientes.Recordset.FindFirst "Código = " & cboFornecedor.Text
  
  If Not datClientes.Recordset.NoMatch Then
    txtNomeFornecedor.Text = datClientes.Recordset.Fields("Nome") & ""
  End If
End Sub

Private Sub cboProduto_CloseUp()
  cboProduto.Text = cboProduto.Columns(0).Text
  cboProduto_LostFocus
End Sub

Private Sub cboProduto_LostFocus()
  Dim rstProdutos As Recordset
  
  txtNomeProduto.Text = ""
  
  Set rstProdutos = db.OpenRecordset("SELECT Código, Nome FROM Produtos WHERE Código = '" & cboProduto.Text & "' AND Código <> '0' ", dbOpenSnapshot)
  
  With rstProdutos
    If Not (.BOF And .EOF) Then
      txtNomeProduto.Text = .Fields("Nome") & ""
    End If
    
    If Not rstProdutos Is Nothing Then .Close
    Set rstProdutos = Nothing
  End With
End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.Text = cboSubClasse.Columns(0).Text
  cboSubClasse_LostFocus
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim rstSubClasses As Recordset
  
  txtNomeSubClasse.Text = ""
  
  If Not IsNumeric(cboSubClasse.Text) Then Exit Sub
  
  Set rstSubClasses = db.OpenRecordset("SELECT Código, Nome FROM [Sub Classes] WHERE Código = " & cboSubClasse.Text, dbOpenSnapshot)
  
  With rstSubClasses
    If Not (.BOF And .EOF) Then
      txtNomeSubClasse.Text = .Fields("Nome") & ""
    End If
    
    If Not rstSubClasses Is Nothing Then .Close
    Set rstSubClasses = Nothing
  End With
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim rstRelCompras As Recordset
  Dim strSQL       As String
  
  Dim dblQtdeTotalDev   As Double: dblQtdeTotalDev = 0
  Dim dblValorTotalDev  As Double: dblValorTotalDev = 0
  
  If (Not IsDate(mskDataInicio.Text)) And (Not IsDate(mskDataFinal.Text)) Then
    mskDataInicio.Text = Data_Atual
    mskDataFinal.Text = Data_Atual
  End If
  
  If Not IsDate(mskDataInicio.Text) Then
    MsgBox "Data inicial inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(mskDataFinal.Text) Then
    MsgBox "Data final inválida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(mskDataInicio.Text) > CDate(mskDataFinal.Text) Then
    MsgBox "A data inicial não pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  dbTemp.Execute "DELETE * FROM tblRelCompras"
  
  '---[ Chamada das funções para geração da tabela temporária ]---'
    If chkTipoNormal.Value = vbChecked Then
      Call StatusMsg("Gerando as informações do tipo normal, aguarde . . . ")
      GeraNormal
    End If
    
    If chkTipoGrade.Value = vbChecked Then
      Call StatusMsg("Gerando as informações do tipo grade, aguarde . . . ")
      GeraGrade
    End If
    
    If chkTipoEdicao.Value = vbChecked Then
      Call StatusMsg("Gerando as informações do tipo edição, aguarde . . . ")
      GeraEdicao
    End If
    
    Call StatusMsg("")
  '---[ Chamada das funções para geração da tabela temporária ]---'
  
  Set rstRelCompras = dbTemp.OpenRecordset("SELECT * FROM tblRelCompras", dbOpenSnapshot)
  
  With rstRelCompras
    If Not (.BOF And .EOF) Then
      With crpView
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        
        
        If optSaidaVideo.Value Then .Destination = crptToWindow
        If optSaidaImpressora.Value Then .Destination = crptToPrinter
        
        .SortFields(0) = "-{tblRelCompras.proTipo}"
        
        If optOrdemCodigo.Value Then .SortFields(1) = "+{Produtos.Código Ordenação}"
        If optOrdemNome.Value Then .SortFields(1) = "+{Produtos.Nome}"
        If optRankingUnidade.Value Then .SortFields(1) = "+{tblRelCompras.comQtde}"
        If optRankingValor.Value Then .SortFields(1) = "+{tblRelCompras.comValor}"
        
        .ReportFileName = gsReportPath & "rptCompras.rpt"
        
        ' Modelo 1 ou 2
        'SetPrinterModeloPwd2 crpView
        
        .DataFiles(0) = gsQuickDBFileName
        .DataFiles(1) = gsQuickDBFileName
        .DataFiles(2) = gsTempDBFileName
        .DataFiles(3) = gsQuickDBFileName
        .DataFiles(4) = gsQuickDBFileName
        .DataFiles(5) = gsTempDBFileName
        
        .Formulas(0) = "Periodo = '" & "De " & mskDataInicio.Text & " até " & mskDataFinal.Text & "'"
        
        If Len(Trim(txtNomeFilial.Text)) > 0 Then .Formulas(1) = "Filtro_Filial = '" & txtNomeFilial.Text & "'"
        If Len(Trim(txtNomeClasse.Text)) > 0 Then .Formulas(2) = "Filtro_Classe = '" & txtNomeClasse.Text & "'"
        If Len(Trim(txtNomeSubClasse.Text)) > 0 Then .Formulas(3) = "Filtro_SubClasse = '" & txtNomeSubClasse.Text & "'"
        If Len(Trim(txtNomeProduto.Text)) > 0 Then .Formulas(4) = "Filtro_Produto = '" & txtNomeProduto.Text & "'"
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then .Formulas(5) = "Filtro_Fornecedor = '" & txtNomeFornecedor.Text & "'"
        
        '25/07/2003 - mpdea
        'Seta a impressora para relatório
        Call SetPrinterName("REL", crpView)
        
        .Action = 1
        pgbProgress.Value = 0
      End With
    Else
      MsgBox "Não existem informações a serem exibidas !", vbInformation, App.Title
    End If
  End With
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
  
  datFiliais.DatabaseName = gsQuickDBFileName
  datClientes.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
End Sub

Private Sub mskDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataFinal.Text = frmCalendario.gsDateCalender(mskDataFinal.Text)
  End If
End Sub

Private Sub mskDataFinal_LostFocus()
  mskDataFinal.Text = Ajusta_Data(mskDataFinal.Text)
End Sub

Private Sub mskDataInicio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    mskDataInicio.Text = frmCalendario.gsDateCalender(mskDataInicio.Text)
  End If
End Sub

Private Sub mskDataInicio_LostFocus()
  mskDataInicio.Text = Ajusta_Data(mskDataInicio.Text)
End Sub

Private Sub GeraNormal()
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstCompras         As Recordset
  Dim rstRelCompras      As Recordset
  Dim rstGrade          As Recordset
  Dim rstProdutos       As Recordset
  
  Dim strCodigoProduto  As String
  Dim intTamanho        As Integer
  Dim intCor            As Integer
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Entradas - Produtos].Código, Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal " & _
           " FROM ((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Sequência = [Entradas - Produtos].Sequência) AND (Entradas.Filial = [Entradas - Produtos].Filial)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN Produtos ON [Entradas - Produtos].Código = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, [Entradas - Produtos].Código, Entradas.Fornecedor, [Operações Entrada].Tipo, Entradas.Efetivada, Entradas.[Nota Cancelada], Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING (([Operações Entrada].Tipo)='C') AND ( Entradas.Efetivada = TRUE ) AND (Entradas.[Nota Cancelada] = FALSE ) "
  
  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboFornecedor.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Entradas - Produtos].Código = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  Set rstCompras = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstCompras
    If (.BOF And .EOF) Then
      Exit Sub
    End If
  End With
  
  rstCompras.MoveLast
  rstCompras.MoveFirst
  
  pgbProgress.min = 0
  pgbProgress.Max = rstCompras.RecordCount + 1
  
  Set rstRelCompras = dbTemp.OpenRecordset("SELECT * FROM tblRelCompras", dbOpenDynaset)

  ws.BeginTrans
  blnInTransaction = True
  
  With rstRelCompras
    rstCompras.MoveFirst
    
    Do While Not rstCompras.EOF
      Set rstProdutos = db.OpenRecordset("SELECT Código FROM Produtos WHERE Código = '" & _
                                         rstCompras.Fields("Código") & "'", dbOpenSnapshot)
      
      intTamanho = 0
      intCor = 0
      
      With rstProdutos
        If Not (.BOF And .EOF) Then
          strCodigoProduto = .Fields("Código")
        Else
          Set rstGrade = db.OpenRecordset(" SELECT * FROM [Códigos da Grade] " & _
                                          " WHERE Código = '" & rstCompras.Fields("Código") & "'", dbOpenSnapshot)
          
          With rstGrade
            If Not (.BOF And .EOF) Then
              strCodigoProduto = .Fields("Código Original")
              intTamanho = Left(Right(.Fields("Código"), 6), 3)
              intCor = Right(.Fields("Código"), 3)
            Else
              strCodigoProduto = "0"
            End If
          End With
        End If
      End With
      
      .AddNew
      
      .Fields("filID") = rstCompras.Fields("Filial")
      .Fields("proID") = strCodigoProduto
      .Fields("proTipo") = "N"
      .Fields("tamID") = intTamanho
      .Fields("corID") = intCor
      .Fields("comData") = rstCompras.Fields("Data")
      .Fields("comQtde") = rstCompras.Fields("ContarDeQtde")
      .Fields("comValor") = rstCompras.Fields("PrecoTotal")
      
      .Update
      
      pgbProgress.Value = rstCompras.AbsolutePosition
      rstCompras.MoveNext
    Loop
  End With
  
  ws.CommitTrans
  blnInTransaction = False
  
  If Not rstRelCompras Is Nothing Then rstRelCompras.Close
  Set rstRelCompras = Nothing
  
  If Not rstCompras Is Nothing Then rstCompras.Close
  Set rstCompras = Nothing
End Sub

Private Sub GeraGrade()
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstCompras         As Recordset
  Dim rstRelCompras      As Recordset
  Dim rstGrade          As Recordset
  Dim rstProdutos       As Recordset
  
  Dim strCodigoProduto  As String
  Dim intTamanho        As Integer
  Dim intCor            As Integer
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Entradas - Produtos].Código, Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal, [Códigos da Grade].[Código Original] " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequência = [Entradas - Produtos].Sequência)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN [Códigos da Grade] ON [Entradas - Produtos].Código = [Códigos da Grade].Código) INNER JOIN Produtos ON [Códigos da Grade].[Código Original] = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, [Entradas - Produtos].Código, [Códigos da Grade].[Código Original], Entradas.Fornecedor, [Operações Entrada].Tipo, Entradas.Efetivada, Entradas.[Nota Cancelada], Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING (([Operações Entrada].Tipo)='C') AND ( Entradas.Efetivada = TRUE ) AND (Entradas.[Nota Cancelada] = FALSE ) "

  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboFornecedor.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Códigos da Grade].[Código Original] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  Set rstCompras = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstCompras
    If (.BOF And .EOF) Then
      Exit Sub
    End If
  End With
  
  rstCompras.MoveLast
  rstCompras.MoveFirst
  
  pgbProgress.min = 0
  pgbProgress.Max = rstCompras.RecordCount + 1
  
  Set rstRelCompras = dbTemp.OpenRecordset("SELECT * FROM tblRelCompras", dbOpenDynaset)

  ws.BeginTrans
  blnInTransaction = True
  
  With rstRelCompras
    rstCompras.MoveFirst
    
    Do While Not rstCompras.EOF
      Set rstProdutos = db.OpenRecordset("SELECT Código FROM Produtos WHERE Código = '" & _
                                         rstCompras.Fields("Código") & "'", dbOpenSnapshot)
      
      intTamanho = 0
      intCor = 0
      
      With rstProdutos
        If Not (.BOF And .EOF) Then
          strCodigoProduto = .Fields("Código")
        Else
          Set rstGrade = db.OpenRecordset(" SELECT * FROM [Códigos da Grade] " & _
                                          " WHERE Código = '" & rstCompras.Fields("Código") & "'", dbOpenSnapshot)
          
          With rstGrade
            If Not (.BOF And .EOF) Then
              strCodigoProduto = .Fields("Código Original")
              intTamanho = Left(Right(.Fields("Código"), 6), 3)
              intCor = Right(.Fields("Código"), 3)
            Else
              strCodigoProduto = "0"
            End If
          End With
        End If
      End With
      
      .AddNew
      
      .Fields("filID") = rstCompras.Fields("Filial")
      .Fields("proID") = strCodigoProduto
      .Fields("proTipo") = "G"
      .Fields("tamID") = intTamanho
      .Fields("corID") = intCor
      .Fields("comData") = rstCompras.Fields("Data")
      .Fields("comQtde") = rstCompras.Fields("ContarDeQtde")
      .Fields("comValor") = rstCompras.Fields("PrecoTotal")
      
      .Update
      
      pgbProgress.Value = rstCompras.AbsolutePosition
      rstCompras.MoveNext
    Loop
  End With
  
  ws.CommitTrans
  blnInTransaction = False
  
  If Not rstRelCompras Is Nothing Then rstRelCompras.Close
  Set rstRelCompras = Nothing
  
  If Not rstCompras Is Nothing Then rstCompras.Close
  Set rstCompras = Nothing
End Sub

Private Sub GeraEdicao()

End Sub
