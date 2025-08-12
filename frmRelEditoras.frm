VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelEditoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Vendas por Editora"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelEditoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   7320
   Begin VB.Data datEditoras 
      Caption         =   "datEditoras"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Pesquisa 2] ORDER BY Nome"
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
      Top             =   5400
      Width           =   3615
      Begin VB.OptionButton optOrdemCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optOrdemNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
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
      Height          =   645
      Left            =   120
      TabIndex        =   45
      Top             =   4680
      Width           =   5415
      Begin VB.CheckBox chkSepararData 
         Caption         =   "Separar por data"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   47
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "até:"
         Height          =   255
         Left            =   1920
         TabIndex        =   46
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
      Top             =   5475
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
      Height          =   855
      Left            =   3840
      TabIndex        =   16
      Top             =   5400
      Width           =   1695
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
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
      TabIndex        =   41
      Top             =   -120
      Width           =   9615
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "frmRelEditoras.frx":058A
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Vendas"
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
         TabIndex        =   44
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelEditoras.frx":23F2
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   2160
         TabIndex        =   43
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "• Caso não queira utilizar algum filtro, basta não preencher o campo"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   1200
         Width           =   5175
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
      Top             =   7200
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
      Top             =   7560
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
      Top             =   7200
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
      Top             =   7560
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
      Top             =   7200
      Width           =   2295
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
      Height          =   3210
      Left            =   120
      TabIndex        =   25
      Top             =   1455
      Width           =   7095
      Begin VB.TextBox txtNomeEditora 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2760
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboEditora 
         Bindings        =   "frmRelEditoras.frx":24AD
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   2760
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   960
         Width           =   4455
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   28
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox txtNomeVendedor 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2400
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmRelEditoras.frx":24C7
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   2400
         Width           =   1335
         DataFieldList   =   "Nome"
         _Version        =   196617
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4101
         Columns(0).Caption=   "Codigo"
         Columns(0).Name =   "Codigo"
         Columns(0).DataField=   "Código"
         Columns(0).FieldLen=   256
         Columns(1).Width=   7355
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelEditoras.frx":24E3
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
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
      Begin SSDataWidgets_B.SSDBCombo cboProduto 
         Bindings        =   "frmRelEditoras.frx":24FD
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
         Bindings        =   "frmRelEditoras.frx":2517
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
         Bindings        =   "frmRelEditoras.frx":2532
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
      Begin SSDataWidgets_B.SSDBCombo cboCliente 
         Bindings        =   "frmRelEditoras.frx":254A
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
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelEditoras.frx":2564
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
      Begin VB.Label Label13 
         Caption         =   "Editora"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Sub-Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filial"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2070
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2430
         Width           =   855
      End
   End
   Begin VB.Data datVendedores 
      Caption         =   "datVendedores"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Funcionários WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   7560
      Width           =   2295
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
      Height          =   885
      Left            =   7560
      TabIndex        =   21
      Top             =   4320
      Width           =   3375
      Begin VB.CheckBox chkTipoNormal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTipoGrade 
         Caption         =   "Grade"
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkTipoEdicao 
         Caption         =   "Edição"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   360
         Value           =   1  'Checked
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   5880
      Width           =   1575
   End
   Begin Crystal.CrystalReport crpView 
      Left            =   240
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   6360
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmRelEditoras"
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

Private Sub cboCliente_CloseUp()
  cboCliente.Text = cboCliente.Columns(0).Text
  cboCliente_LostFocus
End Sub

Private Sub cboCliente_LostFocus()
  Dim rstClientes As Recordset
  
  txtNomeCliente.Text = ""
  If Not IsNumeric(cboCliente.Text) Then Exit Sub
  
  Set rstClientes = db.OpenRecordset("SELECT Código, Nome FROM Cli_For WHERE Código = " & cboCliente.Text, dbOpenSnapshot)
  
  With rstClientes
    If Not (.BOF And .EOF) Then
      txtNomeCliente.Text = .Fields("Nome") & ""
    End If
    
    If Not rstClientes Is Nothing Then .Close
    Set rstClientes = Nothing
  End With
End Sub

Private Sub cboEditora_CloseUp()
  cboEditora.Text = cboEditora.Columns(1).Text
  cboEditora_LostFocus
End Sub

Private Sub cboEditora_LostFocus()
  Dim rstEditoras As Recordset
  
  txtNomeEditora.Text = ""
  If Not IsNumeric(cboEditora.Text) Then Exit Sub
  
  Set rstEditoras = db.OpenRecordset(" SELECT Nome FROM [Pesquisa 2] " & _
                                         " WHERE Código = " & cboEditora.Text, dbOpenDynaset)
  With rstEditoras
    If Not (.BOF And .EOF) Then
      txtNomeEditora.Text = .Fields("Nome") & ""
    End If
    
    .Close
    Set rstEditoras = Nothing
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

Private Sub cboVendedor_CloseUp()
  cboVendedor.Text = cboVendedor.Columns(0).Text
  cboVendedor_LostFocus
End Sub

Private Sub cboVendedor_LostFocus()
  Dim rstFuncionarios As Recordset
  
  txtNomeVendedor.Text = ""
  If Not IsNumeric(cboVendedor.Text) Then Exit Sub
  
  Set rstFuncionarios = db.OpenRecordset(" SELECT Nome FROM Funcionários " & _
                                         " WHERE Código = " & cboVendedor.Text, dbOpenSnapshot)
  With rstFuncionarios
    If Not (.BOF And .EOF) Then
      txtNomeVendedor.Text = .Fields("Nome") & ""
    End If
    
    .Close
    Set rstFuncionarios = Nothing
  End With
End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  Dim rstRelVendas              As Recordset
  Dim strSQL                    As String
  
  Dim dblQtdeTotalDev           As Double: dblQtdeTotalDev = 0
  Dim dblValorTotalDev          As Double: dblValorTotalDev = 0
  Dim dblTotalDescontoSubTotal  As Double: dblTotalDescontoSubTotal = 0
  
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
  
  dbTemp.Execute "DELETE * FROM tblRelVendas"
  
  '---[ Chamada das funções para geração da tabela temporária ]---'
    If chkTipoNormal.Value = vbChecked Then
      Call StatusMsg("Gerando as informações do tipo normal, aguarde . . . ")
      GeraNormal
    End If
    
    If chkTipoGrade.Value = vbChecked Then
      Call StatusMsg("Gerando as informações do tipo grade, aguarde . . . ")
      GeraGrade
    End If

'    If chkTipoEdicao.Value = vbChecked Then
'      Call StatusMsg("Gerando as informações do tipo edição, aguarde . . . ")
'      GeraEdicao
'    End If
    
    Call StatusMsg("")
  '---[ Chamada das funções para geração da tabela temporária ]---'
  
  Set rstRelVendas = dbTemp.OpenRecordset("SELECT * FROM tblRelVendas", dbOpenSnapshot)
  
  With rstRelVendas
    If Not (.BOF And .EOF) Then
      '---[ Gera o total de Descontos do sub-total ]---'
        Call StatusMsg("Analisando descontos no sub-total e devoluções, aguarde . . . ")
        ReturnDescontoSubTotal dblTotalDescontoSubTotal
        ReturnDevolucaoNormal dblValorTotalDev, dblQtdeTotalDev
        ReturnDevolucaoGrade dblValorTotalDev, dblQtdeTotalDev
      '---[ Gera o total de Descontos do sub-total ]---'
      
      With crpView
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        
        If optSaidaVideo.Value Then .Destination = crptToWindow
        If optSaidaImpressora.Value Then .Destination = crptToPrinter
        
        .SortFields(0) = "-{tblRelVendas.proTipo}"
        
        If optOrdemCodigo.Value Then .SortFields(1) = "+{Produtos.Código Ordenação}"
        If optOrdemNome.Value Then .SortFields(1) = "+{Produtos.Nome}"
        If optRankingUnidade.Value Then .SortFields(1) = "+{tblRelVendas.venQtde}"
        If optRankingValor.Value Then .SortFields(1) = "+{tblRelVendas.venValor}"
        
        If chkSepararData.Value = vbChecked Then
          .ReportFileName = gsReportPath & "rptVendasPorData.rpt"
        Else
          .ReportFileName = gsReportPath & "rptVendasEditora.rpt"
        End If
        
        ' Modelo 1 ou 2
        'SetPrinterModeloPwd2 crpView
        
        .DataFiles(0) = gsTempDBFileName
        .DataFiles(1) = gsQuickDBFileName
        .DataFiles(2) = gsQuickDBFileName
        .DataFiles(3) = gsTempDBFileName
        .DataFiles(4) = gsQuickDBFileName
        .DataFiles(5) = gsQuickDBFileName
        .DataFiles(6) = gsQuickDBFileName
        
        .Formulas(0) = "DescSubTotal = " & Replace(Format(CStr(dblTotalDescontoSubTotal), "###0.00"), gsCurrencyDecimal, ".")
        .Formulas(1) = "DevolucoesQtde = " & Replace(Format(CStr(dblQtdeTotalDev), "###0.00"), gsCurrencyDecimal, ".")
        .Formulas(2) = "DevolucoesValor = " & Replace(Format(CStr(dblValorTotalDev), "###0.00"), gsCurrencyDecimal, ".")
        '---[ Preenchimento dos campos de cabeçalho de filtro ]---'
        .Formulas(3) = "Periodo = '" & "De " & mskDataInicio.Text & " até " & mskDataFinal.Text & "'"
          
          If Len(Trim(txtNomeFilial.Text)) > 0 Then .Formulas(4) = "Filtro_Filial = '" & txtNomeFilial.Text & "'"
          If Len(Trim(txtNomeCliente.Text)) > 0 Then .Formulas(5) = "Filtro_Cliente = '" & txtNomeCliente.Text & "'"
          If Len(Trim(txtNomeClasse.Text)) > 0 Then .Formulas(6) = "Filtro_Classe = '" & txtNomeClasse.Text & "'"
          If Len(Trim(txtNomeSubClasse.Text)) > 0 Then .Formulas(7) = "Filtro_SubClasse = '" & txtNomeSubClasse.Text & "'"
          If Len(Trim(txtNomeProduto.Text)) > 0 Then .Formulas(8) = "Filtro_Produto = '" & txtNomeProduto.Text & "'"
          If Len(Trim(txtNomeFornecedor.Text)) > 0 Then .Formulas(9) = "Filtro_Fornecedor = '" & txtNomeFornecedor.Text & "'"
          If Len(Trim(txtNomeVendedor.Text)) > 0 Then .Formulas(10) = "Filtro_Vendedor = '" & txtNomeVendedor.Text & "'"
        '---[ Preenchimento dos campos de cabeçalho de filtro ]---'
        
        '10/05/2004 - Daniel
        'Tratamento de 05 casas decimais após a ','
        'quando Embalavi
        If g_bln5CasasDecimais Then
            .Formulas(11) = "QtdeCasasDecimais = " & "5"
        '30/04/2007 - Anderson - Implementação de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
            .Formulas(11) = "QtdeCasasDecimais = " & "3"
        Else
            .Formulas(11) = "QtdeCasasDecimais = " & "2"
        End If
        
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
  
  Call StatusMsg("")
End Sub

Private Sub GeraNormal()
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstVendas         As Recordset
  Dim rstRelVendas      As Recordset
  Dim rstProdutos       As Recordset
  
  Dim blnProdutoOK      As Boolean
  
  strSQL = " SELECT Saídas.Filial, Saídas.Data, [Saídas - Produtos].[Código sem Grade], Sum([Saídas - Produtos].Qtde) AS SomaQtde, Sum([Saídas - Produtos].[Preço Final]) AS SomaPrecoFinal, [Operações Saída].Tipo "
  strSQL = strSQL & " FROM ((Saídas INNER JOIN [Saídas - Produtos] ON (Saídas.Sequência = [Saídas - Produtos].Sequência) AND (Saídas.Filial = [Saídas - Produtos].Filial)) INNER JOIN Produtos ON [Saídas - Produtos].[Código sem Grade] = Produtos.Código) INNER JOIN [Operações Saída] ON Saídas.Operação = [Operações Saída].Código "
  strSQL = strSQL & " GROUP BY Saídas.Filial, [Saídas - Produtos].[Código sem Grade], Saídas.Efetivada, Saídas.[Nota Cancelada], [Operações Saída].Tipo, Saídas.Data, Saídas.Data, Saídas.Filial, Saídas.Cliente, [Saídas - Produtos].[Código sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Saídas.Digitador "
  strSQL = strSQL & " Having ( Saídas.Efetivada ) AND ( NOT Saídas.[Nota Cancelada]) AND ( [Operações Saída].Tipo = 'V' ) AND Produtos.Tipo = 'N' "
  
  strSQL = strSQL & " AND (Saídas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Saídas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "

  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Cliente = " & cboCliente.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Saídas - Produtos].[Código sem Grade] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Digitador = " & cboVendedor.Text & " )"
  End If
  
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If (.BOF And .EOF) Then
      Exit Sub
    End If
  End With
  
  rstVendas.MoveLast
  rstVendas.MoveFirst
  
  pgbProgress.min = 0
  pgbProgress.Max = rstVendas.RecordCount + 1
  
'  Set rstRelVendas = dbTemp.OpenRecordset("SELECT * FROM tblRelVendas", dbOpenDynaset)

  ws.BeginTrans
  blnInTransaction = True
  
  With rstVendas
    .MoveFirst
    
    Do While Not rstVendas.EOF
      Set rstProdutos = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE Código = '" & .Fields("Código Sem Grade") & "' AND Tipo = 'N' ", dbOpenSnapshot)
      
      blnProdutoOK = Not (rstProdutos.BOF And rstProdutos.EOF)
      
      rstProdutos.Close
      Set rstProdutos = Nothing
      
      If blnProdutoOK Then
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("Código Sem Grade"))
        End If
      End If
      
      If blnProdutoOK Then
        strSQL = " SELECT * FROM tblRelVendas "
        strSQL = strSQL & " WHERE filID = " & .Fields("Filial")
        strSQL = strSQL & " AND proID = '" & .Fields("Código Sem Grade") & "' "
        strSQL = strSQL & " AND proTipo = 'N' "
        strSQL = strSQL & " AND tamID = 0 "
        strSQL = strSQL & " AND corID = 0 "
        If chkSepararData.Value = vbChecked Then strSQL = strSQL & " AND venData = #" & Format(.Fields("Data"), "mm/dd/yyyy") & "# "
        
        Set rstRelVendas = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        If (rstRelVendas.BOF And rstRelVendas.EOF) Then
          rstRelVendas.AddNew
          
          rstRelVendas.Fields("filID") = .Fields("Filial")
          rstRelVendas.Fields("proID") = .Fields("Código Sem Grade")
          rstRelVendas.Fields("proTipo") = "N"
          rstRelVendas.Fields("tamID") = 0
          rstRelVendas.Fields("corID") = 0
          
          If chkSepararData.Value = vbChecked Then
            rstRelVendas.Fields("venData") = .Fields("Data")
          Else
            rstRelVendas.Fields("venData") = Data_Atual
          End If
          
          rstRelVendas.Fields("venQtde") = .Fields("SomaQtde")
          '10/05/2004 - Daniel
          'Caso seja Embalavi, formataremos o valor para
          '5 casas após a vírgula
          If g_bln5CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((.Fields("SomaPrecoFinal")), "##,###,##0.00000")
          '30/04/2007 - Anderson - Implementação de 3 casas decimais
          ElseIf g_bln3CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((.Fields("SomaPrecoFinal")), "##,###,##0.000")
          Else 'Não Embalavi
            rstRelVendas.Fields("venValor") = .Fields("SomaPrecoFinal")
          End If
        Else
          rstRelVendas.Edit
          rstRelVendas.Fields("venQtde") = rstRelVendas.Fields("venQtde") + .Fields("SomaQtde")
          '10/05/2004 - Daniel
          'Caso seja Embalavi, formataremos o valor para
          '5 casas após a vírgula
          If g_bln5CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((rstRelVendas.Fields("venValor") + .Fields("SomaPrecoFinal")), "##,###,##0.00000")
          '30/04/2007 - Anderson - Implementação de 3 casas decimais
          ElseIf g_bln3CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((rstRelVendas.Fields("venValor") + .Fields("SomaPrecoFinal")), "##,###,##0.000")
          Else
            rstRelVendas.Fields("venValor") = rstRelVendas.Fields("venValor") + .Fields("SomaPrecoFinal")
          End If
        End If
        
        rstRelVendas.Update
        
        rstRelVendas.Close
        Set rstRelVendas = Nothing
      End If
      
      pgbProgress.Value = .AbsolutePosition
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

Private Sub GeraGrade()
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstVendas         As Recordset
  Dim rstRelVendas      As Recordset
  Dim rstProdutos       As Recordset
  
  Dim intTamanho        As Integer
  Dim intCor            As Integer
  
  Dim blnProdutoOK      As Boolean
  
  strSQL = " SELECT Saídas.Filial, Saídas.Data, [Saídas - Produtos].Código , [Saídas - Produtos].[Código sem Grade], Sum([Saídas - Produtos].Qtde) AS SomaQtde, Sum([Saídas - Produtos].[Preço Final]) AS SomaPrecoFinal, [Operações Saída].Tipo "
  strSQL = strSQL & " FROM ((Saídas INNER JOIN [Saídas - Produtos] ON (Saídas.Sequência = [Saídas - Produtos].Sequência) AND (Saídas.Filial = [Saídas - Produtos].Filial)) INNER JOIN Produtos ON [Saídas - Produtos].[Código sem Grade] = Produtos.Código) INNER JOIN [Operações Saída] ON Saídas.Operação = [Operações Saída].Código "
  strSQL = strSQL & " GROUP BY Saídas.Filial, [Saídas - Produtos].Código, [Saídas - Produtos].[Código sem Grade], Saídas.Efetivada, Saídas.[Nota Cancelada], [Operações Saída].Tipo, Saídas.Data, Saídas.Data, Saídas.Filial, Saídas.Cliente, [Saídas - Produtos].[Código sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Saídas.Digitador "
  strSQL = strSQL & " Having ( Saídas.Efetivada ) AND ( NOT Saídas.[Nota Cancelada]) AND ( [Operações Saída].Tipo = 'V' ) AND Produtos.Tipo = 'G' "
  
  strSQL = strSQL & " AND (Saídas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Saídas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "

  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Cliente = " & cboCliente.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Saídas - Produtos].[Código sem Grade] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Digitador = " & cboVendedor.Text & " )"
  End If
  
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If (.BOF And .EOF) Then
      Exit Sub
    End If
  End With
  
  rstVendas.MoveLast
  rstVendas.MoveFirst
  
  pgbProgress.min = 0
  pgbProgress.Max = rstVendas.RecordCount + 1
  
'  Set rstRelVendas = dbTemp.OpenRecordset("SELECT * FROM tblRelVendas", dbOpenDynaset)

  ws.BeginTrans
  blnInTransaction = True
  
  With rstVendas
    .MoveFirst
    
    Do While Not rstVendas.EOF
      Set rstProdutos = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE Código = '" & .Fields("Código Sem Grade") & "' AND Tipo = 'G' ", dbOpenSnapshot)
      
      blnProdutoOK = Not (rstProdutos.BOF And rstProdutos.EOF)
      
      rstProdutos.Close
      Set rstProdutos = Nothing
      
      If blnProdutoOK Then
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("Código Sem Grade"))
        End If
      End If
      
      If blnProdutoOK Then
        strSQL = " SELECT * FROM tblRelVendas "
        strSQL = strSQL & " WHERE filID = " & .Fields("Filial")
        strSQL = strSQL & " AND proID = '" & .Fields("Código Sem Grade") & "' "
        strSQL = strSQL & " AND proTipo = 'G' "
        strSQL = strSQL & " AND tamID = " & intTamanho
        strSQL = strSQL & " AND corID = " & intCor
        If chkSepararData.Value = vbChecked Then strSQL = strSQL & " AND venData = #" & Format(.Fields("Data"), "mm/dd/yyyy") & "# "
        
        Set rstRelVendas = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        If (rstRelVendas.BOF And rstRelVendas.EOF) Then
          rstRelVendas.AddNew
          
          intTamanho = Left(Right(.Fields("Código"), 6), 3)
          intCor = Right(.Fields("Código"), 3)
          
          rstRelVendas.Fields("filID") = .Fields("Filial")
          rstRelVendas.Fields("proID") = .Fields("Código Sem Grade")
          rstRelVendas.Fields("proTipo") = "G"
          rstRelVendas.Fields("tamID") = intTamanho
          rstRelVendas.Fields("corID") = intCor
          
          If chkSepararData.Value = vbChecked Then
            rstRelVendas.Fields("venData") = .Fields("Data")
          Else
            rstRelVendas.Fields("venData") = Data_Atual
          End If
          
          rstRelVendas.Fields("venQtde") = .Fields("SomaQtde")
          rstRelVendas.Fields("venValor") = .Fields("SomaPrecoFinal")
        Else
          rstRelVendas.Edit
          rstRelVendas.Fields("venQtde") = rstRelVendas.Fields("venQtde") + .Fields("SomaQtde")
          rstRelVendas.Fields("venValor") = rstRelVendas.Fields("venValor") + .Fields("SomaPrecoFinal")
        End If
        
        rstRelVendas.Update
        
        rstRelVendas.Close
        Set rstRelVendas = Nothing
      End If
      
      pgbProgress.Value = rstVendas.AbsolutePosition
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
  Call CenterForm(Me)
  
  datFiliais.DatabaseName = gsQuickDBFileName
  datClientes.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
  datVendedores.DatabaseName = gsQuickDBFileName
  datEditoras.DatabaseName = gsQuickDBFileName
End Sub

Private Function blnVerificaForncedor(strCodigoProduto As String) As Boolean
  Dim rstFornProd As Recordset

  Set rstFornProd = db.OpenRecordset(" SELECT * FROM Forn_Prod " & _
                                     " WHERE Produto = '" & strCodigoProduto & "' " & _
                                     " AND Fornecedor = " & CLng(cboFornecedor.Text), dbOpenSnapshot)
  
  With rstFornProd
    blnVerificaForncedor = Not (.BOF And .EOF)
    
    rstFornProd.Close
    Set rstFornProd = Nothing
  End With
End Function

Private Function ReturnDevolucaoNormal(ByRef dblValorDevolucao As Double, _
                                       ByRef dblQtdeDevolucao As Double) As Boolean
  
  Dim strSQL As String
  Dim rstDev As Recordset
  Dim blnProdutoOK As Boolean
  
  Dim rstProdutos As Recordset
  Dim rstGrade As Recordset
  
  Dim strCodigoProduto As String
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Entradas - Produtos].Código, Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal " & _
           " FROM ((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Sequência = [Entradas - Produtos].Sequência) AND (Entradas.Filial = [Entradas - Produtos].Filial)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN Produtos ON [Entradas - Produtos].Código = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Digitador, [Entradas - Produtos].Código, Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Operações Entrada].Tipo)='D')) "

  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboCliente.Text & ") "
  End If
  
  '15/03/2004 - Daniel
  'Não estava fazendo o filtro por Vendedor
  'Foi acrescentado também esta linha no GROUP BY:
  'Entradas.Digitador
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Digitador = " & cboVendedor.Text & ") "
  End If
  
  If Len(Trim(cboProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Entradas - Produtos].Código = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(cboClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(cboSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  '15/03/2004 - Daniel
  'Não estava fazendo o filtro por Vendedor
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND Entradas.Digitador =" & CInt(cboVendedor.Text)
  End If
  
  
  Set rstDev = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstDev
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        blnProdutoOK = True
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("Código"))
        End If
        
        If blnProdutoOK Then
          dblValorDevolucao = dblValorDevolucao + CDbl(.Fields("PrecoTotal"))
          dblQtdeDevolucao = dblQtdeDevolucao + CDbl(.Fields("ContarDeQtde"))
        End If
        .MoveNext
      Loop
    End If
  End With
End Function

Private Function ReturnDevolucaoGrade(ByRef dblValorDevolucao As Double, _
                                      ByRef dblQtdeDevolucao As Double) As Boolean
  
  Dim strSQL As String
  Dim rstDev As Recordset
  Dim blnProdutoOK As Boolean
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Códigos da Grade].[Código Original], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Preço Final]) AS PrecoTotal " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequência = [Entradas - Produtos].Sequência)) INNER JOIN [Operações Entrada] ON Entradas.Operação = [Operações Entrada].Código) INNER JOIN [Códigos da Grade] ON [Entradas - Produtos].Código = [Códigos da Grade].Código) INNER JOIN Produtos ON [Códigos da Grade].[Código Original] = Produtos.Código " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Digitador, [Códigos da Grade].[Código Original], Entradas.Fornecedor, [Operações Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Operações Entrada].Tipo)='D')) "


  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboCliente.Text & ") "
  End If
  
  '15/03/2004 - Daniel
  'Não estava fazendo o filtro por Vendedor
  'Foi acrescentado também esta linha no GROUP BY:
  'Entradas.Digitador
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Digitador = " & cboVendedor.Text & ") "
  End If
  
  If Len(Trim(cboProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Códigos da Grade].[Código Original] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(cboClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(cboSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  
  Set rstDev = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstDev
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do While Not .EOF
        blnProdutoOK = True
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("Código Original"))
        End If
        
        If blnProdutoOK Then
          dblValorDevolucao = dblValorDevolucao + CDbl(.Fields("PrecoTotal"))
          dblQtdeDevolucao = dblQtdeDevolucao + CDbl(.Fields("ContarDeQtde"))
        End If
        
        .MoveNext
      Loop
    End If
  End With
End Function

Private Function ReturnDescontoSubTotal(ByRef dblValorDesconto As Double) As Double
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstVendas         As Recordset
  Dim rstProdutos       As Recordset
  Dim rstDescontoSubTotal As Recordset
  
  Dim dblDescontoSubTotal As Double
  Dim dblDescontoSomar  As Double
  Dim blnProdutoOK      As Boolean
  
  strSQL = " SELECT SUM(Saídas.DescontoSubTotal) AS DescontoSubTotal, [Saídas - Produtos].[Código sem Grade], Saídas.Filial, Saídas.Sequência "
  strSQL = strSQL & " FROM ((Saídas INNER JOIN [Saídas - Produtos] ON (Saídas.Sequência = [Saídas - Produtos].Sequência) AND (Saídas.Filial = [Saídas - Produtos].Filial)) INNER JOIN Produtos ON [Saídas - Produtos].[Código sem Grade] = Produtos.Código) INNER JOIN [Operações Saída] ON Saídas.Operação = [Operações Saída].Código "
  strSQL = strSQL & " GROUP BY Saídas.Filial, Saídas.Data, Saídas.Cliente, [Saídas - Produtos].[Código sem Grade], Saídas.Digitador, Produtos.Classe, Produtos.[Sub Classe], Saídas.Efetivada, Saídas.[Nota Cancelada], [Operações Saída].Tipo = 'V', Saídas.Sequência, Saídas.DescontoSubTotal "
  strSQL = strSQL & " HAVING ( Saídas.Efetivada ) AND ( NOT Saídas.[Nota Cancelada]) AND ( [Operações Saída].Tipo = 'V' ) AND Saídas.DescontoSubTotal > 0"
  
  strSQL = strSQL & " AND (Saídas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Saídas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Cliente = " & cboCliente.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Saídas - Produtos].[Código sem Grade] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Saídas.Digitador = " & cboVendedor.Text & " )"
  End If
  
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  
  With rstVendas
    If (.BOF And .EOF) Then
      Exit Function
    End If
    
    .MoveLast
    .MoveFirst
    
    pgbProgress.min = 0
    pgbProgress.Max = .RecordCount + 1
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
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("Código Sem Grade"))
        End If
      End If
      
      If blnProdutoOK Then
        dblDescontoSubTotal = dblDescontoSubTotal + dblDescontoSomar
      End If
      
      pgbProgress.Value = .AbsolutePosition
      .MoveNext
    Loop
  End With
  
  dblValorDesconto = dblDescontoSubTotal
  
  If Not rstVendas Is Nothing Then rstVendas.Close
  Set rstVendas = Nothing
End Function

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
