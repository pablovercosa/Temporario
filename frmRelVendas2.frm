VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelVendas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relat�rio de Vendas"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   540
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
   Icon            =   "frmRelVendas2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   7320
   Begin VB.Data datOperacao 
      Caption         =   "datOperacao"
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
      RecordSource    =   "SELECT C�digo, Nome FROM [Opera��es Sa�da] ORDER BY Nome"
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Frame fraRelatorio 
      Caption         =   "Relat�rio"
      Height          =   855
      Left            =   5580
      TabIndex        =   50
      Top             =   4320
      Width           =   1695
      Begin VB.OptionButton optSintetico 
         Appearance      =   0  'Flat
         Caption         =   "Sint�tico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optAnalitico 
         Appearance      =   0  'Flat
         Caption         =   "Anal�tico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdFechar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6120
      Width           =   3555
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tipo dos Produtos"
      Height          =   885
      Left            =   3690
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
      Begin VB.CheckBox chkTipoEdicao 
         Appearance      =   0  'Flat
         Caption         =   "Edi��o"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.CheckBox chkTipoGrade 
         Appearance      =   0  'Flat
         Caption         =   "Grade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkTipoNormal 
         Appearance      =   0  'Flat
         Caption         =   "Normal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
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
      RecordSource    =   "SELECT C�digo, Nome FROM Funcion�rios WHERE Liberado = TRUE AND Ativo = TRUE ORDER BY Nome"
      Top             =   7200
      Width           =   2295
   End
   Begin Crystal.CrystalReport crpView 
      Left            =   240
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame4 
      Height          =   3090
      Left            =   30
      TabIndex        =   33
      Top             =   1185
      Width           =   7245
      Begin VB.TextBox txtNomeOperacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Data datPrecos 
         Caption         =   "datPrecos"
         Connect         =   "Access 2000;"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5550
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT Tabela FROM [Tabela de Pre�os] WHERE Tabela <> 'CUSTO' ORDER BY Tabela"
         Top             =   2910
         Visible         =   0   'False
         Width           =   1875
      End
      Begin SSDataWidgets_B.SSDBCombo cboTabela 
         Bindings        =   "frmRelVendas2.frx":4E95A
         Height          =   315
         Left            =   5280
         TabIndex        =   8
         Top             =   2640
         Width           =   1695
         DataFieldList   =   "Tabela"
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Tabela"
         Columns(0).Name =   "Tabela"
         Columns(0).CaptionAlignment=   0
         Columns(0).DataField=   "Tabela"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
      End
      Begin VB.TextBox txtNomeVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2280
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmRelVendas2.frx":4E972
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   2280
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
         Columns(0).DataField=   "C�digo"
         Columns(0).FieldLen=   256
         Columns(1).Width=   7355
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.TextBox txtNomeFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelVendas2.frx":4E98E
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
         DataFieldList   =   "Nome"
         _Version        =   196617
         Columns(0).Width=   3200
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.TextBox txtNomeFilial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   120
         Width           =   4455
      End
      Begin VB.TextBox txtNomeCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtNomeClasse 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox txtNomeSubClasse 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtNomeProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboProduto 
         Bindings        =   "frmRelVendas2.frx":4E9A8
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1560
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
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmRelVendas2.frx":4E9C2
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1200
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
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmRelVendas2.frx":4E9DD
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   840
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
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCliente 
         Bindings        =   "frmRelVendas2.frx":4E9F5
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   480
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
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelVendas2.frx":4EA0F
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   120
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
         BackColor       =   12648447
         DataFieldToDisplay=   "Filial"
      End
      Begin SSDataWidgets_B.SSDBCombo cboOperacao 
         Bindings        =   "frmRelVendas2.frx":4EA28
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   2640
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
         Columns(0).DataField=   "C�digo"
         Columns(0).FieldLen=   256
         Columns(1).Width=   7355
         Columns(1).Caption=   "Nome"
         Columns(1).Name =   "Nome"
         Columns(1).DataField=   "Nome"
         Columns(1).FieldLen=   256
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   12648447
         DataFieldToDisplay=   "Nome"
      End
      Begin VB.Label Label14 
         Caption         =   "Opera��o"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2670
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Tabela"
         Height          =   255
         Left            =   4680
         TabIndex        =   51
         Top             =   2670
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filial"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Sub-Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1590
         Width           =   975
      End
   End
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   6540
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
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
      RecordSource    =   "SELECT C�digo, Nome FROM Produtos WHERE C�digo <> '0' ORDER BY Nome"
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
      RecordSource    =   "SELECT C�digo, Nome FROM [Sub Classes] ORDER BY Nome"
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Data datClasse 
      Caption         =   "datClasse"
      Connect         =   "Access 2000;"
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
      RecordSource    =   "SELECT C�digo, Nome FROM Classes ORDER BY Nome"
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
      RecordSource    =   "SELECT C�digo, Nome FROM Cli_For ORDER BY Nome"
      Top             =   7200
      Width           =   2295
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
      RecordSource    =   "SELECT Filial, Nome FROM [Par�metros Filial]"
      Top             =   6840
      Width           =   2295
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
      Height          =   1305
      Left            =   -120
      TabIndex        =   30
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "� Caso n�o queira utilizar algum filtro, basta n�o preencher o campo"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   270
         TabIndex        =   48
         Top             =   930
         Width           =   5595
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRelVendas2.frx":4EA42
         ForeColor       =   &H00808080&
         Height          =   765
         Left            =   270
         TabIndex        =   31
         Top             =   210
         Width           =   6675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sa�da"
      Height          =   855
      Left            =   3690
      TabIndex        =   24
      Top             =   5220
      Width           =   3585
      Begin VB.OptionButton optSaidaImpressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1980
         TabIndex        =   26
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optSaidaVideo 
         Appearance      =   0  'Flat
         Caption         =   "V�deo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6120
      Width           =   3555
   End
   Begin VB.Frame Frame1 
      Caption         =   "Per�odo"
      Height          =   885
      Left            =   30
      TabIndex        =   9
      Top             =   4320
      Width           =   3615
      Begin VB.CheckBox chkSepararData 
         Appearance      =   0  'Flat
         Caption         =   "Separar por data"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
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
         TabIndex        =   10
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12648447
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
      Begin VB.Label Label4 
         Caption         =   "at�"
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "De"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ordem"
      Height          =   855
      Left            =   30
      TabIndex        =   13
      Top             =   5220
      Width           =   3615
      Begin VB.OptionButton optRankingValor 
         Appearance      =   0  'Flat
         Caption         =   "Ranking por valor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optRankingUnidade 
         Appearance      =   0  'Flat
         Caption         =   "Ranking por unidade"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optOrdemNome 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optOrdemCodigo 
         Appearance      =   0  'Flat
         Caption         =   "C�digo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRelVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'06/07/2005 - Daniel
'Vari�vel modular com finalidade de monitorar �s personaliza��es para a empresa Zue
Dim m_blnZue As Boolean
'---------------------------------------------------------------------------
'27/07/2006 - Andrea
'Cria��o de vari�vel para auxiliar na personaliza��o para a empresa BeStar
'QS40011-300 - Para filtrar a tabela de pre�os
Dim m_blnBeStar As Boolean
'----------------------------------------------------------------------------


Private Sub cboClasse_CloseUp()
  cboClasse.Text = cboClasse.Columns(0).Text
  cboClasse_LostFocus
End Sub

Private Sub cboClasse_LostFocus()
  Dim rstClasses As Recordset
  
  txtNomeClasse.Text = ""
  If Not IsNumeric(cboClasse.Text) Then Exit Sub
  
  Set rstClasses = db.OpenRecordset("SELECT C�digo, Nome FROM Classes WHERE C�digo = " & cboClasse.Text, dbOpenSnapshot)
  
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
  
  Set rstClientes = db.OpenRecordset("SELECT C�digo, Nome FROM Cli_For WHERE C�digo = " & cboCliente.Text, dbOpenSnapshot)
  
  With rstClientes
    If Not (.BOF And .EOF) Then
      txtNomeCliente.Text = .Fields("Nome") & ""
    End If
    
    If Not rstClientes Is Nothing Then .Close
    Set rstClientes = Nothing
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
  
  Set rstFiliais = db.OpenRecordset("SELECT Filial, Nome FROM [Par�metros Filial] WHERE Filial = " & cboFilial.Text, dbOpenSnapshot)
  
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
  
  datClientes.Recordset.FindFirst "C�digo = " & cboFornecedor.Text
  
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
  
  Set rstProdutos = db.OpenRecordset("SELECT C�digo, Nome FROM Produtos WHERE C�digo = '" & cboProduto.Text & "' AND C�digo <> '0' ", dbOpenSnapshot)
  
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
  
  Set rstSubClasses = db.OpenRecordset("SELECT C�digo, Nome FROM [Sub Classes] WHERE C�digo = " & cboSubClasse.Text, dbOpenSnapshot)
  
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
  
  Set rstFuncionarios = db.OpenRecordset(" SELECT Nome FROM Funcion�rios " & _
                                         " WHERE C�digo = " & cboVendedor.Text, dbOpenSnapshot)
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

'25/11/2005 - mpdea
'Inclu�do tratamento de erro
Private Sub cmdImprimir_Click()
  Dim rstRelVendas              As Recordset
  Dim strSQL                    As String
  
  Dim dblQtdeTotalDev           As Double: dblQtdeTotalDev = 0
  Dim dblValorTotalDev          As Double: dblValorTotalDev = 0
  Dim dblTotalDescontoSubTotal  As Double: dblTotalDescontoSubTotal = 0
  
  '16/10/2007 - Anderson
  'Vari�vel criada para verificar a quantidade de f�rmulas utilizadas no Crystal Reports
  Dim intFormulas As Integer
  
  On Error GoTo ErrHandler
  
  If txtNomeFilial.Text = "" Then
      MsgBox "Escolha uma Filial !", vbInformation, "Quick Store"
      Exit Sub
  End If
  
  If (Not IsDate(mskDataInicio.Text)) And (Not IsDate(mskDataFinal.Text)) Then
    mskDataInicio.Text = Data_Atual
    mskDataFinal.Text = Data_Atual
  End If
  
  If Not IsDate(mskDataInicio.Text) Then
    MsgBox "Data inicial inv�lida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If Not IsDate(mskDataFinal.Text) Then
    MsgBox "Data final inv�lida !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  If CDate(mskDataInicio.Text) > CDate(mskDataFinal.Text) Then
    MsgBox "A data inicial n�o pode ser maior que a data final !", vbCritical, "Quick Store"
    Exit Sub
  End If
  
  dbTemp.Execute "DELETE * FROM tblRelVendas"
  
  '---[ Chamada das fun��es para gera��o da tabela tempor�ria ]---'
  
    '08/06/2007 - Anderson
    'Criada fun��o para evitar que ocorram erros no relat�rio quando o sistema registra vendas de produtos normais e depois altera para produtos com grade.
    Call GeraNormalGrade
  
    If chkTipoNormal.Value = vbChecked Then
      Call StatusMsg("Gerando as informa��es do tipo normal, aguarde . . . ")
      GeraNormal
    End If
    
    If chkTipoGrade.Value = vbChecked Then
      Call StatusMsg("Gerando as informa��es do tipo grade, aguarde . . . ")
      GeraGrade
    End If

'    If chkTipoEdicao.Value = vbChecked Then
'      Call StatusMsg("Gerando as informa��es do tipo edi��o, aguarde . . . ")
'      GeraEdicao
'    End If
    
    Call StatusMsg("")
  '---[ Chamada das fun��es para gera��o da tabela tempor�ria ]---'
  
  Set rstRelVendas = dbTemp.OpenRecordset("SELECT * FROM tblRelVendas", dbOpenSnapshot)
  
  With rstRelVendas
    If Not (.BOF And .EOF) Then
      '---[ Gera o total de Descontos do sub-total ]---'
        Call StatusMsg("Analisando descontos no sub-total e devolu��es, aguarde . . . ")
        ReturnDescontoSubTotal dblTotalDescontoSubTotal
        ReturnDevolucaoNormal dblValorTotalDev, dblQtdeTotalDev
        ReturnDevolucaoGrade dblValorTotalDev, dblQtdeTotalDev
      '---[ Gera o total de Descontos do sub-total ]---'
      
      '--------------------------------------------------------------------------
      '08/07/2005 - Daniel
      'Adicionado Tratamento para o relat�rio Sint�tico desenvolvido para
      'a empresa Zue de Londrina
      '--------------------------------------------------------------------------
      If optSintetico.Value Then Call AgruparRegistros
            
      With crpView
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        
        If optSaidaVideo.Value Then .Destination = crptToWindow
        If optSaidaImpressora.Value Then .Destination = crptToPrinter
        
        '08/07/2005 - Daniel
        'Adicionado Personaliza��es para a Zue
        If Not optSintetico.Value Then
          'Rotina para Clientes comuns
          .SortFields(0) = "-{tblRelVendas.proTipo}"
        
          If optOrdemCodigo.Value Then .SortFields(1) = "+{Produtos.C�digo Ordena��o}"
          If optOrdemNome.Value Then .SortFields(1) = "+{Produtos.Nome}"
          '12/07/2005 - Daniel
          'Ordena��o para Zue
          If m_blnZue Then
            If optRankingUnidade.Value Then .SortFields(1) = "-{tblRelVendas.venQtde}"
            If optRankingValor.Value Then .SortFields(1) = "-{tblRelVendas.venValor}"
          Else
            If optRankingUnidade.Value Then .SortFields(1) = "+{tblRelVendas.venQtde}"
            If optRankingValor.Value Then .SortFields(1) = "+{tblRelVendas.venValor}"
          End If
          
          If chkSepararData.Value = vbChecked Then
            .ReportFileName = gsReportPath & "rptVendasPorData.rpt"
          Else
            .ReportFileName = gsReportPath & "rptVendas.rpt"
          End If
          
          .DataFiles(0) = gsTempDBFileName
          .DataFiles(1) = gsQuickDBFileName
          .DataFiles(2) = gsQuickDBFileName
          .DataFiles(3) = gsTempDBFileName
          .DataFiles(4) = gsQuickDBFileName
          .DataFiles(5) = gsQuickDBFileName
        
        
        Else
          'Rotina para Zue
          '
          'Tratamento de ordena��o. A princ�pio cliente preferiu ordenar
          'por decrescente a "Qtde" ap�s a Filial
          .SortFields(0) = "+{tblRelVendasGroup.Filial}"
          If optRankingUnidade.Value Then .SortFields(1) = "-{tblRelVendasGroup.Qtde}"  'Ordem decrescente Uni
          If optOrdemCodigo.Value Then .SortFields(1) = "+{tblRelVendasGroup.Produto}"
          If optRankingValor.Value Then .SortFields(1) = "-{tblRelVendasGroup.Valor}"   'Ordem decrescente Val
          
          .ReportFileName = gsReportPath & "rptVendasGroup.rpt"
          .DataFiles(0) = gsTempDBFileName
          .DataFiles(1) = gsQuickDBFileName
          .DataFiles(2) = gsQuickDBFileName
          .DataFiles(3) = gsQuickDBFileName
          
        End If
        
        .Formulas(0) = "DescSubTotal = " & Replace(Format(CStr(dblTotalDescontoSubTotal), "###0.00"), gsCurrencyDecimal, ".")
        .Formulas(1) = "DevolucoesQtde = " & Replace(Format(CStr(dblQtdeTotalDev), "###0.00"), gsCurrencyDecimal, ".")
        .Formulas(2) = "DevolucoesValor = " & Replace(Format(CStr(dblValorTotalDev), "###0.00"), gsCurrencyDecimal, ".")
        '---[ Preenchimento dos campos de cabe�alho de filtro ]---'
        .Formulas(3) = "Periodo = '" & "De " & mskDataInicio.Text & " at� " & mskDataFinal.Text & "'"
          
          '16/10/2007 - Anderson
          'Informa a quantidade de f�rmulas utilizadas no Crystal Reports
          intFormulas = 4
          'Retirado condi��es pois o Crystal exibe apenas os parametros se todos forem digitados.
          'If Len(Trim(txtNomeFilial.Text)) > 0 Then .Formulas(4) = "Filtro_Filial = '" & txtNomeFilial.Text & "'"
          'If Len(Trim(txtNomeCliente.Text)) > 0 Then .Formulas(5) = "Filtro_Cliente = '" & txtNomeCliente.Text & "'"
          'If Len(Trim(txtNomeClasse.Text)) > 0 Then .Formulas(6) = "Filtro_Classe = '" & txtNomeClasse.Text & "'"
          'If Len(Trim(txtNomeSubClasse.Text)) > 0 Then .Formulas(7) = "Filtro_SubClasse = '" & txtNomeSubClasse.Text & "'"
          'If Len(Trim(txtNomeProduto.Text)) > 0 Then .Formulas(8) = "Filtro_Produto = '" & txtNomeProduto.Text & "'"
          'If Len(Trim(txtNomeFornecedor.Text)) > 0 Then .Formulas(9) = "Filtro_Fornecedor = '" & txtNomeFornecedor.Text & "'"
          'If Len(Trim(txtNomeVendedor.Text)) > 0 Then .Formulas(10) = "Filtro_Vendedor = '" & txtNomeVendedor.Text & "'"
          ''-------------------------------------------------
          ''06/07/2006 - Andrea
          ''Inclu�do passagem do par�metro=> tabela de pre�os.
          'If Len(Trim(cboTabela.Text)) > 0 Then .Formulas(11) = "Filtro_Tabela = '" & cboTabela.Text & "'"
          ''-------------------------------------------------
          ''16/10/2007 - Anderson
          ''Customiza��o de relat�rio para Agrotama
          'If Len(Trim(cboOperacao.Text)) > 0 Then .Formulas(12) = "Filtro_Operacao = '" & cboOperacao.Text & " - " & txtNomeOperacao & "'"
          If Len(Trim(txtNomeFilial.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Filial = '" & txtNomeFilial.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(txtNomeCliente.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Cliente = '" & txtNomeCliente.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(txtNomeClasse.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Classe = '" & txtNomeClasse.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_SubClasse = '" & txtNomeSubClasse.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(txtNomeProduto.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Produto = '" & txtNomeProduto.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Fornecedor = '" & txtNomeFornecedor.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(txtNomeVendedor.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Vendedor = '" & txtNomeVendedor.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(cboTabela.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Tabela = '" & cboTabela.Text & "'"
            intFormulas = intFormulas + 1
          End If
          If Len(Trim(cboOperacao.Text)) > 0 Then
            .Formulas(intFormulas) = "Filtro_Operacao = '" & cboOperacao.Text & " - " & txtNomeOperacao & "'"
            intFormulas = intFormulas + 1
          End If
        '---[ Preenchimento dos campos de cabe�alho de filtro ]---'
        
        '10/05/2004 - Daniel
        'Tratamento de 05 casas decimais ap�s a ','
        'quando Embalavi
        If g_bln5CasasDecimais Then
            .Formulas(12) = "QtdeCasasDecimais = " & "5"
        '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
        ElseIf g_bln3CasasDecimais Then
            .Formulas(12) = "QtdeCasasDecimais = " & "3"
        Else
            .Formulas(12) = "QtdeCasasDecimais = " & "2"
        End If
        
        
        ' Modelo 1 ou 2
        ''SetPrinterModeloPwd1 crpView
        
        '25/07/2003 - mpdea
        'Seta a impressora para relat�rio
        Call SetPrinterName("REL", crpView)
        
        
        .Action = 1
        pgbProgress.Value = 0
      End With
    Else
      MsgBox "N�o existem informa��es a serem exibidas !", vbInformation, App.Title
    End If
  End With
  
  Call StatusMsg("")
  
  Exit Sub
  
ErrHandler:
  Call StatusMsg("")
  MsgBox "Erro ao imprimir relat�rio: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
  
End Sub

Private Sub GeraNormal()
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstVendas         As Recordset
  Dim rstRelVendas      As Recordset
  Dim rstProdutos       As Recordset
  
  Dim blnProdutoOK      As Boolean
  
  strSQL = " SELECT Sa�das.Filial, Sa�das.Data, [Sa�das - Produtos].[C�digo sem Grade], Sum([Sa�das - Produtos].Qtde) AS SomaQtde, Sum([Sa�das - Produtos].[Pre�o Final]) AS SomaPrecoFinal, [Opera��es Sa�da].Tipo, Sa�das.Tabela, Sa�das.Opera��o  "
  strSQL = strSQL & " FROM ((Sa�das INNER JOIN [Sa�das - Produtos] ON (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia) AND (Sa�das.Filial = [Sa�das - Produtos].Filial)) INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
  
  ''''' comentado em 11/03/2022 strSQL = strSQL & " GROUP BY Sa�das.Filial, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Efetivada, Sa�das.[Nota Cancelada], [Opera��es Sa�da].Tipo, Sa�das.Data, Sa�das.Data, Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Sa�das.Digitador, Sa�das.Tabela, Sa�das.Opera��o "
  ''''' comentado em 11/03/2022 strSQL = strSQL & " Having ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Produtos.Tipo = 'N' "
'''''  strSQL = strSQL & " GROUP BY Sa�das.Filial, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Efetivada, Sa�das.[Nota Cancelada], Sa�das.[Movimenta��o Desfeita], [Opera��es Sa�da].Tipo, Sa�das.Data, Sa�das.Data, Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Sa�das.Digitador, Sa�das.Tabela, Sa�das.Opera��o "
'''''  strSQL = strSQL & " Having ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( NOT Sa�das.[Movimenta��o Desfeita]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Produtos.Tipo = 'N' "

'''''  ALTERA��O inclui filtro [Movimenta��o Desfeita] e remove filtro [Nota Cancelada] 11/03/2022 Pablo
  strSQL = strSQL & " GROUP BY Sa�das.Filial, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Efetivada, Sa�das.[Movimenta��o Desfeita], [Opera��es Sa�da].Tipo, Sa�das.Data, Sa�das.Cliente, Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Sa�das.Digitador, Sa�das.Tabela, Sa�das.Opera��o "
  strSQL = strSQL & " Having ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Movimenta��o Desfeita]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Produtos.Tipo = 'N' "

  strSQL = strSQL & " AND (Sa�das.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Sa�das.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "

  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Cliente = " & cboCliente.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Sa�das - Produtos].[C�digo sem Grade] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND (Sa�das.Digitador = " & cboVendedor.Text & " )"
  End If
  '-------------------------------------------------------
  '06/07/2006 - Andrea
  'Inclu�do filtro tabela de pre�os
  If Len(Trim(cboTabela.Text)) > 0 Then
    strSQL = strSQL & " AND (Sa�das.Tabela = '" & cboTabela.Text & "')"
  End If
  '-------------------------------------------------------
  
  '16/10/2007 - Anderson
  'Implementa��o do filtro opera��o
  'Solicitado por: Agrotama
  If Len(Trim(cboOperacao.Text)) > 0 Then
    strSQL = strSQL & " AND (Sa�das.Opera��o = " & Trim(cboOperacao.Text) & ") "
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
      Set rstProdutos = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE C�digo = '" & .Fields("C�digo Sem Grade") & "' AND Tipo = 'N' ", dbOpenSnapshot)
      
      blnProdutoOK = Not (rstProdutos.BOF And rstProdutos.EOF)
      
      rstProdutos.Close
      Set rstProdutos = Nothing
      
      If blnProdutoOK Then
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("C�digo Sem Grade"))
        End If
      End If
      
      If blnProdutoOK Then
        strSQL = " SELECT * FROM tblRelVendas "
        strSQL = strSQL & " WHERE filID = " & .Fields("Filial")
        strSQL = strSQL & " AND proID = '" & .Fields("C�digo Sem Grade") & "' "
        strSQL = strSQL & " AND proTipo = 'N' "
        strSQL = strSQL & " AND tamID = 0 "
        strSQL = strSQL & " AND corID = 0 "
        If chkSepararData.Value = vbChecked Then strSQL = strSQL & " AND venData = #" & Format(.Fields("Data"), "mm/dd/yyyy") & "# "
        
        Set rstRelVendas = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        If (rstRelVendas.BOF And rstRelVendas.EOF) Then
          rstRelVendas.AddNew
          
          rstRelVendas.Fields("filID") = .Fields("Filial")
          rstRelVendas.Fields("proID") = .Fields("C�digo Sem Grade")
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
          '5 casas ap�s a v�rgula
          If g_bln5CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((.Fields("SomaPrecoFinal")), "##,###,##0.00000")
          '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
          ElseIf g_bln3CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((.Fields("SomaPrecoFinal")), "##,###,##0.000")
          Else 'N�o Embalavi
            rstRelVendas.Fields("venValor") = .Fields("SomaPrecoFinal")
          End If
        Else
          rstRelVendas.Edit
          rstRelVendas.Fields("venQtde") = rstRelVendas.Fields("venQtde") + .Fields("SomaQtde")
          '10/05/2004 - Daniel
          'Caso seja Embalavi, formataremos o valor para
          '5 casas ap�s a v�rgula
          If g_bln5CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((rstRelVendas.Fields("venValor") + .Fields("SomaPrecoFinal")), "##,###,##0.00000")
          '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
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
  
  strSQL = " SELECT Sa�das.Filial, Sa�das.Data, [Sa�das - Produtos].C�digo , [Sa�das - Produtos].[C�digo sem Grade], Sum([Sa�das - Produtos].Qtde) AS SomaQtde, Sum([Sa�das - Produtos].[Pre�o Final]) AS SomaPrecoFinal, [Opera��es Sa�da].Tipo, Sa�das.Tabela, Sa�das.Opera��o "
  strSQL = strSQL & " FROM ((Sa�das INNER JOIN [Sa�das - Produtos] ON (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia) AND (Sa�das.Filial = [Sa�das - Produtos].Filial)) INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
'''  strSQL = strSQL & " GROUP BY Sa�das.Filial, [Sa�das - Produtos].C�digo, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Efetivada, Sa�das.[Nota Cancelada], [Opera��es Sa�da].Tipo, Sa�das.Data, Sa�das.Data, Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Sa�das.Digitador, Sa�das.Tabela, Sa�das.Opera��o "
'''  strSQL = strSQL & " Having ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Produtos.Tipo = 'G' "
  strSQL = strSQL & " GROUP BY Sa�das.Filial, [Sa�das - Produtos].C�digo, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Efetivada, Sa�das.[Nota Cancelada], Sa�das.[Movimenta��o Desfeita], [Opera��es Sa�da].Tipo, Sa�das.Data, Sa�das.Data, Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Sa�das.Digitador, Sa�das.Tabela, Sa�das.Opera��o "
  strSQL = strSQL & " Having ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( NOT Sa�das.[Movimenta��o Desfeita]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Produtos.Tipo = 'G' "
  
  
  strSQL = strSQL & " AND (Sa�das.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Sa�das.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  strSQL = strSQL & " AND [Sa�das - Produtos].C�digo<>[Sa�das - Produtos].[C�digo sem Grade] "

  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Cliente = " & cboCliente.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Sa�das - Produtos].[C�digo sem Grade] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Digitador = " & cboVendedor.Text & " )"
  End If
  '----------------------------------------------------------
  '06/07/2006 - Andrea
  'Inclu�do filtro tabela de pre�os
  If Len(Trim(cboTabela.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Tabela = '" & cboTabela.Text & "' )"
  End If
  '----------------------------------------------------------
  
  '16/10/2007 - Anderson
  'Implementa��o do filtro opera��o
  'Solicitado por: Agrotama
  If Len(Trim(cboOperacao.Text)) > 0 Then
    strSQL = strSQL & " AND (Sa�das.Opera��o = " & Trim(cboOperacao.Text) & ") "
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
      Set rstProdutos = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE C�digo = '" & .Fields("C�digo Sem Grade") & "' AND Tipo = 'G' ", dbOpenSnapshot)
      
      blnProdutoOK = Not (rstProdutos.BOF And rstProdutos.EOF)
      
      rstProdutos.Close
      Set rstProdutos = Nothing
      
      If blnProdutoOK Then
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("C�digo Sem Grade"))
        End If
      End If
      
      '05/05/2005 - Daniel
      'Corre��o a partir da vers�o beta 6.52.0.40
      'BUG: O Relat�rio estava exibindo �s vendas de produtos com
      'grade n�o separando por tamanho e a cor
      intTamanho = Left(Right(.Fields("C�digo"), 6), 3)
      intCor = Right(.Fields("C�digo"), 3)
      '----------------------------------------------------------
      
      If blnProdutoOK Then
        strSQL = " SELECT * FROM tblRelVendas "
        strSQL = strSQL & " WHERE filID = " & .Fields("Filial")
        strSQL = strSQL & " AND proID = '" & .Fields("C�digo Sem Grade") & "' "
        strSQL = strSQL & " AND proTipo = 'G' "
        strSQL = strSQL & " AND tamID = " & intTamanho
        strSQL = strSQL & " AND corID = " & intCor
        If chkSepararData.Value = vbChecked Then strSQL = strSQL & " AND venData = #" & Format(.Fields("Data"), "mm/dd/yyyy") & "# "
        
        Set rstRelVendas = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        If (rstRelVendas.BOF And rstRelVendas.EOF) Then
          rstRelVendas.AddNew
          
          intTamanho = Left(Right(.Fields("C�digo"), 6), 3)
          intCor = Right(.Fields("C�digo"), 3)
          
          rstRelVendas.Fields("filID") = .Fields("Filial")
          rstRelVendas.Fields("proID") = .Fields("C�digo Sem Grade")
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
  '16/10/2007 - Anderson
  'Foi retidado por outro programador, por�m a chamada � feita mais abaixo do c�digo
  'Call CenterForm(Me)
  
  datFiliais.DatabaseName = gsQuickDBFileName
  datClientes.DatabaseName = gsQuickDBFileName
  datClasse.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
  datProdutos.DatabaseName = gsQuickDBFileName
  datVendedores.DatabaseName = gsQuickDBFileName
  '----------------------------------------------------------
  '06/07/2006 - Andrea
  datPrecos.DatabaseName = gsQuickDBFileName
  '----------------------------------------------------------

  '16/10/2007 - Anderson
  'Customiza��o do relat�rio para Agrotama
  datOperacao.DatabaseName = gsQuickDBFileName

  '-----------------------------------------------------------------------------
  '06/07/2005 - Daniel
  'Personaliza��es para a empresa Zue o relat�rio de vendas oferecer� a op��o
  'de exibir o relat�rio de vendas Sint�tico agrupando os valores de produtos
  'que utilizam grade e tamb�m os normais
  m_blnZue = CheckSerialCaseMod("QS71258-374")
  '-----------------------------------------------------------------------------
  '27/07/2006 - Andrea
  'Personaliza��o para a empresa BeStar, o relat�rio de vendas oferecer� a op��o
  'de selecionar a tabela de pre�os
  m_blnBeStar = CheckSerialCaseMod("QS40011-300")
  '-----------------------------------------------------------------------------
  If Not m_blnBeStar Then ' Case demais empresas
    Label13.Visible = False
    cboTabela.Visible = False
    '16/10/2007 - Anderson
    'Customiza��o do relat�rio da Agrotama
    txtNomeOperacao.Width = 4455
  End If
  
  If Not m_blnZue Then 'Case demais empresas
    fraRelatorio.Visible = False
    optAnalitico.Visible = False
    optSintetico.Visible = False
    Frame5.Width = 3375
    chkTipoNormal.Left = 120
    chkTipoNormal.Top = 360
    chkTipoGrade.Left = 1320
    chkTipoGrade.Top = 360
    chkTipoEdicao.Left = 2400
    chkTipoEdicao.Top = 360
  Else 'Case Zue
    fraRelatorio.Visible = True
    optAnalitico.Visible = True
    optSintetico.Visible = True
    Frame5.Width = 1815
    chkTipoNormal.Left = 120
    chkTipoNormal.Top = 240
    chkTipoGrade.Left = 120
    chkTipoGrade.Top = 480
    chkTipoEdicao.Left = 960
    chkTipoEdicao.Top = 240
  End If
  
  Call CenterForm(Me)
  '-----------------------------------------------------------------------------
  
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
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [Entradas - Produtos].C�digo, Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Pre�o Final]) AS PrecoTotal " & _
           " FROM ((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Sequ�ncia = [Entradas - Produtos].Sequ�ncia) AND (Entradas.Filial = [Entradas - Produtos].Filial)) INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo) INNER JOIN Produtos ON [Entradas - Produtos].C�digo = Produtos.C�digo " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Digitador, [Entradas - Produtos].C�digo, Entradas.Fornecedor, [Opera��es Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Opera��es Entrada].Tipo)='D')) "

  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboCliente.Text & ") "
  End If
  
  '15/03/2004 - Daniel
  'N�o estava fazendo o filtro por Vendedor
  'Foi acrescentado tamb�m esta linha no GROUP BY:
  'Entradas.Digitador
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Digitador = " & cboVendedor.Text & ") "
  End If
  
  If Len(Trim(cboProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Entradas - Produtos].C�digo = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(cboClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(cboSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  '15/03/2004 - Daniel
  'N�o estava fazendo o filtro por Vendedor
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
          blnProdutoOK = blnVerificaForncedor(.Fields("C�digo"))
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
  
  strSQL = " SELECT Entradas.Filial, Entradas.Data, [C�digos da Grade].[C�digo Original], Sum([Entradas - Produtos].Qtde) AS ContarDeQtde, Sum([Entradas - Produtos].[Pre�o Final]) AS PrecoTotal " & _
           " FROM (((Entradas INNER JOIN [Entradas - Produtos] ON (Entradas.Filial = [Entradas - Produtos].Filial) AND (Entradas.Sequ�ncia = [Entradas - Produtos].Sequ�ncia)) INNER JOIN [Opera��es Entrada] ON Entradas.Opera��o = [Opera��es Entrada].C�digo) INNER JOIN [C�digos da Grade] ON [Entradas - Produtos].C�digo = [C�digos da Grade].C�digo) INNER JOIN Produtos ON [C�digos da Grade].[C�digo Original] = Produtos.C�digo " & _
           " GROUP BY Entradas.Filial, Entradas.Data, Entradas.Digitador, [C�digos da Grade].[C�digo Original], Entradas.Fornecedor, [Opera��es Entrada].Tipo, Produtos.Classe, Produtos.[Sub Classe] " & _
           " HAVING ((([Opera��es Entrada].Tipo)='D')) "


  strSQL = strSQL & " AND (Entradas.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) " & _
                    " AND (Entradas.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Fornecedor = " & cboCliente.Text & ") "
  End If
  
  '15/03/2004 - Daniel
  'N�o estava fazendo o filtro por Vendedor
  'Foi acrescentado tamb�m esta linha no GROUP BY:
  'Entradas.Digitador
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Entradas.Digitador = " & cboVendedor.Text & ") "
  End If
  
  If Len(Trim(cboProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([C�digos da Grade].[C�digo Original] = '" & cboProduto.Text & "') "
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
          blnProdutoOK = blnVerificaForncedor(.Fields("C�digo Original"))
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
  
  '23/04/2009 - mpdea
  'Modificado forma de calcular o desconto rateado, pois n�o contemplava v�rias situa��es
  'Agora obt�m o valor rateado de acordo com cada item (Desc ST / ST * Pre�o final item)
  '22/02/2007 - Anderson
  'Acrescentado o campo quantidade, pois estava gerando problemas quando ao imprimir o relat�rio o valor total de desconto no sub total estava se multiplicando de acordo com os itens
  'repetidos. Isso era comum em casos onde o leitor de c�digo de barras era utilizado.
  'Para resolver o problema o campo quantidade foi acrescentado para que fosse dividido pelo valor total de desconto.
  'strSQL = " SELECT SUM(Sa�das.DescontoSubTotal) AS DescontoSubTotal, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Filial, Sa�das.Sequ�ncia, Sa�das.Tabela "
  strSQL = "SELECT SUM(Sa�das.DescontoSubTotal) AS DescontoSubTotal, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Filial, Sa�das.Sequ�ncia, Sa�das.Tabela, Sum([Sa�das - Produtos].Qtde) as Qtde "
  '23/04/2009 - mpdea
  strSQL = strSQL & ", SUM(Sa�das.DescontoSubTotal / Sa�das.Produtos * [Sa�das - Produtos].[Pre�o Final]) AS DescontoRateado "
  strSQL = strSQL & "FROM ((Sa�das INNER JOIN [Sa�das - Produtos] ON (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia) AND (Sa�das.Filial = [Sa�das - Produtos].Filial)) INNER JOIN Produtos ON [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo) INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
'''  strSQL = strSQL & "GROUP BY Sa�das.Filial, Sa�das.Data, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Digitador, Produtos.Classe, Produtos.[Sub Classe], Sa�das.Efetivada, Sa�das.[Nota Cancelada], [Opera��es Sa�da].Tipo = 'V', Sa�das.Sequ�ncia, Sa�das.DescontoSubTotal, Sa�das.Tabela, Sa�das.Opera��o "
  strSQL = strSQL & "GROUP BY Sa�das.Filial, Sa�das.Data, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Digitador, Produtos.Classe, Produtos.[Sub Classe], Sa�das.Efetivada, Sa�das.[Nota Cancelada], Sa�das.[Movimenta��o Desfeita], [Opera��es Sa�da].Tipo = 'V', Sa�das.Sequ�ncia, Sa�das.DescontoSubTotal, Sa�das.Tabela, Sa�das.Opera��o "
  '23/04/2009 - mpdea
  strSQL = strSQL & ", Sa�das.DescontoSubTotal / Sa�das.Produtos * [Sa�das - Produtos].[Pre�o Final] "
'''  strSQL = strSQL & "HAVING ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Sa�das.DescontoSubTotal > 0 "
  strSQL = strSQL & "HAVING ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( NOT Sa�das.[Movimenta��o Desfeita]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Sa�das.DescontoSubTotal > 0 "
  strSQL = strSQL & "AND (Sa�das.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & "AND (Sa�das.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & "AND ( Sa�das.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & "AND ( Sa�das.Cliente = " & cboCliente.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & "AND ([Sa�das - Produtos].[C�digo sem Grade] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & "AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & "AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " ) "
  End If
  
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & "AND ( Sa�das.Digitador = " & cboVendedor.Text & " ) "
  End If
  '-------------------------------------------------------------------------------
  '27/07/2006 - Andrea
  'Inclu�do esta linha para que se o usu�rio selecionar uma tabela, seja filtrada
  'na hora de somar os descontos
  If Len(Trim(cboTabela.Text)) > 0 Then
    strSQL = strSQL & "AND (Sa�das.Tabela = '" & cboTabela.Text & "') "
  End If
  '-------------------------------------------------------------------------------
  
  '16/10/2007 - Anderson
  'Implementa��o do filtro opera��o
  'Solicitado por: Agrotama
  If Len(Trim(cboOperacao.Text)) > 0 Then
    strSQL = strSQL & "AND (Sa�das.Opera��o = " & Trim(cboOperacao.Text) & ") "
  End If
  
  dbTemp.Execute "DELETE * FROM tblRelVendasDescontoSubTotal"
  
  Set rstVendas = db.OpenRecordset(strSQL, dbOpenSnapshot)
  With rstVendas
    If Not (.BOF And .EOF) Then
      .MoveLast
      .MoveFirst
      pgbProgress.min = 0
      pgbProgress.Max = .RecordCount + 1
      .MoveFirst
      
      Do While Not .EOF
        strSQL = " SELECT * FROM tblRelVendasDescontoSubTotal WHERE filID = " & .Fields("Filial")
        strSQL = strSQL & " AND movSequencia = " & .Fields("Sequ�ncia")
        
        '23/04/2009 - mpdea
        If CDbl(.Fields("DescontoRateado")) > 0 Then
        'If CDbl(.Fields("DescontoSubTotal")) > 0 Then
          Set rstDescontoSubTotal = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
          
          If (rstDescontoSubTotal.BOF And rstDescontoSubTotal.EOF) Then
            '23/04/2009 - mpdea
            'Obt�m desconto rateado
            '22/02/2007 - Anderson
            'Acrescentado o campo quantidade, pois estava gerando problemas quando ao imprimir o relat�rio o valor total de desconto no sub total estava se multiplicando de acordo com os itens
            'repetidos. Isso era comum em casos onde o leitor de c�digo de barras era utilizado.
            'Para resolver o problema o campo quantidade foi acrescentado para que fosse dividido pelo valor total de desconto.
            'dblDescontoSomar = .Fields("DescontoSubTotal")
            'dblDescontoSomar = .Fields("DescontoSubTotal") / .Fields("Qtde")
            dblDescontoSomar = .Fields("DescontoRateado")
            
            rstDescontoSubTotal.AddNew
          Else
            '23/04/2009 - mpdea
            rstDescontoSubTotal.Edit
            dblDescontoSomar = .Fields("DescontoRateado")
            'dblDescontoSomar = 0
          End If
          rstDescontoSubTotal.Fields("filID") = .Fields("Filial")
          rstDescontoSubTotal.Fields("movSequencia") = .Fields("Sequ�ncia")
          rstDescontoSubTotal.Fields("movValorDesconto") = dblDescontoSomar
          rstDescontoSubTotal.Update
          rstDescontoSubTotal.Close
          Set rstDescontoSubTotal = Nothing
        Else
          dblDescontoSomar = 0
        End If
        
        Set rstProdutos = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE C�digo = '" & .Fields("C�digo Sem Grade") & "' ", dbOpenSnapshot)
        
        blnProdutoOK = Not (rstProdutos.BOF And rstProdutos.EOF)
        
        rstProdutos.Close
        Set rstProdutos = Nothing
        
        If blnProdutoOK Then
          If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
            blnProdutoOK = blnVerificaForncedor(.Fields("C�digo Sem Grade"))
          End If
        End If
        
        If blnProdutoOK Then
          dblDescontoSubTotal = dblDescontoSubTotal + dblDescontoSomar
        End If
        
        pgbProgress.Value = .AbsolutePosition
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rstVendas = Nothing
  
  dblValorDesconto = dblDescontoSubTotal

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

Private Sub optAnalitico_Click()
  '08/07/2005 - Daniel
  'Case: Zue
  chkSepararData.Value = vbUnchecked 'Limpamos o objeto
  chkSepararData.Enabled = True      'Habilitamos
  optOrdemNome.Value = True
  optOrdemNome.Enabled = True
End Sub

Private Sub optSintetico_Click()
  '08/07/2005 - Daniel
  'Case: Zue
  chkSepararData.Value = vbUnchecked 'Limpamos o objeto
  chkSepararData.Enabled = False     'Desabilitamos
  optOrdemNome.Value = False
  optOrdemNome.Enabled = False
  optRankingUnidade.Value = True
End Sub

Private Sub AgruparRegistros()
  '--------------------------------------------------------------------------
  '08/07/2005 - Daniel
  'Adicionado Tratamento para o relat�rio Sint�tico desenvolvido para
  'a empresa Zue de Londrina - Agrupamento dos registros da tblRelVendas
  '--------------------------------------------------------------------------
  Dim rstVendas      As Recordset
  Dim rstVendasGroup As Recordset
  Dim strSQL         As String
  
  On Error GoTo TratarErro
  
  'Delete na tempor�ria..
  dbTemp.Execute "DELETE FROM tblRelVendasGroup"
  'Abrimos a tempor�ria para .addnew
  Set rstVendasGroup = dbTemp.OpenRecordset("tblRelVendasGroup", dbOpenDynaset)
  
  strSQL = "SELECT filID, proID, SUM(venQtde) AS Qtde, SUM(venValor) as Valor"
  strSQL = strSQL & " FROM tblRelVendas GROUP BY filID, proID "
  
  Set rstVendas = dbTemp.OpenRecordset(strSQL, dbOpenSnapshot)
  
  If rstVendas.RecordCount = 0 Then Exit Sub
  
  With rstVendas
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      Do Until .EOF
        rstVendasGroup.AddNew
         rstVendasGroup.Fields("Filial").Value = .Fields("filID").Value
         rstVendasGroup.Fields("Produto").Value = .Fields("proID").Value & ""
         rstVendasGroup.Fields("Qtde").Value = .Fields("Qtde").Value
         rstVendasGroup.Fields("Valor").Value = Format(.Fields("Valor").Value, FORMAT_VALUE)
        rstVendasGroup.Update
      
       .MoveNext
      Loop
      
    End If
    .Close
  End With
  
  Set rstVendas = Nothing
  
  rstVendasGroup.Close
  Set rstVendasGroup = Nothing
  
  Exit Sub
  
TratarErro:
  MsgBox "Erro no Agrupamento de registros." & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"
  Err.Clear
  Exit Sub

End Sub

Private Sub GeraNormalGrade()
  Dim strSQL            As String
  Dim blnInTransaction  As Boolean
  
  Dim rstVendas         As Recordset
  Dim rstRelVendas      As Recordset
  Dim rstProdutos       As Recordset
  
  Dim blnProdutoOK      As Boolean
  
  strSQL = " SELECT Sa�das.Filial, Sa�das.Data, [Sa�das - Produtos].C�digo , [Sa�das - Produtos].[C�digo sem Grade], "
  strSQL = strSQL & " Sum([Sa�das - Produtos].Qtde) AS SomaQtde, Sum([Sa�das - Produtos].[Pre�o Final]) AS SomaPrecoFinal, "
  strSQL = strSQL & " [Opera��es Sa�da].Tipo, Sa�das.Tabela, Sa�das.Opera��o "
  strSQL = strSQL & " FROM ("
          strSQL = strSQL & " ("
          strSQL = strSQL & " Sa�das INNER JOIN [Sa�das - Produtos] ON "
          strSQL = strSQL & " (Sa�das.Sequ�ncia = [Sa�das - Produtos].Sequ�ncia) AND "
          strSQL = strSQL & " (Sa�das.Filial = [Sa�das - Produtos].Filial)"
          strSQL = strSQL & " ) "
          strSQL = strSQL & " INNER JOIN Produtos ON "
          strSQL = strSQL & " [Sa�das - Produtos].[C�digo sem Grade] = Produtos.C�digo"
  strSQL = strSQL & " ) "
  strSQL = strSQL & " INNER JOIN [Opera��es Sa�da] ON Sa�das.Opera��o = [Opera��es Sa�da].C�digo "
'''  strSQL = strSQL & " GROUP BY Sa�das.Filial, [Sa�das - Produtos].C�digo, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Efetivada, Sa�das.[Nota Cancelada], [Opera��es Sa�da].Tipo, Sa�das.Data, Sa�das.Data, Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Sa�das.Digitador, Sa�das.Tabela, Sa�das.Opera��o "
'''  strSQL = strSQL & " Having ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Produtos.Tipo = 'G' "
  strSQL = strSQL & " GROUP BY Sa�das.Filial, [Sa�das - Produtos].C�digo, [Sa�das - Produtos].[C�digo sem Grade], Sa�das.Efetivada, Sa�das.[Nota Cancelada], Sa�das.[Movimenta��o Desfeita], [Opera��es Sa�da].Tipo, Sa�das.Data, Sa�das.Data, Sa�das.Filial, Sa�das.Cliente, [Sa�das - Produtos].[C�digo sem Grade], Produtos.Classe, Produtos.[Sub Classe], Produtos.Tipo, Sa�das.Digitador, Sa�das.Tabela, Sa�das.Opera��o "
  strSQL = strSQL & " Having ( Sa�das.Efetivada ) AND ( NOT Sa�das.[Nota Cancelada]) AND ( NOT Sa�das.[Movimenta��o Desfeita]) AND ( [Opera��es Sa�da].Tipo = 'V' ) AND Produtos.Tipo = 'G' "
  
  strSQL = strSQL & " AND (Sa�das.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Sa�das.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  strSQL = strSQL & " AND [Sa�das - Produtos].C�digo=[Sa�das - Produtos].[C�digo sem Grade] "

  If Len(Trim(txtNomeFilial.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Filial = " & cboFilial.Text & ") "
  End If
  
  If Len(Trim(txtNomeCliente.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Cliente = " & cboCliente.Text & ") "
  End If
  
  If Len(Trim(txtNomeProduto.Text)) > 0 Then
    strSQL = strSQL & " AND ([Sa�das - Produtos].[C�digo sem Grade] = '" & cboProduto.Text & "') "
  End If
  
  If Len(Trim(txtNomeClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.Classe = " & cboClasse.Text & ") "
  End If
  
  If Len(Trim(txtNomeSubClasse.Text)) > 0 Then
    strSQL = strSQL & " AND (Produtos.[Sub Classe] = " & cboSubClasse.Text & " )"
  End If
  
  If Len(Trim(txtNomeVendedor.Text)) > 0 Then
    strSQL = strSQL & " AND ( Sa�das.Digitador = " & cboVendedor.Text & " )"
  End If

  '-------------------------------------------------------
  '06/07/2006 - Andrea
  'Inclu�do filtro tabela de pre�os
  If Len(Trim(cboTabela.Text)) > 0 Then
    strSQL = strSQL & " AND (Sa�das.Tabela = '" & cboTabela.Text & "')"
  End If
  '-------------------------------------------------------
  
  '16/10/2007 - Anderson
  'Implementa��o do filtro opera��o
  'Solicitado por: Agrotama
  If Len(Trim(cboOperacao.Text)) > 0 Then
    strSQL = strSQL & " AND (Sa�das.Opera��o = " & Trim(cboOperacao.Text) & ") "
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
      Set rstProdutos = db.OpenRecordset("SELECT Tipo FROM Produtos WHERE C�digo = '" & .Fields("C�digo Sem Grade") & "' AND Tipo = 'G' ", dbOpenSnapshot)
      
      blnProdutoOK = Not (rstProdutos.BOF And rstProdutos.EOF)
      
      rstProdutos.Close
      Set rstProdutos = Nothing
      
      If blnProdutoOK Then
        If Len(Trim(txtNomeFornecedor.Text)) > 0 Then
          blnProdutoOK = blnVerificaForncedor(.Fields("C�digo Sem Grade"))
        End If
      End If
      
      If blnProdutoOK Then
        strSQL = " SELECT * FROM tblRelVendas "
        strSQL = strSQL & " WHERE filID = " & .Fields("Filial")
        strSQL = strSQL & " AND proID = '" & .Fields("C�digo Sem Grade") & "' "
        strSQL = strSQL & " AND proTipo = 'N' "
        strSQL = strSQL & " AND tamID = 0 "
        strSQL = strSQL & " AND corID = 0 "
        If chkSepararData.Value = vbChecked Then strSQL = strSQL & " AND venData = #" & Format(.Fields("Data"), "mm/dd/yyyy") & "# "
        
        Set rstRelVendas = dbTemp.OpenRecordset(strSQL, dbOpenDynaset)
        If (rstRelVendas.BOF And rstRelVendas.EOF) Then
          rstRelVendas.AddNew
          
          rstRelVendas.Fields("filID") = .Fields("Filial")
          rstRelVendas.Fields("proID") = .Fields("C�digo Sem Grade")
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
          '5 casas ap�s a v�rgula
          If g_bln5CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((.Fields("SomaPrecoFinal")), "##,###,##0.00000")
          '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
          ElseIf g_bln3CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((.Fields("SomaPrecoFinal")), "##,###,##0.000")
          Else 'N�o Embalavi
            rstRelVendas.Fields("venValor") = .Fields("SomaPrecoFinal")
          End If
        Else
          rstRelVendas.Edit
          rstRelVendas.Fields("venQtde") = rstRelVendas.Fields("venQtde") + .Fields("SomaQtde")
          '10/05/2004 - Daniel
          'Caso seja Embalavi, formataremos o valor para
          '5 casas ap�s a v�rgula
          If g_bln5CasasDecimais Then
            rstRelVendas.Fields("venValor") = Format((rstRelVendas.Fields("venValor") + .Fields("SomaPrecoFinal")), "##,###,##0.00000")
          '30/04/2007 - Anderson - Implementa��o de 3 casas decimais
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

'16/10/2007 - Anderson
'Customiza��o de Relat�rio para Agrotama
Private Sub cboOperacao_CloseUp()
  cboOperacao.Text = cboOperacao.Columns(0).Text
  cboOperacao_LostFocus
End Sub

'16/10/2007 - Anderson
'Customiza��o de Relat�rio para Agrotama
Private Sub cboOperacao_LostFocus()
  Dim rstOpSaida As Recordset
  
  txtNomeOperacao.Text = ""
  If Not IsNumeric(cboOperacao.Text) Then Exit Sub
  
  Set rstOpSaida = db.OpenRecordset(" SELECT Nome FROM [Opera��es Sa�da] " & _
                                         " WHERE C�digo = " & cboOperacao.Text, dbOpenSnapshot)
  With rstOpSaida
    If Not (.BOF And .EOF) Then
      txtNomeOperacao.Text = .Fields("Nome") & ""
    End If
    
    .Close
    Set rstOpSaida = Nothing
  End With
End Sub

