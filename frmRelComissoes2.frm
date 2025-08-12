VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRelComissoes2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Comissões"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelComissoes2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   7335
   Begin VB.Frame Frame8 
      Caption         =   "Nível de detalhamento"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   5040
      Width           =   2295
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optDetalhado 
         Caption         =   "Detalhado"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
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
      Top             =   7920
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
      RecordSource    =   "SELECT Código, Nome FROM Cli_For"
      Top             =   8280
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
      RecordSource    =   "SELECT Código, Nome FROM Classes"
      Top             =   7920
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
      RecordSource    =   "SELECT Código, Nome FROM [Sub Classes]"
      Top             =   8280
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
      RecordSource    =   "SELECT Código, Nome FROM Produtos WHERE Código <> '0'"
      Top             =   7920
      Width           =   2295
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
      RecordSource    =   "SELECT Código, Nome FROM Funcionários"
      Top             =   8280
      Width           =   2295
   End
   Begin VB.Frame Frame7 
      Caption         =   "Exibição do valor"
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   6000
      Width           =   3975
      Begin VB.ComboBox cboQtdeCasasDecimais 
         Height          =   315
         ItemData        =   "frmRelComissoes2.frx":058A
         Left            =   2400
         List            =   "frmRelComissoes2.frx":059D
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Quantidade de casas decimais"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   300
         Width           =   2295
      End
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
      TabIndex        =   13
      Top             =   5040
      Width           =   3135
      Begin VB.OptionButton optOrdemCodigo 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optOrdemNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optRankingUnidade 
         Caption         =   "Ranking por unidade"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optRankingValor 
         Caption         =   "Ranking por valor"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
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
      TabIndex        =   47
      Top             =   4335
      Width           =   3615
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   315
         Left            =   2280
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
      Begin MSMask.MaskEdBox mskDataInicio 
         Height          =   315
         Left            =   480
         TabIndex        =   7
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
         TabIndex        =   49
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "até:"
         Height          =   255
         Left            =   1920
         TabIndex        =   48
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5955
      Width           =   1455
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
      Left            =   5760
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
      Begin VB.OptionButton optSaidaVideo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSaidaImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   23
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
      TabIndex        =   44
      Top             =   -120
      Width           =   9615
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   240
         Picture         =   "frmRelComissoes2.frx":05B0
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório de Comissões de produtos"
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
         TabIndex        =   46
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "• Caso não queira utilizar algum filtro, basta não preenche o campo"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   1200
         Width           =   5175
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
      Height          =   2850
      Left            =   120
      TabIndex        =   28
      Top             =   1440
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2400
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBCombo cboVendedor 
         Bindings        =   "frmRelComissoes2.frx":2418
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   2400
         Width           =   1335
         DataFieldList   =   "Código"
         _Version        =   196617
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFornecedor 
         Bindings        =   "frmRelComissoes2.frx":2434
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
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
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Bindings        =   "frmRelComissoes2.frx":244E
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
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
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
         Bindings        =   "frmRelComissoes2.frx":2468
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
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
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboClasse 
         Bindings        =   "frmRelComissoes2.frx":2483
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   1335
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
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboCliente 
         Bindings        =   "frmRelComissoes2.frx":249B
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1335
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
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Código"
      End
      Begin SSDataWidgets_B.SSDBCombo cboFilial 
         Bindings        =   "frmRelComissoes2.frx":24B5
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataFieldToDisplay=   "Filial"
      End
      Begin VB.Label Label7 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Sub-Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Classe"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filial"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2070
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2430
         Width           =   855
      End
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
      TabIndex        =   9
      Top             =   4335
      Width           =   3375
      Begin VB.CheckBox chkTipoNormal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTipoGrade 
         Caption         =   "Grade"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkTipoEdicao 
         Caption         =   "Edição"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5760
      TabIndex        =   27
      Top             =   6360
      Width           =   1455
   End
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   6840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin Crystal.CrystalReport crpView 
      Left            =   120
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRelComissoes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub cmdImprimir_Click()
  
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
  
  dbTemp.Execute " DELETE * FROM tblRelComissoes ", dbFailOnError
  
  If chkTipoNormal.Value = vbChecked Then
    Call StatusMsg("Gerando as informações sobre comissões dos produtos normais, aguarde . . . ")
    GeraNormal
  End If
  
  Call StatusMsg("")
  
  With crpView
    .Reset
    
    .DataFiles(0) = gsTempDBFileName
    .DataFiles(1) = gsQuickDBFileName
    .DataFiles(2) = gsQuickDBFileName
    .DataFiles(3) = gsQuickDBFileName
    .DataFiles(4) = gsQuickDBFileName
    .DataFiles(5) = gsQuickDBFileName
    .DataFiles(6) = gsQuickDBFileName
    
    If optDetalhado.Value Then .ReportFileName = gsReportPath & "rptComissoesDetalhado.rpt"
    If optNormal.Value Then .ReportFileName = gsReportPath & "rptComissoesNormal.rpt"
    .WindowState = crptMaximized
    
    ' Modelo 1 ou 2
    'SetPrinterModeloPwd2 crpView
    
    If optSaidaImpressora.Value Then .Destination = crptToPrinter
    If optSaidaVideo.Value Then .Destination = crptToWindow
    
    .Action = 1
  End With
End Sub

Private Sub GeraNormal()
  Dim rstComissoes    As Recordset
  Dim rstRelComissoes As Recordset
  Dim strSQL          As String
  Dim blnInTransaction As Boolean
  
  On Error GoTo Erro:
  
  strSQL = " SELECT Comissão.Data, Comissão.Vendedor, Comissão.Produto, Produtos.Nome, Produtos.Tipo, Comissão.Qtde, Comissão.Valor, Comissão.Comissão, Comissão.Sequência, Comissão.Cliente, Comissão.Tabela, Comissão.Filial, Produtos.Classe, Produtos.[Sub Classe] " & _
           " FROM (Comissão INNER JOIN Produtos ON Comissão.Produto = Produtos.Código) INNER JOIN Saídas ON (Comissão.Sequência = Saídas.Sequência) AND (Comissão.Filial = Saídas.Filial) " & _
           " WHERE (Saídas.[Movimentação Desfeita] = FALSE) "
  
  strSQL = strSQL & " AND (Comissão.Data >= #" & Format(mskDataInicio.Text, "mm/dd/yyyy") & "#) "
  strSQL = strSQL & " AND (Comissão.Data <= #" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "#) "
  
  strSQL = strSQL & " ORDER BY Comissão.Contador "
  
  ws.BeginTrans: blnInTransaction = True
  
  Set rstComissoes = db.OpenRecordset(strSQL, dbOpenSnapshot)
  Set rstRelComissoes = dbTemp.OpenRecordset("SELECT * FROM tblRelComissoes", dbOpenDynaset)
  
  With rstComissoes
    If Not (.BOF And .EOF) Then
      .MoveLast
      .MoveFirst
      
      pgbProgress.Max = .RecordCount + 1
      pgbProgress.Value = 0
      
      Do While Not .EOF
        rstRelComissoes.AddNew
        
        rstRelComissoes.Fields("comData") = .Fields("Data")
        rstRelComissoes.Fields("comVendedor") = .Fields("Vendedor")
        rstRelComissoes.Fields("proID") = .Fields("Produto")
        rstRelComissoes.Fields("proNome") = .Fields("Nome")
        rstRelComissoes.Fields("proTipo") = "N"
        rstRelComissoes.Fields("proQtde") = .Fields("Qtde")
        rstRelComissoes.Fields("proValor") = .Fields("Valor")
        rstRelComissoes.Fields("comComissao") = .Fields("Comissão")
        rstRelComissoes.Fields("comSequencia") = .Fields("Sequência")
        rstRelComissoes.Fields("cliID") = .Fields("Cliente")
        rstRelComissoes.Fields("tabNome") = .Fields("Tabela")
        rstRelComissoes.Fields("filID") = .Fields("Filial")
        
        rstRelComissoes.Update
        
        pgbProgress.Value = .AbsolutePosition
        
        .MoveNext
      Loop
    End If
    
    .Close
    Set rstComissoes = Nothing
  End With
  
  ws.CommitTrans: blnInTransaction = False
  
  Exit Sub
  
Erro:
  If blnInTransaction Then ws.Rollback
  Err.Raise Err.Number, Err.Description
End Sub

Private Sub Form_Load()
  Call CenterForm(Me)
End Sub

