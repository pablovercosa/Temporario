VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmImprimePreco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Relatório de Tabela de Preços"
   ClientHeight    =   6945
   ClientLeft      =   2640
   ClientTop       =   690
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1790
   Icon            =   "ImprimePreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6945
   ScaleWidth      =   10860
   Begin VB.CheckBox chk_relatorioHorizontal_prod_em_2linhas 
      Appearance      =   0  'Flat
      Caption         =   "Relatório horizontal com nome do produto em duas linhas"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   51
      Top             =   3540
      Width           =   4500
   End
   Begin VB.TextBox txt_codFornecedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8340
      TabIndex        =   50
      Top             =   5250
      Width           =   1395
   End
   Begin VB.Data datMoedas 
      Caption         =   "datMoedas"
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
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Código, Nome FROM Moedas WHERE Nome <> 'REAL' OR Nome <> 'Real' OR Nome <> 'real' ORDER BY Código"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox Título 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1635
      MaxLength       =   40
      TabIndex        =   24
      Text            =   "Tabela de Preços"
      Top             =   4890
      Width           =   3975
   End
   Begin VB.Data datSubClasse 
      Caption         =   "Data2"
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
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome, Código FROM [Sub Classes] ORDER BY Nome"
      Top             =   6960
      Width           =   1665
   End
   Begin VB.Data datPrecos 
      Caption         =   "Data1"
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
      Left            =   240
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Tabela FROM Preços ORDER BY Tabela"
      Top             =   7440
      Width           =   1665
   End
   Begin VB.CheckBox O_Não_Zero 
      Appearance      =   0  'Flat
      Caption         =   "Não imprimir produtos com preço igual a 0 (zero)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   3240
      Width           =   5160
   End
   Begin VB.CheckBox O_Estoque 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir somente produtos com estoque"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   2880
      Width           =   5160
   End
   Begin VB.CheckBox O_Cabeçalho 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir o cabeçalho"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   2220
      Width           =   5160
   End
   Begin VB.CheckBox O_Desconto 
      Appearance      =   0  'Flat
      Caption         =   "Incluir o desconto no preço do produto"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   28
      Top             =   2580
      Width           =   5160
   End
   Begin VB.CheckBox O_Inativos 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir também produtos inativos"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   6
      Top             =   1500
      Width           =   5160
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Con_Classe"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impressão"
      Height          =   3570
      Left            =   5640
      TabIndex        =   38
      Top             =   1380
      Width           =   5115
      Begin VB.TextBox txtNomeMoeda 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3090
         Width           =   3855
      End
      Begin VB.OptionButton optMoedaEstrangeira 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "6 tabelas de preço - impressão horizontal (paisagem) com moeda estrangeira e real"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   180
         TabIndex        =   21
         Top             =   2205
         Width           =   4815
      End
      Begin VB.OptionButton O_2_S 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "2 tabelas de preços - impressão (retrato) "
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   18
         Top             =   1275
         Width           =   4815
      End
      Begin VB.OptionButton O_6V 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "6 tabelas de preço - impressão vertical (retrato)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   630
         Width           =   4815
      End
      Begin VB.OptionButton O_1_A 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "1 tabela de preço - impressão vertical (retrato) - layout 2"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   180
         TabIndex        =   20
         Top             =   1860
         Width           =   4815
      End
      Begin VB.OptionButton O_1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "1 tabela de preço - impressão vertical (retrato) - layout 1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   1590
         Width           =   4815
      End
      Begin VB.OptionButton O_3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "3 tabelas de preço - impressão vertical (retrato)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   945
         Width           =   4815
      End
      Begin VB.OptionButton O_6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "6 tabelas de preço - impressão horizontal (paisagem)"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   4815
      End
      Begin SSDataWidgets_B.SSDBCombo cboMoedas 
         Bindings        =   "ImprimePreco.frx":4E95A
         Height          =   345
         Left            =   180
         TabIndex        =   22
         Top             =   3090
         Width           =   885
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
         _ExtentX        =   1561
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Enabled         =   0   'False
         DataFieldToDisplay=   "Código"
      End
      Begin VB.Label lblMoedaEstrangeira 
         AutoSize        =   -1  'True
         Caption         =   "Moeda Estrangeira"
         Enabled         =   0   'False
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Top             =   2820
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saída"
      Height          =   625
      Left            =   2700
      TabIndex        =   37
      Top             =   3825
      Width           =   2535
      Begin VB.OptionButton O_Impressora 
         Appearance      =   0  'Flat
         Caption         =   "Impressora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1035
         TabIndex        =   14
         Top             =   240
         Width           =   1245
      End
      Begin VB.OptionButton O_Vídeo 
         Appearance      =   0  'Flat
         Caption         =   "Vídeo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   810
      End
   End
   Begin VB.CommandButton B_Emite 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gerar Relatório"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6360
      Width           =   10665
   End
   Begin VB.CheckBox Classes 
      Appearance      =   0  'Flat
      Caption         =   "Tabela de Preços separada por Classes"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   1860
      Width           =   5160
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem de Impressão"
      Height          =   625
      Left            =   90
      TabIndex        =   35
      Top             =   3825
      Width           =   2535
      Begin VB.OptionButton O_Nome 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1380
         TabIndex        =   12
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton O_Código 
         Appearance      =   0  'Flat
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin Crystal.CrystalReport Rel 
      Left            =   10020
      Top             =   6090
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "ImprimePreco.frx":4E972
      DataSource      =   "datPrecos"
      Height          =   345
      Index           =   0
      Left            =   945
      TabIndex        =   0
      Top             =   105
      Width           =   4305
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      _ExtentX        =   7594
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "ImprimePreco.frx":4E98A
      DataSource      =   "datPrecos"
      Height          =   345
      Index           =   1
      Left            =   6465
      TabIndex        =   1
      Top             =   75
      Width           =   4305
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      _ExtentX        =   7594
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "ImprimePreco.frx":4E9A2
      DataSource      =   "datPrecos"
      Height          =   345
      Index           =   2
      Left            =   945
      TabIndex        =   2
      Top             =   555
      Width           =   4305
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      _ExtentX        =   7594
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "ImprimePreco.frx":4E9BA
      DataSource      =   "datPrecos"
      Height          =   345
      Index           =   3
      Left            =   6465
      TabIndex        =   3
      Top             =   525
      Width           =   4305
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      _ExtentX        =   7594
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "ImprimePreco.frx":4E9D2
      DataSource      =   "datPrecos"
      Height          =   345
      Index           =   4
      Left            =   945
      TabIndex        =   4
      Top             =   1035
      Width           =   4305
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      _ExtentX        =   7594
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Tabela"
   End
   Begin SSDataWidgets_B.SSDBCombo cboLista 
      Bindings        =   "ImprimePreco.frx":4E9EA
      DataSource      =   "datPrecos"
      Height          =   345
      Index           =   5
      Left            =   6465
      TabIndex        =   5
      Top             =   1005
      Width           =   4305
      DataFieldList   =   "Tabela"
      MaxDropDownItems=   16
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
      _ExtentX        =   7594
      _ExtentY        =   609
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   12648447
      DataFieldToDisplay=   "Tabela"
   End
   Begin MSMask.MaskEdBox Data_Altera 
      Height          =   315
      Left            =   4410
      TabIndex        =   23
      ToolTipText     =   "Pressione F2 para Calendário"
      Top             =   4530
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      Enabled         =   0   'False
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
   Begin SSDataWidgets_B.SSDBCombo Combo_Classe 
      Bindings        =   "ImprimePreco.frx":4EA02
      DataSource      =   "Data1"
      Height          =   345
      Left            =   1815
      TabIndex        =   25
      ToolTipText     =   "Use 0 para imprimir todas as classes"
      Top             =   5565
      Width           =   855
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4974
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1640
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
      Enabled         =   0   'False
   End
   Begin SSDataWidgets_B.SSDBCombo cboSubClasse 
      Bindings        =   "ImprimePreco.frx":4EA16
      DataSource      =   "datSubClasse"
      Height          =   345
      Left            =   1815
      TabIndex        =   26
      ToolTipText     =   "Use 0 para imprimir todas as subclasses"
      Top             =   5940
      Width           =   855
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
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4974
      Columns(0).Caption=   "Nome"
      Columns(0).Name =   "Nome"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nome"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1640
      Columns(1).Caption=   "Código"
      Columns(1).Name =   "Código"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Código"
      Columns(1).DataType=   3
      Columns(1).FieldLen=   256
      _ExtentX        =   1508
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   12648447
      Enabled         =   0   'False
   End
   Begin VB.Label Label14 
      Caption         =   "Código do Fornecedor"
      Height          =   195
      Left            =   6450
      TabIndex        =   49
      Top             =   5310
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Classe de produtos"
      Height          =   255
      Left            =   1815
      TabIndex        =   46
      Top             =   5295
      Width           =   1575
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Título para a tabela"
      Height          =   195
      Left            =   90
      TabIndex        =   45
      Top             =   4935
      Width           =   1395
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Subclasse de produtos"
      Height          =   195
      Left            =   90
      TabIndex        =   44
      Top             =   6000
      Width           =   1620
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Classe de produtos"
      Height          =   195
      Left            =   90
      TabIndex        =   43
      Top             =   5625
      Width           =   1380
   End
   Begin VB.Label lblSubClasse 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2700
      TabIndex        =   42
      Top             =   5940
      Width           =   2940
   End
   Begin VB.Label Label9 
      Caption         =   "Subclasse de produtos :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   41
      Top             =   5970
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label10 
      Caption         =   "Título para a tabela :"
      BeginProperty Font 
         Name            =   "WeblySleek UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1710
      TabIndex        =   40
      Top             =   4950
      Width           =   1590
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Somente produtos com preços alterados a partir de =>"
      Enabled         =   0   'False
      Height          =   195
      Left            =   90
      TabIndex        =   39
      Top             =   4590
      Width           =   3975
   End
   Begin VB.Label Nome_Classe 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2700
      TabIndex        =   36
      Top             =   5565
      Width           =   2940
   End
   Begin VB.Label Label7 
      Caption         =   "Tabela 6"
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Tabela 5"
      Height          =   255
      Left            =   90
      TabIndex        =   33
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label5 
      Caption         =   "Tabela 4"
      Height          =   255
      Left            =   5640
      TabIndex        =   32
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label4 
      Caption         =   "Tabela 3"
      Height          =   255
      Left            =   90
      TabIndex        =   31
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Tabela 2"
      Height          =   255
      Left            =   5640
      TabIndex        =   30
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Tabela 1"
      Height          =   255
      Left            =   90
      TabIndex        =   29
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmImprimePreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPreços        As Recordset
Dim rsPreços_Tempo  As Recordset
Dim rsProdutos      As Recordset
Dim rsClasses       As Recordset
Dim rsSubclasses    As Recordset
Dim rsTabelas       As Recordset
Dim rsParametros    As Recordset
Dim rsFornProduto   As Recordset
'03/11/2004 - Daniel
'Case de Moeda Estrangeira
Dim m_blnMoedaEstrangeira As Boolean

Private Sub B_Emite_Click()
  Dim sSql        As String
  Dim Cód         As String
  Dim Fim         As Integer
  Dim Preço       As Double
  Dim Str_Rel     As String
  Dim Str1        As String
  Dim Classe      As Integer
  Dim nSubClasse  As Integer
  Dim Prod_OK     As Integer
  Dim Tem_Tab     As Boolean
  Dim Est         As Double
  Dim Erro        As Integer
  Dim Aux_Str     As String
  Dim nI          As Integer
  
  '16/06/2005 - Daniel
  'Adicionado rotina para tratamento de erros
  On Error GoTo TratarErro
  
  If Classes.Value = 1 Then
    Rel.WindowShowGroupTree = True
  Else
    Rel.WindowShowGroupTree = False
  End If
  
  Tem_Tab = False
  
  For nI = 0 To 5
    If cboLista(nI) <> "" Then Tem_Tab = True
    If cboLista(nI).Enabled = True Then
      If cboLista(nI).Text <> "" Then
        Tem_Tab = True
      End If
    End If
  Next nI
   
  If Tem_Tab = False Then
    DisplayMsg "Escolha ao menos uma tabela."
    Exit Sub
  End If
  
  '03/11/2004 - Daniel
  'Tratamento para impressão de valores
  'em moeda estrangeira
  If optMoedaEstrangeira.Value Then
    If Len(txtNomeMoeda.Text) <= 0 Then
      MsgBox "Selecione uma moeda estrangeira para geração do relatório.", vbExclamation, "Quick Store"
      cboMoedas.SetFocus
      Exit Sub
    End If
  End If
  
  If optMoedaEstrangeira.Value Then m_blnMoedaEstrangeira = True
  '-------------------------------------------------------------
  
  'Apaga o que tiver no arquivo [Preços - Tempo]
  sSql = "Delete * From [Preços - Tempo]"
  db.Execute sSql
  
  Call StatusMsg("Verificando preços...")
  
  Classe = 0
  nSubClasse = 0
  If Classes.Value = 1 Then
    If Nome_Classe.Caption <> "" Then Classe = Val(Combo_Classe.Text)
    If lblSubClasse.Caption <> "" Then nSubClasse = Val(cboSubClasse.Text)
  End If
  
  Cód = ""
  Fim = False
  rsProdutos.Index = "Código"
  rsPreços.Index = "Tabela"
  rsTabelas.Index = "Tabela"
  Do
    rsProdutos.Seek ">", Cód
    If rsProdutos.NoMatch Then Fim = True
    If Fim = False Then
      Prod_OK = True
      If Classe <> 0 And Classe <> rsProdutos("Classe") Then Prod_OK = False
      If nSubClasse <> 0 And nSubClasse <> rsProdutos("Sub Classe") Then Prod_OK = False
      Cód = rsProdutos("Código")
      
      If Cód = "0" Then Prod_OK = False
      
      If Prod_OK = True Then
        If rsProdutos("Não Incluir na Tabela") = True Then Prod_OK = False
      End If
      
      Rem Verifica Estoque
      If Prod_OK = True Then
        If O_Estoque.Value = 1 Then
          If rsProdutos("Tipo") <> "G" Then Est = Acha_Estoque(gnCodFilial, Cód, 0, 0, 0, Erro)
          If rsProdutos("Tipo") = "G" Then Est = Acha_Estoque_Grade(gnCodFilial, Cód, 0, 0, 0, Erro)
          If Erro <> 0 Then Prod_OK = False
          If Erro = 0 And Est <= 0 Then Prod_OK = False
        End If
      End If
      
      
      Rem Verifica data de alteração
      If O_1.Value = True Or O_1_A.Value = True And Prod_OK = True Then
        If IsDate(Data_Altera.Text) Then
          rsPreços.Seek "=", cboLista(0).Text, Cód
          If rsPreços.NoMatch Then Prod_OK = False
          If Not rsPreços.NoMatch Then
            If Not IsDate(rsPreços("Data Alteração")) Then Prod_OK = False
            If IsDate(rsPreços("Data Alteração")) Then
              If CDate(rsPreços("Data Alteração")) < CDate(Data_Altera.Text) Then Prod_OK = False
            End If
          End If
        End If
      End If
      
      ' Tratamento para filtrar por Código do Fornecedor (se foi preenchido na tela)
      If txt_codFornecedor.Text <> "" Then
          If Not (rsFornProduto.EOF And rsFornProduto.BOF) Then
              rsFornProduto.MoveFirst
          End If
          
          Prod_OK = False
          While Not rsFornProduto.EOF
              If rsFornProduto.Fields("Fornecedor") = txt_codFornecedor.Text And rsFornProduto.Fields("Produto") = Cód Then
                  Prod_OK = True
              End If
          
              rsFornProduto.MoveNext
          Wend
      End If
      ' Fim tratamento Código do Fornecedor
      
      If Prod_OK = True Then
        If rsProdutos("Desativado") = True Then
          If O_Inativos.Value = 0 Then Prod_OK = False
        End If
      End If
      
        
      If Prod_OK = True Then
        '03/11/2004 - Daniel
        'Case Moeda Estrangeira
        If m_blnMoedaEstrangeira Then
          Dim dblCotacao          As Double
          Dim blnMoedaSelecionada As Boolean
        
          dblCotacao = 0
          blnMoedaSelecionada = False
        
          Call BuscarUltimaCotacao(Cód, dblCotacao, blnMoedaSelecionada)
        End If
      
        'Criamos os registros somente para os produtos cadastrados com
        'a moeda selecionada
        If m_blnMoedaEstrangeira And blnMoedaSelecionada Then
        
          rsPreços_Tempo.AddNew
          rsPreços_Tempo("Produto") = Cód
          rsPreços_Tempo("Classe") = rsProdutos("Classe")
          rsPreços_Tempo("SubClasse") = rsProdutos("Sub Classe")
        
          For nI = 0 To 5
            rsPreços_Tempo("Preço " & CStr(nI + 1)) = 0
            rsPreços_Tempo("PreçoNacional " & CStr(nI + 1)) = 0
            If cboLista(nI).Text <> "" Then
              rsPreços.Seek "=", cboLista(nI).Text, Cód
              If Not rsPreços.NoMatch Then
                 If IsNull(rsPreços("Preço")) Then
                   Preço = 0
                 Else
                   Preço = rsPreços("Preço")
                 End If
                 'If O_Desconto.Value = 1 Then Preço = Preço - (rsProdutos("Desconto") * Preço / 100)
                 rsPreços_Tempo("Preço " & CStr(nI + 1)) = Preço
                 rsPreços_Tempo("PreçoNacional " & CStr(nI + 1)) = Format(Preço * dblCotacao, FORMAT_VALUE)
                 
                 rsTabelas.Seek "=", rsPreços("Tabela")
                 If Not rsTabelas.NoMatch Then
                   If rsTabelas("Dividir") <> 0 Then
                     rsPreços_Tempo("Preço " & CStr(nI + 1)) = Format((Preço / rsTabelas("Dividir")), "############.00")
                     rsPreços_Tempo("PreçoNacional " & CStr(nI + 1)) = Format(((Format((Preço / rsTabelas("Dividir")), "############.00")) * dblCotacao), FORMAT_VALUE)
                   End If
                 End If
              End If
            End If
          Next nI
        
          rsPreços_Tempo.Update
        
        End If
        
        If Not m_blnMoedaEstrangeira Then
          
          rsPreços_Tempo.AddNew
          rsPreços_Tempo("Produto") = Cód
          rsPreços_Tempo("Classe") = rsProdutos("Classe")
          rsPreços_Tempo("SubClasse") = rsProdutos("Sub Classe")
        
          For nI = 0 To 5
            rsPreços_Tempo("Preço " & CStr(nI + 1)) = 0
            If cboLista(nI).Text <> "" Then
              rsPreços.Seek "=", cboLista(nI).Text, Cód
              If Not rsPreços.NoMatch Then
                 If IsNull(rsPreços("Preço")) Then
                   Preço = 0
                 Else
                   Preço = rsPreços("Preço")
                 End If
                 If O_Desconto.Value = 1 Then Preço = Preço - (rsProdutos("Desconto") * Preço / 100)
                 rsPreços_Tempo("Preço " & CStr(nI + 1)) = Preço
                 
                 rsTabelas.Seek "=", rsPreços("Tabela")
                 If Not rsTabelas.NoMatch Then
                   If rsTabelas("Dividir") <> 0 Then
                     rsPreços_Tempo("Preço " & CStr(nI + 1)) = Format((Preço / rsTabelas("Dividir")), "############.00")
                   End If
                 End If
              End If
            End If
          Next nI
         
          rsPreços_Tempo.Update
         
        End If 'If Not m_blnMoedaEstrangeira Then
        
      End If
   
   End If
  
  Loop Until Fim = True
  
  Call StatusMsg("")
  
  If O_1.Value = True Or O_1_A.Value = True Then
    If O_Não_Zero.Value = 1 Then
     sSql = "Delete * From [Preços - Tempo] Where [Preço 1] = 0"
     db.Execute sSql
    End If
  End If
  
  Rem  Seta Valores e Manda Relatório
  
  Rem  Nome do BD
  Str1 = gsQuickDBFileName
  Rel.DataFiles(0) = Str1
    
  Rem Saída
  If O_Vídeo = True Then Rel.Destination = 0
  If O_Impressora = True Then Rel.Destination = 1
  
  Rem Nome do arquivo .rpt
  If O_6.Value = True Then 'Horizontal
    If Classes.Value = 0 Then
        If chk_relatorioHorizontal_prod_em_2linhas.Value = vbChecked Then
            Str1 = gsReportPath & "Precos1_prod2linhas.rpt"
        Else
            Str1 = gsReportPath & "PRECOS1.RPT"
        End If
    Else
        If chk_relatorioHorizontal_prod_em_2linhas.Value = vbChecked Then
            Str1 = gsReportPath & "Precos2_prod2linhas.rpt"
        Else
            Str1 = gsReportPath & "PRECOS2.RPT"
        End If
    End If
  End If
  
  If O_6V.Value = True Then 'Vertical
    If Classes.Value = 0 Then
      Str1 = gsReportPath & "PRECOS1V.RPT"
    Else
      If MsgBox("Deseja visualizar as Subclasses ?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
        '04/08/2005 - Daniel
        'Adicionado relatório de listas de preços na vertical com 6 tabelas
        'e exibição de classes x subclasses
        Str1 = gsReportPath & "PRECOSCLASSE_SUB.RPT"
      Else
        Str1 = gsReportPath & "PRECOS2V.RPT"
      End If
    End If
  End If
  
  If O_6V.Value Or O_6.Value Or optMoedaEstrangeira.Value Then
   For nI = 0 To 5
    Str_Rel = "TAB" & CStr(nI + 1) & " = '"
    Str_Rel = Str_Rel + cboLista(nI) + "'"
    Rel.Formulas(nI) = Str_Rel
   Next nI
  End If
  
  If O_3.Value = True Then
    If Classes.Value = 0 Then
      Str1 = gsReportPath & "PRECOS1M.RPT"
    Else
      Str1 = gsReportPath & "PRECOS2M.RPT"
    End If
    
    For nI = 0 To 2
      Str_Rel = "TAB" & CStr(nI + 1) & " = '"
      Str_Rel = Str_Rel + cboLista(nI) + "'"
      Rel.Formulas(nI) = Str_Rel
    Next nI
    
    Rel.Formulas(3) = ""
    Rel.Formulas(4) = ""
    Rel.Formulas(5) = ""
  End If
  
  If O_2_S.Value = True Then
    For nI = 0 To 1
        Str_Rel = "TAB" & CStr(nI + 1) & " = '"
        Str_Rel = Str_Rel + cboLista(nI) + "'"
        Rel.Formulas(nI) = Str_Rel
    Next nI
    If Classes.Value = 0 Then
      Str1 = gsReportPath & "PRECOSSIMP.RPT"
    Else
      Str1 = gsReportPath & "PRECOSSIMP1.RPT"
    End If
  End If
  
  If O_1.Value = True Then
    If Classes.Value = 0 Then
      Str1 = gsReportPath & "PRECOS1P.RPT"
    Else
      Str1 = gsReportPath & "PRECOS2P.RPT"
    End If
  
    Str_Rel = "TAB1 = '"
    Str_Rel = Str_Rel + cboLista(0).Text + "'"
    Rel.Formulas(0) = Str_Rel
    Rel.Formulas(1) = ""
    Rel.Formulas(2) = ""
    Rel.Formulas(3) = ""
    Rel.Formulas(4) = ""
    Rel.Formulas(5) = ""
  End If
   
  If O_1_A.Value = True Then
    If Classes.Value = 0 Then
      Str1 = gsReportPath & "PRECOS1A.RPT"
    Else
      Str1 = gsReportPath & "PRECOS2A.RPT"
    End If
    
    Str_Rel = "TAB1 = '"
    Str_Rel = Str_Rel + cboLista(0) + "'"
    Rel.Formulas(0) = Str_Rel
    Rel.Formulas(1) = ""
    Rel.Formulas(2) = ""
    Rel.Formulas(3) = ""
    Rel.Formulas(4) = ""
    Rel.Formulas(5) = ""
  End If
   
  '04/11/2004 - Daniel
  'Adicionado Lista de Preços com valores
  'em moeda estrangeira e real
  If optMoedaEstrangeira.Value Then
    Str1 = gsReportPath & "PRECOSMOEDAESTRANGEIRA.RPT"
  End If
  
  Str_Rel = "Título = '" + Título.Text + "'"
  Rel.Formulas(6) = Str_Rel
  
  Str_Rel = "Cabe1 = '" + rsParametros("Lista 1") & "" + "'"
  Rel.Formulas(7) = Str_Rel
  
  Str_Rel = "Cabe2 = '" + rsParametros("Lista 2") & "" + "'"
  Rel.Formulas(8) = Str_Rel
   
  Str_Rel = "Cabe3 = '" + rsParametros("Lista 3") & "" + "'"
  Rel.Formulas(9) = Str_Rel
   
  Str_Rel = "Cabe4 = '" + rsParametros("Lista 4") & "" + "'"
  Rel.Formulas(10) = Str_Rel
   
  Str_Rel = "Cabe5 = '" + rsParametros("Lista 5") & "" + "'"
  Rel.Formulas(11) = Str_Rel
   
   
  If O_Cabeçalho.Value = 1 Then
    Str_Rel = "Cabeçalho = '1'"
  Else
    Str_Rel = "Cabeçalho = '0'"
  End If
  
  Rel.Formulas(12) = Str_Rel
   
  Rel.ReportFileName = Str1
  
  Str_Rel = "nome_empresa = '"
  Str_Rel = Str_Rel + gsNomeEmpresa + "'"
  Rel.Formulas(13) = Str_Rel
  
  '10/05/2004 - Daniel
  'Caso seja Embalavi, formataremos o valor para
  '5 casas após a vírgula para o preço
  If g_bln5CasasDecimais Then
    Rel.Formulas(14) = "QtdeCasasDecimais = " & "5"
  '30/04/2007 - Anderson - Implementação de 3 casas decimais
  ElseIf g_bln3CasasDecimais Then
    Rel.Formulas(14) = "QtdeCasasDecimais = " & "3"
  Else
    Rel.Formulas(14) = "QtdeCasasDecimais = " & "2"
  End If
  
  If Classes.Value = 1 Then
    If O_Código.Value = True Then
     ' Rel.SortFields(0) = "+{Preços - Tempo.Classe}"
     ' Rel.SortFields(1) = "+{Preços - Tempo.SubClasse} "
      Rel.SortFields(0) = "+{Produtos.Código Ordenação}"
      Rel.SortFields(1) = ""
      Rel.SortFields(2) = ""
    End If
    If O_Nome.Value = True Then
      
      'Rel.SortFields(0) = "+{Preços - Tempo.Classe}"
      'Rel.SortFields(1) = "+{Preços - Tempo.SubClasse} "
      Rel.SortFields(0) = "+{Produtos.Nome}"
      Rel.SortFields(1) = ""
      Rel.SortFields(2) = ""
    End If
  Else
    If O_Código.Value = True Then
      Rel.SortFields(0) = "+{Produtos.Código Ordenação}"
      Rel.SortFields(1) = ""
      Rel.SortFields(2) = ""
    End If
    If O_Nome.Value = True Then
      Rel.SortFields(0) = "+{Produtos.Nome}"
      Rel.SortFields(1) = ""
      Rel.SortFields(2) = ""
    End If
  End If
  
  Call StatusMsg("Aguarde, imprimindo...")
  MousePointer = vbHourglass
  
  '25/07/2003 - mpdea
  'Seta a impressora para relatório
  Call SetPrinterName("REL", Rel)
    
  Rel.Action = 1
  
  Call StatusMsg("")
  MousePointer = vbDefault
  
  Exit Sub

TratarErro:
  Call StatusMsg("")
  MousePointer = vbDefault
  MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbExclamation, "Quick Store"

End Sub

Private Sub cboMoedas_CloseUp()
  cboMoedas.Text = cboMoedas.Columns(0).Text
  cboMoedas_LostFocus
End Sub

Private Sub cboMoedas_LostFocus()
  Dim rstMoedas As Recordset
  
  txtNomeMoeda.Text = ""
  
  If Not IsNumeric(cboMoedas.Text) Then Exit Sub
  
  Set rstMoedas = db.OpenRecordset("SELECT Código, Nome FROM Moedas WHERE Código = " & CByte(cboMoedas.Text), dbOpenDynaset)

  With rstMoedas
    If Not (.BOF And .EOF) Then
      txtNomeMoeda.Text = .Fields("Nome") & ""
    End If
  End With

  rstMoedas.Close
  Set rstMoedas = Nothing

End Sub

Private Sub Classes_Click()
   Combo_Classe.Enabled = -Classes.Value
   Nome_Classe.Enabled = -Classes.Value
   cboSubClasse.Enabled = -Classes.Value
   lblSubClasse.Enabled = -Classes.Value
''
'' If Classes.Value = 1 Then
''   Combo_Classe.Enabled = True
''   Nome_Classe.Enabled = True
''   cboSubClasse.Enabled = True
''   lblSubClasse.Enabled = True
'' Else
''   Combo_Classe.Enabled = False
''   Nome_Classe.Enabled = False
''   cboSubClasse.Enabled = False
''   lblSubClasse.Enabled = False
'' End If
''
End Sub

Private Sub Combo_Classe_CloseUp()
  Combo_Classe.Text = Combo_Classe.Columns(1).Text
  Combo_Classe_LostFocus
End Sub

Private Sub Combo_Classe_LostFocus()
  Dim Aux As Variant
  
  Nome_Classe.Caption = ""
  Aux = Combo_Classe.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsClasses.Index = "Código"
  rsClasses.Seek "=", Val(Aux)
  If rsClasses.NoMatch Then Exit Sub
  
  Nome_Classe.Caption = rsClasses("Nome")

End Sub

Private Sub cboSubClasse_CloseUp()
  cboSubClasse.Text = cboSubClasse.Columns(1).Text
  cboSubClasse_LostFocus
End Sub

Private Sub cboSubClasse_LostFocus()
  Dim Aux As Variant
  
  lblSubClasse.Caption = ""
  Aux = cboSubClasse.Text
  If IsNull(Aux) Then Exit Sub
  If Not IsNumeric(Aux) Then Exit Sub
  If Val(Aux) <= 0 Then Exit Sub
  If Val(Aux) > 9999 Then Exit Sub
  
  rsSubclasses.Index = "Código"
  rsSubclasses.Seek "=", Val(Aux)
  If rsSubclasses.NoMatch Then Exit Sub
  
  lblSubClasse.Caption = rsSubclasses("Nome")

End Sub

Private Sub Data_Altera_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
      SendKeys "+{TAB}{HOME}"
    Case vbKeyDown
      SendKeys "{TAB}{HOME}"
    Case vbKeyF2
      Data_Altera.Text = frmCalendario.gsDateCalender(Data_Altera.Text)
  End Select
End Sub

Private Sub Data_Altera_LostFocus()
  Data_Altera.Text = Ajusta_Data(Data_Altera.Text)
End Sub

Private Sub Form_Load()
  Dim Aux As String
  
  Call CenterForm(Me)
  
  Set rsPreços = db.OpenRecordset("Preços", , dbReadOnly)
  Set rsClasses = db.OpenRecordset("Classes", , dbReadOnly)
  Set rsSubclasses = db.OpenRecordset("Sub Classes", , dbReadOnly)
  Set rsProdutos = db.OpenRecordset("Produtos", , dbReadOnly)
  Set rsPreços_Tempo = db.OpenRecordset("Preços - Tempo")
  Set rsTabelas = db.OpenRecordset("Tabela de Preços", , dbReadOnly)
  Set rsParametros = db.OpenRecordset("Parâmetros Filial", , dbReadOnly)
  Set rsFornProduto = db.OpenRecordset("Forn_Prod", , dbReadOnly)
  
  Data1.DatabaseName = gsQuickDBFileName
  datPrecos.DatabaseName = gsQuickDBFileName
  datSubClasse.DatabaseName = gsQuickDBFileName
  datMoedas.DatabaseName = gsQuickDBFileName
  
  Call GetSettings
  
  rsParametros.Index = "Filial"
  rsParametros.Seek "=", gnCodFilial
  If rsParametros.NoMatch Then Exit Sub
 
End Sub

Private Sub GetSettings()
  cboLista(0).Text = GetSetting("QuickStore", "RelPrecos", "Tab1", "")
  cboLista(1).Text = GetSetting("QuickStore", "RelPrecos", "Tab2", "")
  cboLista(2).Text = GetSetting("QuickStore", "RelPrecos", "Tab3", "")
  cboLista(3).Text = GetSetting("QuickStore", "RelPrecos", "Tab4", "")
  cboLista(4).Text = GetSetting("QuickStore", "RelPrecos", "Tab5", "")
  cboLista(5).Text = GetSetting("QuickStore", "RelPrecos", "Tab6", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SaveSetting("QuickStore", "RelPrecos", "Tab1", cboLista(0).Text)
  Call SaveSetting("QuickStore", "RelPrecos", "Tab2", cboLista(1).Text)
  Call SaveSetting("QuickStore", "RelPrecos", "Tab3", cboLista(2).Text)
  Call SaveSetting("QuickStore", "RelPrecos", "Tab4", cboLista(3).Text)
  Call SaveSetting("QuickStore", "RelPrecos", "Tab5", cboLista(4).Text)
  Call SaveSetting("QuickStore", "RelPrecos", "Tab6", cboLista(5).Text)
  
  rsFornProduto.Close
  Set rsFornProduto = Nothing
End Sub

Private Sub O_1_A_Click()
  cboLista(0).Enabled = True
  cboLista(1).Enabled = False
  cboLista(2).Enabled = False
  cboLista(3).Enabled = False
  cboLista(4).Enabled = False
  cboLista(5).Enabled = False
  Data_Altera.Enabled = True
  Label8.Enabled = True
  O_Não_Zero.Enabled = True
  '03/11/2004 - Daniel
  cboMoedas.Text = ""
  cboMoedas_LostFocus
  lblMoedaEstrangeira.Enabled = False
  cboMoedas.Enabled = False
  Classes.Enabled = True
  O_Cabeçalho.Enabled = True
  O_Desconto.Enabled = True
  m_blnMoedaEstrangeira = False
End Sub

Private Sub O_1_Click()
  cboLista(0).Enabled = True
  cboLista(1).Enabled = False
  cboLista(2).Enabled = False
  cboLista(3).Enabled = False
  cboLista(4).Enabled = False
  cboLista(5).Enabled = False
  Data_Altera.Enabled = True
  Label8.Enabled = True
  O_Não_Zero.Enabled = True
  '03/11/2004 - Daniel
  cboMoedas.Text = ""
  cboMoedas_LostFocus
  lblMoedaEstrangeira.Enabled = False
  cboMoedas.Enabled = False
  Classes.Enabled = True
  O_Cabeçalho.Enabled = True
  O_Desconto.Enabled = True
  m_blnMoedaEstrangeira = False
End Sub

Private Sub O_2_S_Click()
  cboLista(0).Enabled = True
  cboLista(1).Enabled = True
  cboLista(2).Enabled = False
  cboLista(3).Enabled = False
  cboLista(4).Enabled = False
  cboLista(5).Enabled = False
  Data_Altera.Enabled = False
  '03/11/2004 - Daniel
  cboMoedas.Text = ""
  cboMoedas_LostFocus
  lblMoedaEstrangeira.Enabled = False
  cboMoedas.Enabled = False
  Classes.Enabled = True
  O_Cabeçalho.Enabled = True
  O_Desconto.Enabled = True
  m_blnMoedaEstrangeira = False
End Sub

Private Sub O_3_Click()
  cboLista(0).Enabled = True
  cboLista(1).Enabled = True
  cboLista(2).Enabled = True
  cboLista(3).Enabled = False
  cboLista(4).Enabled = False
  cboLista(5).Enabled = False
  Data_Altera.Enabled = False
  Label8.Enabled = False
  O_Não_Zero.Enabled = False
  '03/11/2004 - Daniel
  cboMoedas.Text = ""
  cboMoedas_LostFocus
  lblMoedaEstrangeira.Enabled = False
  cboMoedas.Enabled = False
  Classes.Enabled = True
  O_Cabeçalho.Enabled = True
  O_Desconto.Enabled = True
  m_blnMoedaEstrangeira = False
End Sub

Private Sub O_6_Click()
  cboLista(0).Enabled = True
  cboLista(1).Enabled = True
  cboLista(2).Enabled = True
  cboLista(3).Enabled = True
  cboLista(4).Enabled = True
  cboLista(5).Enabled = True
  Data_Altera.Enabled = True
  Label8.Enabled = False
  O_Não_Zero.Enabled = False
  '03/11/2004 - Daniel
  cboMoedas.Text = ""
  cboMoedas_LostFocus
  lblMoedaEstrangeira.Enabled = False
  cboMoedas.Enabled = False
  Classes.Enabled = True
  O_Cabeçalho.Enabled = True
  O_Desconto.Enabled = True
  m_blnMoedaEstrangeira = False
End Sub

Private Sub O_6V_Click()
  '03/11/2004 - Daniel
  cboMoedas.Text = ""
  cboMoedas_LostFocus
  lblMoedaEstrangeira.Enabled = False
  cboMoedas.Enabled = False
  Classes.Enabled = True
  O_Cabeçalho.Enabled = True
  O_Desconto.Enabled = True
  m_blnMoedaEstrangeira = False
End Sub

Private Sub BuscarUltimaCotacao(ByVal CodProduto As String, ByRef Cotacao As Double, ByRef MoedaSelecionada As Boolean)
  Dim rstCotacoes As Recordset
  Dim rstProdutos As Recordset

  Cotacao = 0

  Set rstCotacoes = db.OpenRecordset("SELECT * FROM Cotações WHERE Moeda = " & CByte(cboMoedas.Text) & " ORDER BY Data ", dbOpenDynaset)

  With rstCotacoes
    If Not (.BOF And .EOF) Then
      .MoveLast
      
      Cotacao = Format((.Fields("Cotação").Value), FORMAT_VALUE)
    End If
    .Close
  End With

  Set rstCotacoes = Nothing

  'Verificação da moeda do Produto
  Set rstProdutos = db.OpenRecordset("SELECT Moeda FROM Produtos WHERE Código = '" & CodProduto & "'", dbOpenDynaset)

  With rstProdutos
    If Not (.BOF And .EOF) Then
      .MoveFirst
    
      If .Fields("Moeda").Value = CByte(cboMoedas.Text) Then MoedaSelecionada = True
    
    End If
    .Close
  End With

  Set rstProdutos = Nothing

End Sub

Private Sub optMoedaEstrangeira_Click()
  lblMoedaEstrangeira.Enabled = True
  cboMoedas.Enabled = True
  
  Classes.Value = vbUnchecked
  Classes.Enabled = False
  O_Cabeçalho.Value = vbUnchecked
  O_Cabeçalho.Enabled = False
  O_Desconto.Value = vbUnchecked
  O_Desconto.Enabled = False
  m_blnMoedaEstrangeira = True
End Sub

